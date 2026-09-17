import os
import sqlite3
from datetime import datetime

import psycopg
from psycopg.rows import dict_row

SQL_FILE = os.path.join(os.path.dirname(__file__), 'reedoy_payroll_migration.sql')
LEGACY_DB = os.path.join(os.path.dirname(__file__), '_legacy_migration_tmp.sqlite')
MARKER = 'legacy_sqlite_v2'


def pg_dsn():
    url = os.environ.get('DATABASE_URL', '').strip()
    if not url:
        raise RuntimeError('DATABASE_URL is not set in Render environment.')
    return url.replace('postgres://', 'postgresql://', 1)


def load_legacy():
    if not os.path.exists(SQL_FILE):
        raise FileNotFoundError(f'Not found: {SQL_FILE}')
    if os.path.exists(LEGACY_DB):
        os.remove(LEGACY_DB)
    con = sqlite3.connect(LEGACY_DB)
    con.row_factory = sqlite3.Row
    try:
        with open(SQL_FILE, 'r', encoding='utf-8') as f:
            con.executescript(f.read())
        return con
    except Exception:
        con.close()
        raise


def create_pg_schema(cur):
    stmts = [
        """CREATE TABLE IF NOT EXISTS users (
            id SERIAL PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT,
            role TEXT NOT NULL DEFAULT 'User',
            active INTEGER NOT NULL DEFAULT 1,
            created_at TEXT,
            last_login TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS company_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS workers (
            id SERIAL PRIMARY KEY,
            name TEXT NOT NULL,
            bangla_name TEXT,
            department TEXT NOT NULL,
            designation TEXT,
            basic_salary DOUBLE PRECISION NOT NULL DEFAULT 0,
            ot_rate DOUBLE PRECISION NOT NULL DEFAULT 0,
            refreshment_bill DOUBLE PRECISION NOT NULL DEFAULT 0
        )""",
        """CREATE TABLE IF NOT EXISTS attendance (
            id SERIAL PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            month_year TEXT NOT NULL,
            present_days INTEGER DEFAULT 0,
            absent_days INTEGER DEFAULT 0,
            ot_hours DOUBLE PRECISION DEFAULT 0,
            UNIQUE(worker_id, month_year)
        )""",
        """CREATE TABLE IF NOT EXISTS worker_advances (
            id SERIAL PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            month_year TEXT NOT NULL,
            advance_date TEXT NOT NULL,
            amount DOUBLE PRECISION NOT NULL,
            note TEXT
        )""",
        """CREATE TABLE IF NOT EXISTS activity_log (
            id SERIAL PRIMARY KEY,
            username TEXT,
            action TEXT,
            log_time TEXT NOT NULL
        )""",
        """CREATE TABLE IF NOT EXISTS daily_attendance (
            id SERIAL PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            month_year TEXT NOT NULL,
            day INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'P',
            UNIQUE(worker_id, month_year, day)
        )""",
        """CREATE TABLE IF NOT EXISTS migration_meta (
            key TEXT PRIMARY KEY,
            completed_at TEXT NOT NULL
        )""",
    ]
    for sql in stmts:
        cur.execute(sql)


def scalar(cur, sql, args=()):
    cur.execute(sql, args)
    row = cur.fetchone()
    return row[0] if row else None


def migrate():
    legacy = load_legacy()
    try:
        with psycopg.connect(pg_dsn()) as pg:
            with pg.cursor() as cur:
                create_pg_schema(cur)
                if scalar(cur, 'SELECT 1 FROM migration_meta WHERE key=%s', (MARKER,)):
                    print('Migration already completed. Nothing to do.')
                    return

                # 1) Workers — preserve legacy IDs.
                rows = legacy.execute('''
                    SELECT id, name, bangla_name, department, designation,
                           basic_salary, ot_rate, refreshment_bill
                    FROM workers ORDER BY id
                ''').fetchall()
                for r in rows:
                    cur.execute('''
                        INSERT INTO workers
                          (id,name,bangla_name,department,designation,basic_salary,ot_rate,refreshment_bill)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (id) DO UPDATE SET
                          name=EXCLUDED.name,
                          bangla_name=EXCLUDED.bangla_name,
                          department=EXCLUDED.department,
                          designation=EXCLUDED.designation,
                          basic_salary=EXCLUDED.basic_salary,
                          ot_rate=EXCLUDED.ot_rate,
                          refreshment_bill=EXCLUDED.refreshment_bill
                    ''', tuple(r))
                print('workers:', len(rows))

                # 2) Monthly attendance. Legacy has an extra advance_deduction column;
                # current app does not use it, so it is intentionally ignored.
                rows = legacy.execute('''
                    SELECT id, worker_id, month_year, present_days, absent_days, ot_hours
                    FROM attendance ORDER BY id
                ''').fetchall()
                for r in rows:
                    cur.execute('''
                        INSERT INTO attendance
                          (id,worker_id,month_year,present_days,absent_days,ot_hours)
                        VALUES (%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (worker_id,month_year) DO UPDATE SET
                          present_days=EXCLUDED.present_days,
                          absent_days=EXCLUDED.absent_days,
                          ot_hours=EXCLUDED.ot_hours
                    ''', tuple(r))
                print('attendance:', len(rows))

                # 3) Daily attendance.
                rows = legacy.execute('''
                    SELECT id, worker_id, month_year, day, status
                    FROM daily_attendance ORDER BY id
                ''').fetchall()
                for r in rows:
                    cur.execute('''
                        INSERT INTO daily_attendance
                          (id,worker_id,month_year,day,status)
                        VALUES (%s,%s,%s,%s,%s)
                        ON CONFLICT (worker_id,month_year,day) DO UPDATE SET
                          status=EXCLUDED.status
                    ''', tuple(r))
                print('daily_attendance:', len(rows))

                # 4) Users. Legacy columns are:
                # id, username, password, role, password_hash, full_name, active, created_at, last_login
                rows = legacy.execute('''
                    SELECT id, username, password_hash, full_name, role, active, created_at, last_login
                    FROM users ORDER BY id
                ''').fetchall()
                for r in rows:
                    password_hash = r['password_hash'] or r['password'] if 'password' in r.keys() else r['password_hash']
                    cur.execute('''
                        INSERT INTO users
                          (id,username,password_hash,full_name,role,active,created_at,last_login)
                        VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (id) DO UPDATE SET
                          username=EXCLUDED.username,
                          password_hash=EXCLUDED.password_hash,
                          full_name=EXCLUDED.full_name,
                          role=EXCLUDED.role,
                          active=EXCLUDED.active,
                          created_at=EXCLUDED.created_at,
                          last_login=EXCLUDED.last_login
                    ''', (r['id'],r['username'],password_hash,r['full_name'],r['role'],r['active'],r['created_at'],r['last_login']))
                print('users:', len(rows))

                # 5) Current-format worker advances. The legacy dump has 0 rows here.
                rows = legacy.execute('''
                    SELECT id, worker_id, month_year, advance_date, amount, note
                    FROM worker_advances ORDER BY id
                ''').fetchall()
                for r in rows:
                    cur.execute('''
                        INSERT INTO worker_advances
                          (id,worker_id,month_year,advance_date,amount,note)
                        VALUES (%s,%s,%s,%s,%s,%s)
                        ON CONFLICT (id) DO UPDATE SET
                          worker_id=EXCLUDED.worker_id,
                          month_year=EXCLUDED.month_year,
                          advance_date=EXCLUDED.advance_date,
                          amount=EXCLUDED.amount,
                          note=EXCLUDED.note
                    ''', tuple(r))
                print('worker_advances:', len(rows))

                # 6) Activity log — avoid duplicate IDs on reruns.
                rows = legacy.execute('''
                    SELECT id, username, action, log_time
                    FROM activity_log ORDER BY id
                ''').fetchall()
                for r in rows:
                    cur.execute('''
                        INSERT INTO activity_log (id,username,action,log_time)
                        VALUES (%s,%s,%s,%s)
                        ON CONFLICT (id) DO NOTHING
                    ''', tuple(r))
                print('activity_log:', len(rows))

                # 7) Company settings. Prefer the current company_settings table;
                # fill any missing keys from the legacy settings table.
                rows = legacy.execute('SELECT key,value FROM company_settings').fetchall()
                if not rows:
                    rows = legacy.execute('SELECT key,value FROM settings').fetchall()
                for r in rows:
                    cur.execute('''
                        INSERT INTO company_settings(key,value) VALUES(%s,%s)
                        ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value
                    ''', (r['key'], r['value']))
                print('company_settings:', len(rows))

                # 8) Reset PostgreSQL sequences after preserving legacy IDs.
                for table in ['users','workers','attendance','worker_advances','activity_log','daily_attendance']:
                    cur.execute(f"""
                        SELECT setval(
                          pg_get_serial_sequence('{table}','id'),
                          GREATEST(COALESCE((SELECT MAX(id) FROM {table}),0),1),
                          (SELECT COUNT(*) > 0 FROM {table})
                        )
                    """)

                cur.execute('''
                    INSERT INTO migration_meta(key,completed_at)
                    VALUES(%s,%s)
                    ON CONFLICT(key) DO NOTHING
                ''', (MARKER, datetime.now().isoformat(timespec='seconds')))

            pg.commit()
        print('SUCCESS: legacy migration completed.')
    finally:
        legacy.close()
        try:
            os.remove(LEGACY_DB)
        except OSError:
            pass


if __name__ == '__main__':
    migrate()
