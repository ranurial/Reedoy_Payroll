import os
import sqlite3
from datetime import datetime

import psycopg2
from psycopg2.extras import RealDictCursor

SQL_FILE = os.path.join(os.path.dirname(__file__), 'reedoy_payroll_migration.sql')
LEGACY_DB = os.path.join(os.path.dirname(__file__), '_legacy_migration_tmp.sqlite')
MARKER = 'legacy_sqlite_v2'


def pg_dsn():
    url = os.environ.get('DATABASE_URL', '').strip()
    if not url:
        raise RuntimeError('DATABASE_URL is not set in Render environment.')
    return url.replace('postgres://', 'postgresql://', 1)


def table_exists(con, table_name):
    row = con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,),
    ).fetchone()
    return row is not None


def table_columns(con, table_name):
    return {row[1] for row in con.execute(f'PRAGMA table_info({table_name})').fetchall()}


def load_legacy():
    if not os.path.exists(SQL_FILE):
        raise FileNotFoundError(
            f'Migration SQL file not found: {SQL_FILE}. '
            'Keep reedoy_payroll_migration.sql in the repository root.'
        )

    if os.path.exists(LEGACY_DB):
        os.remove(LEGACY_DB)

    con = sqlite3.connect(LEGACY_DB)
    con.row_factory = sqlite3.Row
    try:
        with open(SQL_FILE, 'r', encoding='utf-8') as f:
            sql = f.read()
        con.executescript(sql)
        con.commit()
    except Exception:
        con.close()
        if os.path.exists(LEGACY_DB):
            os.remove(LEGACY_DB)
        raise
    return con


def create_pg_schema(cur):
    cur.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT,
            role TEXT DEFAULT 'Administrator',
            active BOOLEAN DEFAULT TRUE,
            created_at TEXT,
            last_login TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS company_settings (
            id INTEGER PRIMARY KEY,
            company_name TEXT,
            address TEXT,
            phone TEXT,
            email TEXT,
            logo_path TEXT,
            language TEXT DEFAULT 'bn',
            created_at TEXT,
            updated_at TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS workers (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL,
            bangla_name TEXT,
            department TEXT,
            designation TEXT,
            basic_salary NUMERIC DEFAULT 0,
            ot_rate NUMERIC DEFAULT 0,
            refreshment_bill NUMERIC DEFAULT 0,
            created_at TEXT,
            updated_at TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS attendance (
            id INTEGER PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            month_year TEXT NOT NULL,
            present_days NUMERIC DEFAULT 0,
            absent_days NUMERIC DEFAULT 0,
            ot_hours NUMERIC DEFAULT 0,
            advance_deduction NUMERIC DEFAULT 0,
            UNIQUE(worker_id, month_year)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS worker_advances (
            id INTEGER PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            amount NUMERIC DEFAULT 0,
            date TEXT,
            reason TEXT,
            notes TEXT,
            created_at TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY,
            user_id INTEGER,
            action TEXT,
            details TEXT,
            log_time TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS daily_attendance (
            id INTEGER PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            month_year TEXT NOT NULL,
            day INTEGER NOT NULL,
            status TEXT,
            UNIQUE(worker_id, month_year, day)
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS migration_meta (
            key TEXT PRIMARY KEY,
            created_at TEXT NOT NULL
        )
    ''')


def scalar(cur, sql, params=()):
    cur.execute(sql, params)
    row = cur.fetchone()
    return row[0] if row else None


def value(row, key, default=None):
    try:
        return row[key]
    except (KeyError, IndexError):
        return default


def migrate():
    legacy = load_legacy()
    pg = None
    cur = None

    try:
        pg = psycopg2.connect(pg_dsn())
        pg.autocommit = False
        cur = pg.cursor(cursor_factory=RealDictCursor)

        create_pg_schema(cur)

        if scalar(cur, 'SELECT 1 FROM migration_meta WHERE key=%s', (MARKER,)):
            print('Migration already completed. Nothing to do.')
            pg.commit()
            return

        # ------------------------------------------------------------
        # 1. WORKERS
        # ------------------------------------------------------------
        worker_count = 0
        expected_worker_ids = []
        if table_exists(legacy, 'workers'):
            rows = legacy.execute('''
                SELECT id, name, bangla_name, department, designation,
                       basic_salary, ot_rate, refreshment_bill
                FROM workers
                ORDER BY id
            ''').fetchall()

            for r in rows:
                cur.execute('''
                    INSERT INTO workers
                        (id, name, bangla_name, department, designation,
                         basic_salary, ot_rate, refreshment_bill)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        name=EXCLUDED.name,
                        bangla_name=EXCLUDED.bangla_name,
                        department=EXCLUDED.department,
                        designation=EXCLUDED.designation,
                        basic_salary=EXCLUDED.basic_salary,
                        ot_rate=EXCLUDED.ot_rate,
                        refreshment_bill=EXCLUDED.refreshment_bill
                ''', (
                    r['id'], r['name'], value(r, 'bangla_name'),
                    value(r, 'department'), value(r, 'designation'),
                    value(r, 'basic_salary', 0), value(r, 'ot_rate', 0),
                    value(r, 'refreshment_bill', 0),
                ))
                worker_count += 1
                expected_worker_ids.append(r['id'])

        print(f'workers migrated: {worker_count}')

        # ------------------------------------------------------------
        # 2. MONTHLY ATTENDANCE
        # ------------------------------------------------------------
        attendance_count = 0
        if table_exists(legacy, 'attendance'):
            rows = legacy.execute('''
                SELECT id, worker_id, month_year,
                       present_days, absent_days, ot_hours
                FROM attendance
                ORDER BY id
            ''').fetchall()

            for r in rows:
                cur.execute('''
                    INSERT INTO attendance
                        (id, worker_id, month_year,
                         present_days, absent_days, ot_hours)
                    VALUES (%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (worker_id, month_year) DO UPDATE SET
                        id=EXCLUDED.id,
                        present_days=EXCLUDED.present_days,
                        absent_days=EXCLUDED.absent_days,
                        ot_hours=EXCLUDED.ot_hours
                ''', (
                    r['id'], r['worker_id'], r['month_year'],
                    value(r, 'present_days', 0), value(r, 'absent_days', 0),
                    value(r, 'ot_hours', 0),
                ))
                attendance_count += 1

        print(f'attendance migrated: {attendance_count}')

        # ------------------------------------------------------------
        # 3. DAILY ATTENDANCE
        # ------------------------------------------------------------
        # Use the primary-key id for conflict handling. This works even
        # when the live table was created earlier without a composite
        # UNIQUE(worker_id, month_year, day) constraint.
        daily_count = 0
        if table_exists(legacy, 'daily_attendance'):
            rows = legacy.execute('''
                SELECT id, worker_id, month_year, day, status
                FROM daily_attendance
                ORDER BY id
            ''').fetchall()

            for r in rows:
                cur.execute('''
                    INSERT INTO daily_attendance
                        (id, worker_id, month_year, day, status)
                    VALUES (%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        worker_id=EXCLUDED.worker_id,
                        month_year=EXCLUDED.month_year,
                        day=EXCLUDED.day,
                        status=EXCLUDED.status
                ''', (
                    r['id'], r['worker_id'], r['month_year'],
                    r['day'], value(r, 'status'),
                ))
                daily_count += 1

        print(f'daily_attendance migrated: {daily_count}')

        # ------------------------------------------------------------
        # 4. USERS
        # ------------------------------------------------------------
        user_count = 0
        if table_exists(legacy, 'users'):
            user_cols = table_columns(legacy, 'users')
            if 'password_hash' in user_cols:
                password_col = 'password_hash'
            elif 'password' in user_cols:
                password_col = 'password'
            else:
                raise RuntimeError(
                    'Legacy users table has neither password_hash nor password column.'
                )

            # These columns are known in the normal migration package.  The
            # fallback expressions make the script tolerant of older exports.
            def col_or_null(name):
                return name if name in user_cols else 'NULL'

            query = f'''
                SELECT id, username, {password_col} AS password_hash,
                       {col_or_null('full_name')} AS full_name,
                       {col_or_null('role')} AS role,
                       {col_or_null('active')} AS active,
                       {col_or_null('created_at')} AS created_at,
                       {col_or_null('last_login')} AS last_login
                FROM users
                ORDER BY id
            '''
            rows = legacy.execute(query).fetchall()

            for r in rows:
                password_hash = value(r, 'password_hash')
                if password_hash is None or str(password_hash).strip() == '':
                    raise RuntimeError(
                        f'User id {r["id"]} has an empty password hash. '
                        'Migration stopped to avoid creating an unusable login.'
                    )

                cur.execute('''
                    INSERT INTO users
                        (id, username, password_hash, full_name, role,
                         active, created_at, last_login)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        username=EXCLUDED.username,
                        password_hash=EXCLUDED.password_hash,
                        full_name=EXCLUDED.full_name,
                        role=EXCLUDED.role,
                        active=EXCLUDED.active,
                        created_at=EXCLUDED.created_at,
                        last_login=EXCLUDED.last_login
                ''', (
                    r['id'], r['username'], password_hash,
                    value(r, 'full_name'), value(r, 'role', 'Administrator'),
                    True if value(r, 'active') is None else value(r, 'active'),
                    value(r, 'created_at'), value(r, 'last_login'),
                ))
                user_count += 1

        print(f'users migrated: {user_count}')

        # ------------------------------------------------------------
        # 5. WORKER ADVANCES
        # ------------------------------------------------------------
        advance_count = 0
        if table_exists(legacy, 'worker_advances'):
            advance_cols = table_columns(legacy, 'worker_advances')

            def advance_col(name):
                return name if name in advance_cols else 'NULL'

            query = f'''
                SELECT id, worker_id,
                       {advance_col('amount')} AS amount,
                       {advance_col('date')} AS date,
                       {advance_col('reason')} AS reason,
                       {advance_col('notes')} AS notes,
                       {advance_col('created_at')} AS created_at
                FROM worker_advances
                ORDER BY id
            '''
            rows = legacy.execute(query).fetchall()

            for r in rows:
                cur.execute('''
                    INSERT INTO worker_advances
                        (id, worker_id, amount, date, reason, notes, created_at)
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        worker_id=EXCLUDED.worker_id,
                        amount=EXCLUDED.amount,
                        date=EXCLUDED.date,
                        reason=EXCLUDED.reason,
                        notes=EXCLUDED.notes,
                        created_at=EXCLUDED.created_at
                ''', (
                    r['id'], r['worker_id'], value(r, 'amount', 0),
                    value(r, 'date'), value(r, 'reason'), value(r, 'notes'),
                    value(r, 'created_at'),
                ))
                advance_count += 1

        print(f'worker_advances migrated: {advance_count}')

        # ------------------------------------------------------------
        # 6. ACTIVITY LOG
        # ------------------------------------------------------------
        activity_count = 0
        if table_exists(legacy, 'activity_log'):
            activity_cols = table_columns(legacy, 'activity_log')

            def activity_col(name):
                return name if name in activity_cols else 'NULL'

            query = f'''
                SELECT id,
                       {activity_col('user_id')} AS user_id,
                       {activity_col('action')} AS action,
                       {activity_col('details')} AS details,
                       {activity_col('log_time')} AS log_time
                FROM activity_log
                ORDER BY id
            '''
            rows = legacy.execute(query).fetchall()

            for r in rows:
                log_time = value(r, 'log_time')
                if log_time is None:
                    log_time = datetime.now().isoformat(timespec='seconds')

                cur.execute('''
                    INSERT INTO activity_log
                        (id, user_id, action, details, log_time)
                    VALUES (%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        user_id=EXCLUDED.user_id,
                        action=EXCLUDED.action,
                        details=EXCLUDED.details,
                        log_time=EXCLUDED.log_time
                ''', (
                    r['id'], value(r, 'user_id'), value(r, 'action'),
                    value(r, 'details'), log_time,
                ))
                activity_count += 1

        print(f'activity_log migrated: {activity_count}')

        # ------------------------------------------------------------
        # 7. COMPANY SETTINGS
        # ------------------------------------------------------------
        settings_count = 0
        settings_table = None
        if table_exists(legacy, 'company_settings'):
            settings_table = 'company_settings'
        elif table_exists(legacy, 'settings'):
            settings_table = 'settings'

        if settings_table:
            settings_cols = table_columns(legacy, settings_table)

            def setting_col(name):
                return name if name in settings_cols else 'NULL'

            query = f'''
                SELECT id,
                       {setting_col('company_name')} AS company_name,
                       {setting_col('address')} AS address,
                       {setting_col('phone')} AS phone,
                       {setting_col('email')} AS email,
                       {setting_col('logo_path')} AS logo_path,
                       {setting_col('language')} AS language,
                       {setting_col('created_at')} AS created_at,
                       {setting_col('updated_at')} AS updated_at
                FROM {settings_table}
                ORDER BY id
            '''
            rows = legacy.execute(query).fetchall()

            for r in rows:
                cur.execute('''
                    INSERT INTO company_settings
                        (id, company_name, address, phone, email,
                         logo_path, language, created_at, updated_at)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        company_name=EXCLUDED.company_name,
                        address=EXCLUDED.address,
                        phone=EXCLUDED.phone,
                        email=EXCLUDED.email,
                        logo_path=EXCLUDED.logo_path,
                        language=EXCLUDED.language,
                        created_at=EXCLUDED.created_at,
                        updated_at=EXCLUDED.updated_at
                ''', (
                    r['id'], value(r, 'company_name'), value(r, 'address'),
                    value(r, 'phone'), value(r, 'email'), value(r, 'logo_path'),
                    value(r, 'language', 'bn'), value(r, 'created_at'),
                    value(r, 'updated_at'),
                ))
                settings_count += 1

        print(f'company_settings migrated: {settings_count}')

        # ------------------------------------------------------------
        # 8. RESET POSTGRES SERIAL SEQUENCES
        # ------------------------------------------------------------
        for table in (
            'users',
            'company_settings',
            'workers',
            'attendance',
            'worker_advances',
            'activity_log',
            'daily_attendance',
        ):
            cur.execute(
                "SELECT pg_get_serial_sequence(%s, 'id')",
                (f'public.{table}',),
            )
            row = cur.fetchone()
            seq = row['pg_get_serial_sequence'] if row else None
            if seq:
                cur.execute(f'SELECT MAX(id) AS max_id FROM {table}')
                max_row = cur.fetchone()
                max_id = max_row['max_id'] if max_row else None
                if max_id is not None:
                    cur.execute(
                        'SELECT setval(%s::regclass, %s, true)',
                        (seq, max_id),
                    )

        # ------------------------------------------------------------
        # 9. BASIC VALIDATION BEFORE COMMIT
        # ------------------------------------------------------------
        if expected_worker_ids:
            cur.execute('SELECT COUNT(*) AS n, MAX(id) AS max_id FROM workers')
            check = cur.fetchone()
            if check['n'] < worker_count or (check['max_id'] or 0) < max(expected_worker_ids):
                raise RuntimeError('Worker migration validation failed.')

        # Marker is written only after every migration step and validation
        # succeeds. This makes a failed run safe to retry.
        cur.execute(
            '''INSERT INTO migration_meta (key, created_at)
               VALUES (%s, %s)
               ON CONFLICT (key) DO NOTHING''',
            (MARKER, datetime.now().isoformat(timespec='seconds')),
        )

        pg.commit()

        print('')
        print('==============================================')
        print('SUCCESS: legacy migration completed.')
        print(f'Workers:           {worker_count}')
        print(f'Attendance:        {attendance_count}')
        print(f'Daily attendance:  {daily_count}')
        print(f'Users:             {user_count}')
        print(f'Advances:          {advance_count}')
        print(f'Activity log:      {activity_count}')
        print(f'Company settings:  {settings_count}')
        print('==============================================')

    except Exception:
        if pg is not None:
            pg.rollback()
        print('')
        print('MIGRATION FAILED - PostgreSQL transaction rolled back.')
        raise
    finally:
        if cur is not None:
            cur.close()
        if pg is not None:
            pg.close()
        legacy.close()
        if os.path.exists(LEGACY_DB):
            try:
                os.remove(LEGACY_DB)
            except OSError:
                pass


if __name__ == '__main__':
    migrate()
