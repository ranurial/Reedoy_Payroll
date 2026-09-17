import os
import sqlite3
from datetime import datetime

import psycopg2
from psycopg2.extras import RealDictCursor

BASE = os.path.dirname(os.path.abspath(__file__))
SQL_FILE = os.path.join(BASE, "reedoy_payroll_migration.sql")
LEGACY_DB = os.path.join(BASE, "_legacy_migration_tmp.sqlite")
MARKER = "legacy_sqlite_v3"


def pg_dsn():
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        raise RuntimeError("DATABASE_URL is not set in Render environment.")
    return url.replace("postgres://", "postgresql://", 1)


def table_exists(con, name):
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def columns(con, name):
    return {r[1] for r in con.execute(f'PRAGMA table_info("{name}")').fetchall()}


def val(row, key, default=None):
    try:
        v = row[key]
        return default if v is None and default is not None else v
    except (KeyError, IndexError):
        return default


def load_legacy():
    if not os.path.exists(SQL_FILE):
        raise FileNotFoundError(
            f"Migration SQL file not found: {SQL_FILE}. "
            "Keep reedoy_payroll_migration.sql in the repository root."
        )

    if os.path.exists(LEGACY_DB):
        os.remove(LEGACY_DB)

    con = sqlite3.connect(LEGACY_DB)
    con.row_factory = sqlite3.Row
    try:
        with open(SQL_FILE, "r", encoding="utf-8-sig") as f:
            con.executescript(f.read())
        con.commit()
        return con
    except Exception:
        con.close()
        if os.path.exists(LEGACY_DB):
            os.remove(LEGACY_DB)
        raise


def scalar(cur, sql, params=()):
    cur.execute(sql, params)
    row = cur.fetchone()
    return row[0] if row else None


def migrate():
    legacy = load_legacy()
    pg = None
    cur = None

    try:
        # Import the live application so its own init_db() creates the exact
        # PostgreSQL schema expected by the running Reedoy application.
        import app  # noqa: F401

        pg = psycopg2.connect(pg_dsn(), sslmode="require")
        pg.autocommit = False
        cur = pg.cursor(cursor_factory=RealDictCursor)

        cur.execute("""
            CREATE TABLE IF NOT EXISTS migration_meta (
                key TEXT PRIMARY KEY,
                created_at TEXT NOT NULL
            )
        """)

        if scalar(cur, "SELECT 1 FROM migration_meta WHERE key=%s", (MARKER,)):
            print("Migration already completed. Nothing to do.")
            pg.commit()
            return

        # ------------------------------------------------------------
        # 1. WORKERS
        # ------------------------------------------------------------
        worker_count = 0
        if table_exists(legacy, "workers"):
            rows = legacy.execute("""
                SELECT id, name, basic_salary, ot_rate, department,
                       designation, refreshment_bill, bangla_name
                FROM workers ORDER BY id
            """).fetchall()

            for r in rows:
                cur.execute("""
                    INSERT INTO workers
                        (id, name, basic_salary, ot_rate, department,
                         designation, refreshment_bill, bangla_name)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        name=EXCLUDED.name,
                        basic_salary=EXCLUDED.basic_salary,
                        ot_rate=EXCLUDED.ot_rate,
                        department=EXCLUDED.department,
                        designation=EXCLUDED.designation,
                        refreshment_bill=EXCLUDED.refreshment_bill,
                        bangla_name=EXCLUDED.bangla_name
                """, (
                    r["id"], r["name"], val(r, "basic_salary", 0),
                    val(r, "ot_rate", 0), val(r, "department", "General"),
                    val(r, "designation", "Worker"), val(r, "refreshment_bill", 0),
                    val(r, "bangla_name"),
                ))
                worker_count += 1
        print(f"workers migrated: {worker_count}")

        # ------------------------------------------------------------
        # 2. MONTHLY ATTENDANCE
        # Use the primary id for conflict handling. The live app does NOT
        # define a UNIQUE(worker_id, month_year) constraint.
        # ------------------------------------------------------------
        attendance_count = 0
        if table_exists(legacy, "attendance"):
            rows = legacy.execute("""
                SELECT id, worker_id, month_year, present_days,
                       absent_days, ot_hours, advance_deduction
                FROM attendance ORDER BY id
            """).fetchall()

            for r in rows:
                cur.execute("""
                    INSERT INTO attendance
                        (id, worker_id, month_year, present_days,
                         absent_days, ot_hours, advance_deduction)
                    VALUES (%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        worker_id=EXCLUDED.worker_id,
                        month_year=EXCLUDED.month_year,
                        present_days=EXCLUDED.present_days,
                        absent_days=EXCLUDED.absent_days,
                        ot_hours=EXCLUDED.ot_hours,
                        advance_deduction=EXCLUDED.advance_deduction
                """, (
                    r["id"], r["worker_id"], r["month_year"],
                    val(r, "present_days", 0), val(r, "absent_days", 0),
                    val(r, "ot_hours", 0), val(r, "advance_deduction", 0),
                ))
                attendance_count += 1
        print(f"attendance migrated: {attendance_count}")

        # ------------------------------------------------------------
        # 3. DAILY ATTENDANCE
        # The live app has only id as PRIMARY KEY, so use ON CONFLICT(id).
        # ------------------------------------------------------------
        daily_count = 0
        if table_exists(legacy, "daily_attendance"):
            rows = legacy.execute("""
                SELECT id, worker_id, month_year, day, status
                FROM daily_attendance ORDER BY id
            """).fetchall()

            for r in rows:
                cur.execute("""
                    INSERT INTO daily_attendance
                        (id, worker_id, month_year, day, status)
                    VALUES (%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        worker_id=EXCLUDED.worker_id,
                        month_year=EXCLUDED.month_year,
                        day=EXCLUDED.day,
                        status=EXCLUDED.status
                """, (
                    r["id"], r["worker_id"], r["month_year"],
                    r["day"], val(r, "status", "P"),
                ))
                daily_count += 1
        print(f"daily_attendance migrated: {daily_count}")

        # ------------------------------------------------------------
        # 4. USERS
        # Live app schema: id, username, password_hash, full_name, role,
        # active, created_at, last_login, password.
        # ------------------------------------------------------------
        user_count = 0
        if table_exists(legacy, "users"):
            uc = columns(legacy, "users")
            password_col = "password_hash" if "password_hash" in uc else "password"

            def ucol(name):
                return name if name in uc else "NULL"

            rows = legacy.execute(f"""
                SELECT id, username, {password_col} AS password_hash,
                       {ucol('full_name')} AS full_name,
                       {ucol('role')} AS role,
                       {ucol('active')} AS active,
                       {ucol('created_at')} AS created_at,
                       {ucol('last_login')} AS last_login,
                       {ucol('password')} AS password
                FROM users ORDER BY id
            """).fetchall()

            for r in rows:
                ph = val(r, "password_hash")
                plain = val(r, "password")
                if (ph is None or str(ph).strip() == "") and plain:
                    ph = plain
                if ph is None or str(ph).strip() == "":
                    raise RuntimeError(f"User id {r['id']} has no usable password value.")

                cur.execute("""
                    INSERT INTO users
                        (id, username, password_hash, full_name, role,
                         active, created_at, last_login, password)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        username=EXCLUDED.username,
                        password_hash=EXCLUDED.password_hash,
                        full_name=EXCLUDED.full_name,
                        role=EXCLUDED.role,
                        active=EXCLUDED.active,
                        created_at=EXCLUDED.created_at,
                        last_login=EXCLUDED.last_login,
                        password=EXCLUDED.password
                """, (
                    r["id"], r["username"], ph,
                    val(r, "full_name"), val(r, "role", "Operator"),
                    1 if val(r, "active") is None else val(r, "active"),
                    val(r, "created_at"), val(r, "last_login"), plain,
                ))
                user_count += 1
        print(f"users migrated: {user_count}")

        # ------------------------------------------------------------
        # 5. WORKER ADVANCES
        # Legacy migration SQL contains worker_advances with the same
        # current-format fields used by the web application.
        # ------------------------------------------------------------
        advance_count = 0
        if table_exists(legacy, "worker_advances"):
            ac = columns(legacy, "worker_advances")
            rows = legacy.execute("""
                SELECT id, worker_id, month_year, advance_date, amount, note
                FROM worker_advances ORDER BY id
            """).fetchall()

            for r in rows:
                cur.execute("""
                    INSERT INTO worker_advances
                        (id, worker_id, month_year, advance_date, amount, note)
                    VALUES (%s,%s,%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        worker_id=EXCLUDED.worker_id,
                        month_year=EXCLUDED.month_year,
                        advance_date=EXCLUDED.advance_date,
                        amount=EXCLUDED.amount,
                        note=EXCLUDED.note
                """, (
                    r["id"], r["worker_id"], r["month_year"],
                    r["advance_date"], val(r, "amount", 0), val(r, "note", ""),
                ))
                advance_count += 1
        print(f"worker_advances migrated: {advance_count}")

        # ------------------------------------------------------------
        # 6. ACTIVITY LOG
        # IMPORTANT: live app schema is (id, username, action, log_time),
        # NOT (id, user_id, details, log_time).
        # ------------------------------------------------------------
        activity_count = 0
        if table_exists(legacy, "activity_log"):
            cols_ = columns(legacy, "activity_log")

            def acol(name):
                return name if name in cols_ else "NULL"

            rows = legacy.execute(f"""
                SELECT id,
                       {acol('username')} AS username,
                       {acol('action')} AS action,
                       {acol('log_time')} AS log_time
                FROM activity_log ORDER BY id
            """).fetchall()

            for r in rows:
                log_time = val(r, "log_time") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                cur.execute("""
                    INSERT INTO activity_log
                        (id, username, action, log_time)
                    VALUES (%s,%s,%s,%s)
                    ON CONFLICT (id) DO UPDATE SET
                        username=EXCLUDED.username,
                        action=EXCLUDED.action,
                        log_time=EXCLUDED.log_time
                """, (
                    r["id"], val(r, "username", ""), val(r, "action", ""), log_time
                ))
                activity_count += 1
        print(f"activity_log migrated: {activity_count}")

        # ------------------------------------------------------------
        # 7. COMPANY SETTINGS
        # Live app schema is key/value. Legacy package has company_settings
        # and settings; both use key/value.
        # ------------------------------------------------------------
        settings_count = 0
        source_settings = None
        if table_exists(legacy, "company_settings"):
            source_settings = "company_settings"
        elif table_exists(legacy, "settings"):
            source_settings = "settings"

        if source_settings:
            rows = legacy.execute(
                f'SELECT key, value FROM "{source_settings}" ORDER BY key'
            ).fetchall()
            for r in rows:
                cur.execute("""
                    INSERT INTO company_settings (key, value)
                    VALUES (%s,%s)
                    ON CONFLICT (key) DO UPDATE SET value=EXCLUDED.value
                """, (r["key"], r["value"]))
                settings_count += 1
        print(f"company_settings migrated: {settings_count}")

        # ------------------------------------------------------------
        # 8. RESET SERIAL SEQUENCES
        # ------------------------------------------------------------
        for table in (
            "users", "workers", "attendance", "daily_attendance",
            "worker_advances", "activity_log"
        ):
            cur.execute("SELECT pg_get_serial_sequence(%s, 'id')", (f"public.{table}",))
            row = cur.fetchone()
            seq = row["pg_get_serial_sequence"] if row else None
            if seq:
                cur.execute(f'SELECT MAX(id) AS max_id FROM "{table}"')
                m = cur.fetchone()["max_id"]
                if m is not None:
                    cur.execute("SELECT setval(%s::regclass, %s, true)", (seq, m))

        # ------------------------------------------------------------
        # 9. VALIDATION
        # ------------------------------------------------------------
        checks = [
            ("workers", worker_count),
            ("attendance", attendance_count),
            ("daily_attendance", daily_count),
            ("users", user_count),
            ("worker_advances", advance_count),
            ("activity_log", activity_count),
        ]
        for table, expected in checks:
            cur.execute(f'SELECT COUNT(*) AS n FROM "{table}"')
            actual = cur.fetchone()["n"]
            if actual < expected:
                raise RuntimeError(
                    f"Validation failed for {table}: expected at least {expected}, found {actual}."
                )

        cur.execute("""
            INSERT INTO migration_meta(key, created_at)
            VALUES(%s,%s)
            ON CONFLICT(key) DO NOTHING
        """, (MARKER, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

        pg.commit()

        print("")
        print("==============================================")
        print("SUCCESS: legacy migration completed.")
        print(f"Workers:           {worker_count}")
        print(f"Attendance:        {attendance_count}")
        print(f"Daily attendance:  {daily_count}")
        print(f"Users:             {user_count}")
        print(f"Advances:          {advance_count}")
        print(f"Activity log:      {activity_count}")
        print(f"Company settings:  {settings_count}")
        print("==============================================")

    except Exception:
        if pg is not None:
            pg.rollback()
        print("")
        print("MIGRATION FAILED - PostgreSQL transaction rolled back.")
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


if __name__ == "__main__":
    migrate()
