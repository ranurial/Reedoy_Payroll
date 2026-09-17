import os
import sqlite3
from datetime import datetime

import psycopg2
from psycopg2 import sql
from psycopg2.extras import RealDictCursor

BASE = os.path.dirname(os.path.abspath(__file__))
SQL_FILE = os.path.join(BASE, "reedoy_payroll_migration.sql")
LEGACY_DB = os.path.join(BASE, "_legacy_migration_tmp.sqlite")
MARKER = "legacy_sqlite_v4"


def pg_dsn():
    url = os.environ.get("DATABASE_URL", "").strip()
    if not url:
        raise RuntimeError("DATABASE_URL is not set in Render environment.")
    return url.replace("postgres://", "postgresql://", 1)


def sqlite_table_exists(con, name):
    return con.execute(
        "SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)
    ).fetchone() is not None


def sqlite_columns(con, name):
    return {r[1] for r in con.execute(f'PRAGMA table_info("{name}")').fetchall()}


def pg_columns(cur, table):
    cur.execute("""
        SELECT column_name
        FROM information_schema.columns
        WHERE table_schema='public' AND table_name=%s
    """, (table,))
    return {r["column_name"] for r in cur.fetchall()}


def val(row, key, default=None):
    try:
        v = row[key]
        if v is None and default is not None:
            return default
        return v
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


def scalar(cur, query, params=()):
    cur.execute(query, params)
    row = cur.fetchone()
    return row[0] if row else None


def insert_upsert(cur, table, values, conflict="id"):
    """Insert using only columns that actually exist in the live PostgreSQL table."""
    live = pg_columns(cur, table)
    values = {k: v for k, v in values.items() if k in live}

    if conflict not in values or not values:
        raise RuntimeError(f"Cannot migrate {table}: required conflict column '{conflict}' is missing.")

    cols = list(values.keys())
    if conflict not in live:
        raise RuntimeError(f"Cannot migrate {table}: live table has no '{conflict}' column.")

    col_sql = sql.SQL(", ").join(sql.Identifier(c) for c in cols)
    placeholders = sql.SQL(", ").join(sql.Placeholder() for _ in cols)
    update_cols = [c for c in cols if c != conflict]

    if update_cols:
        update_sql = sql.SQL(", ").join(
            sql.SQL("{c}=EXCLUDED.{c}").format(c=sql.Identifier(c))
            for c in update_cols
        )
        query = sql.SQL("""
            INSERT INTO {table} ({cols})
            VALUES ({vals})
            ON CONFLICT ({conflict}) DO UPDATE SET {updates}
        """).format(
            table=sql.Identifier(table),
            cols=col_sql,
            vals=placeholders,
            conflict=sql.Identifier(conflict),
            updates=update_sql,
        )
    else:
        query = sql.SQL("""
            INSERT INTO {table} ({cols})
            VALUES ({vals})
            ON CONFLICT ({conflict}) DO NOTHING
        """).format(
            table=sql.Identifier(table),
            cols=col_sql,
            vals=placeholders,
            conflict=sql.Identifier(conflict),
        )

    cur.execute(query, [values[c] for c in cols])


def migrate():
    legacy = load_legacy()
    pg = None
    cur = None

    try:
        # Let the real application create the exact schema it expects.
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

        # If v4 already completed, do not repeat the migration.
        if scalar(cur, "SELECT 1 FROM migration_meta WHERE key=%s", (MARKER,)):
            print("Migration already completed. Nothing to do.")
            pg.commit()
            return

        # Show the important live schema difference that caused the previous failure.
        print("Live attendance columns:", sorted(pg_columns(cur, "attendance")))
        print("Live activity_log columns:", sorted(pg_columns(cur, "activity_log")))

        # ------------------------------------------------------------
        # 1. WORKERS
        # ------------------------------------------------------------
        worker_count = 0
        if sqlite_table_exists(legacy, "workers"):
            rows = legacy.execute("""
                SELECT id, name, basic_salary, ot_rate, department,
                       designation, refreshment_bill, bangla_name
                FROM workers ORDER BY id
            """).fetchall()
            for r in rows:
                insert_upsert(cur, "workers", {
                    "id": r["id"],
                    "name": r["name"],
                    "basic_salary": val(r, "basic_salary", 0),
                    "ot_rate": val(r, "ot_rate", 0),
                    "department": val(r, "department", "General"),
                    "designation": val(r, "designation", "Worker"),
                    "refreshment_bill": val(r, "refreshment_bill", 0),
                    "bangla_name": val(r, "bangla_name"),
                })
                worker_count += 1
        print(f"workers migrated: {worker_count}")

        # ------------------------------------------------------------
        # 2. MONTHLY ATTENDANCE
        # IMPORTANT: do NOT assume advance_deduction exists.
        # The live app currently does not have that column.
        # ------------------------------------------------------------
        attendance_count = 0
        if sqlite_table_exists(legacy, "attendance"):
            legacy_cols = sqlite_columns(legacy, "attendance")
            select_cols = [
                c for c in ["id", "worker_id", "month_year", "present_days",
                            "absent_days", "ot_hours", "advance_deduction"]
                if c in legacy_cols
            ]
            rows = legacy.execute(
                "SELECT " + ", ".join(select_cols) + " FROM attendance ORDER BY id"
            ).fetchall()
            for r in rows:
                insert_upsert(cur, "attendance", {
                    "id": val(r, "id"),
                    "worker_id": val(r, "worker_id"),
                    "month_year": val(r, "month_year"),
                    "present_days": val(r, "present_days", 0),
                    "absent_days": val(r, "absent_days", 0),
                    "ot_hours": val(r, "ot_hours", 0),
                    # Included only if the live DB has it.
                    "advance_deduction": val(r, "advance_deduction", 0),
                })
                attendance_count += 1
        print(f"attendance migrated: {attendance_count}")

        # ------------------------------------------------------------
        # 3. DAILY ATTENDANCE
        # ------------------------------------------------------------
        daily_count = 0
        if sqlite_table_exists(legacy, "daily_attendance"):
            legacy_cols = sqlite_columns(legacy, "daily_attendance")
            select_cols = [c for c in ["id", "worker_id", "month_year", "day", "status"] if c in legacy_cols]
            rows = legacy.execute(
                "SELECT " + ", ".join(select_cols) + " FROM daily_attendance ORDER BY id"
            ).fetchall()
            for r in rows:
                insert_upsert(cur, "daily_attendance", {
                    "id": val(r, "id"),
                    "worker_id": val(r, "worker_id"),
                    "month_year": val(r, "month_year"),
                    "day": val(r, "day"),
                    "status": val(r, "status", "P"),
                })
                daily_count += 1
        print(f"daily_attendance migrated: {daily_count}")

        # ------------------------------------------------------------
        # 4. USERS
        # ------------------------------------------------------------
        user_count = 0
        if sqlite_table_exists(legacy, "users"):
            lc = sqlite_columns(legacy, "users")
            password_col = "password_hash" if "password_hash" in lc else "password"
            rows = legacy.execute(f"""
                SELECT id, username,
                       {password_col} AS password_hash,
                       {('full_name' if 'full_name' in lc else 'NULL')} AS full_name,
                       {('role' if 'role' in lc else 'NULL')} AS role,
                       {('active' if 'active' in lc else 'NULL')} AS active,
                       {('created_at' if 'created_at' in lc else 'NULL')} AS created_at,
                       {('last_login' if 'last_login' in lc else 'NULL')} AS last_login,
                       {('password' if 'password' in lc else 'NULL')} AS password
                FROM users ORDER BY id
            """).fetchall()

            for r in rows:
                ph = val(r, "password_hash")
                plain = val(r, "password")
                if (ph is None or str(ph).strip() == "") and plain:
                    ph = plain
                if ph is None or str(ph).strip() == "":
                    raise RuntimeError(f"User id {r['id']} has no usable password value.")

                insert_upsert(cur, "users", {
                    "id": r["id"],
                    "username": r["username"],
                    "password_hash": ph,
                    "full_name": val(r, "full_name"),
                    "role": val(r, "role", "Operator"),
                    "active": 1 if val(r, "active") is None else val(r, "active"),
                    "created_at": val(r, "created_at"),
                    "last_login": val(r, "last_login"),
                    "password": plain,
                })
                user_count += 1
        print(f"users migrated: {user_count}")

        # ------------------------------------------------------------
        # 5. WORKER ADVANCES
        # ------------------------------------------------------------
        advance_count = 0
        if sqlite_table_exists(legacy, "worker_advances"):
            rows = legacy.execute("""
                SELECT id, worker_id, month_year, advance_date, amount, note
                FROM worker_advances ORDER BY id
            """).fetchall()
            for r in rows:
                insert_upsert(cur, "worker_advances", {
                    "id": r["id"],
                    "worker_id": r["worker_id"],
                    "month_year": r["month_year"],
                    "advance_date": r["advance_date"],
                    "amount": val(r, "amount", 0),
                    "note": val(r, "note", ""),
                })
                advance_count += 1
        print(f"worker_advances migrated: {advance_count}")

        # ------------------------------------------------------------
        # 6. ACTIVITY LOG
        # Live schema: id, username, action, log_time.
        # ------------------------------------------------------------
        activity_count = 0
        if sqlite_table_exists(legacy, "activity_log"):
            lc = sqlite_columns(legacy, "activity_log")
            def acol(name):
                return name if name in lc else "NULL"

            rows = legacy.execute(f"""
                SELECT id,
                       {acol('username')} AS username,
                       {acol('action')} AS action,
                       {acol('log_time')} AS log_time
                FROM activity_log ORDER BY id
            """).fetchall()
            for r in rows:
                log_time = val(r, "log_time") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                insert_upsert(cur, "activity_log", {
                    "id": r["id"],
                    "username": val(r, "username", ""),
                    "action": val(r, "action", ""),
                    "log_time": log_time,
                })
                activity_count += 1
        print(f"activity_log migrated: {activity_count}")

        # ------------------------------------------------------------
        # 7. COMPANY SETTINGS
        # ------------------------------------------------------------
        settings_count = 0
        source_settings = None
        if sqlite_table_exists(legacy, "company_settings"):
            source_settings = "company_settings"
        elif sqlite_table_exists(legacy, "settings"):
            source_settings = "settings"

        if source_settings:
            rows = legacy.execute(
                f'SELECT key, value FROM "{source_settings}" ORDER BY key'
            ).fetchall()
            for r in rows:
                insert_upsert(cur, "company_settings", {
                    "key": r["key"],
                    "value": r["value"],
                }, conflict="key")
                settings_count += 1
        print(f"company_settings migrated: {settings_count}")

        # ------------------------------------------------------------
        # 8. RESET SERIAL SEQUENCES
        # ------------------------------------------------------------
        for table in (
            "users", "workers", "attendance", "daily_attendance",
            "worker_advances", "activity_log"
        ):
            if "id" not in pg_columns(cur, table):
                continue
            cur.execute("SELECT pg_get_serial_sequence(%s, 'id') AS seq", (f"public.{table}",))
            row = cur.fetchone()
            seq = row["seq"] if row else None
            if seq:
                cur.execute(sql.SQL("SELECT MAX(id) AS max_id FROM {}" ).format(sql.Identifier(table)))
                m = cur.fetchone()["max_id"]
                if m is not None:
                    cur.execute("SELECT setval(%s::regclass, %s, true)", (seq, m))

        # ------------------------------------------------------------
        # 9. VALIDATE
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
            cur.execute(sql.SQL("SELECT COUNT(*) AS n FROM {}" ).format(sql.Identifier(table)))
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
