import os
import csv
import io
import sqlite3
import hashlib
import calendar
import datetime
from functools import wraps

from flask import (
    Flask, render_template, request, redirect, url_for,
    session, flash, send_file, abort, jsonify
)


# =========================================================
# OPTIONAL EXPORT LIBRARIES
# =========================================================

try:
    import openpyxl
    from openpyxl import Workbook
except Exception:
    openpyxl = None
    Workbook = None


try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle
    )
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
except Exception:
    SimpleDocTemplate = None
    Paragraph = None
    Spacer = None
    Table = None
    TableStyle = None
    A4 = None
    landscape = None


# =========================================================
# FLASK APP
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "CHANGE-ME-IN-RENDER"
)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

DB_PATH = os.environ.get(
    "SQLITE_DB",
    "reedoy_payroll.db"
)

DEFAULT_COMPANY_NAME = (
    "REEDOY TEXTILE DYEING PRINTING & FINISHING"
)


# =========================================================
# MONTHS / LANGUAGE
# =========================================================

MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December"
]


DEPT_BN = {
    "All Departments": "সব বিভাগ",
    "Cutting": "কাটিং",
    "Sewing": "সেলাই",
    "Finishing": "ফিনিশিং",
    "Quality Control": "কোয়ালিটি কন্ট্রোল",
    "Knitting": "নিটিং",
    "Dyeing": "ডাইং",
    "Printing": "প্রিন্টিং",
    "Washing": "ওয়াশিং",
    "Iron & Packing": "আয়রন ও প্যাকিং",
    "Maintenance": "মেইনটেন্যান্স",
    "Store & Logistics": "স্টোর ও লজিস্টিকস",
    "Human Resources & Admin": "মানবসম্পদ ও প্রশাসন",
    "Accounts & Commercial": "অ্যাকাউন্টস ও কমার্শিয়াল",
    "Utility": "ইউটিলিটি"
}


LANG = {
    "Dashboard": "ড্যাশবোর্ড",
    "Workers Management": "কর্মী ব্যবস্থাপনা",
    "Attendance & Calendar": "উপস্থিতি ও ক্যালেন্ডার",
    "Single Payslip": "একক পে-স্লিপ",
    "Advance Salary": "অগ্রিম বেতন",
    "Department Salary Sheet": "বিভাগভিত্তিক বেতন শীট",
    "Settings": "সেটিংস",
    "Worker Name": "কর্মীর নাম",
    "Department": "বিভাগ",
    "Designation": "পদবি",
    "Basic Salary": "মূল বেতন",
    "OT Rate": "OT হার",
    "Nasta Rate": "নাস্তা হার",
    "Present": "উপস্থিত",
    "Absent": "অনুপস্থিত",
    "Absent Deduction": "অনুপস্থিতির কর্তন",
    "OT Amt": "OT টাকা",
    "Nasta": "নাস্তা",
    "Gross Salary": "মোট বেতন",
    "Advance": "অগ্রিম",
    "Net Payable": "নেট প্রদেয়",
    "Payroll Month": "বেতন মাস",
    "Advance Date": "অগ্রিমের তারিখ",
    "Amount (BDT)": "পরিমাণ (টাকা)",
    "Note": "নোট",
    "Save": "সংরক্ষণ",
    "Update": "আপডেট",
    "Delete": "মুছুন",
    "Search": "অনুসন্ধান",
    "Refresh": "রিফ্রেশ",
    "Generate": "তৈরি করুন",
    "Export Excel": "এক্সেল রপ্তানি",
    "Export PDF": "PDF রপ্তানি"
}


# =========================================================
# DATABASE CONNECTION
# =========================================================

def is_postgres():
    return bool(
        DATABASE_URL
        and not DATABASE_URL.startswith("sqlite://")
    )


_PG_POOL = None


def _get_pg_pool():
    global _PG_POOL

    if _PG_POOL is None:
        from psycopg2.pool import ThreadedConnectionPool

        minconn = int(
            os.environ.get("PG_POOL_MIN", "1")
        )

        maxconn = int(
            os.environ.get("PG_POOL_MAX", "4")
        )

        _PG_POOL = ThreadedConnectionPool(
            minconn,
            maxconn,
            DATABASE_URL,
            sslmode="require",
            connect_timeout=10,
            keepalives=1,
            keepalives_idle=30,
            keepalives_interval=10,
            keepalives_count=3
        )

    return _PG_POOL


def db_connect():
    if is_postgres():
        conn = _get_pg_pool().getconn()
        return conn

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def db_release(conn):
    if conn is None:
        return

    if is_postgres():
        try:
            try:
                if getattr(conn, "status", None) != 1:
                    conn.rollback()
            except Exception:
                pass

            _get_pg_pool().putconn(conn)

        except Exception:
            try:
                conn.close()
            except Exception:
                pass

    else:
        try:
            conn.close()
        except Exception:
            pass


def placeholders(sql):
    if is_postgres():
        return sql.replace("?", "%s")

    return sql


def execute(
    sql,
    params=(),
    fetch=False,
    many=False,
    commit=False
):
    conn = db_connect()
    cur = conn.cursor()

    try:
        sql2 = placeholders(sql)

        if many:
            cur.executemany(sql2, params)
        else:
            cur.execute(sql2, params)

        rows = cur.fetchall() if fetch else None

        if commit:
            conn.commit()

        return rows

    finally:
        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)


def row_dict(cur, row):
    if row is None:
        return None

    if hasattr(row, "keys"):
        return dict(row)

    return {
        d[0]: row[i]
        for i, d in enumerate(cur.description)
    }


def fetch_all(sql, params=()):
    conn = db_connect()
    cur = conn.cursor()

    try:
        cur.execute(
            placeholders(sql),
            params
        )

        rows = cur.fetchall()

        return [
            row_dict(cur, r)
            for r in rows
        ]

    finally:
        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)


def fetch_one(sql, params=()):
    rows = fetch_all(sql, params)

    return rows[0] if rows else None


def scalar(sql, params=(), default=0):
    r = fetch_one(sql, params)

    if not r:
        return default

    return next(iter(r.values()))


# =========================================================
# DATABASE HELPERS
# =========================================================

def table_exists(name):
    if is_postgres():
        return bool(
            scalar(
                """
                SELECT EXISTS (
                    SELECT 1
                    FROM information_schema.tables
                    WHERE table_schema='public'
                    AND table_name=?
                )
                """,
                (name,),
                False
            )
        )

    return bool(
        scalar(
            """
            SELECT COUNT(*)
            FROM sqlite_master
            WHERE type='table'
            AND name=?
            """,
            (name,),
            0
        )
    )


def columns(name):
    if not table_exists(name):
        return set()

    if is_postgres():
        return {
            r["column_name"]
            for r in fetch_all(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema='public'
                AND table_name=?
                """,
                (name,)
            )
        }

    conn = db_connect()
    cur = conn.cursor()

    try:
        cur.execute(
            "PRAGMA table_info(" + name + ")"
        )

        return {
            r[1]
            for r in cur.fetchall()
        }

    finally:
        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)


def add_column(name, col, typ):
    if col in columns(name):
        return

    execute(
        f"ALTER TABLE {name} ADD COLUMN {col} {typ}",
        commit=True
    )


def next_id(table_name):
    """
    PostgreSQL migrated tables may have INTEGER PRIMARY KEY
    without automatic sequence. Therefore explicitly generate
    the next ID when necessary.
    """

    try:
        value = scalar(
            f"SELECT COALESCE(MAX(id),0) FROM {table_name}",
            default=0
        )

        return int(value or 0) + 1

    except Exception:
        return 1


def hash_password(password):
    return hashlib.sha256(
        str(password).encode("utf-8")
    ).hexdigest()


def nowstr():
    return datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

def init_db():
    """
    IMPORTANT:
    This function only creates missing structures.
    It NEVER drops, truncates or replaces migrated data.
    """

    conn = db_connect()
    cur = conn.cursor()

    try:

        cur.execute(
            placeholders(
                """
                CREATE TABLE IF NOT EXISTS workers (
                    id INTEGER PRIMARY KEY,
                    name TEXT NOT NULL,
                    basic_salary REAL NOT NULL DEFAULT 0,
                    ot_rate REAL NOT NULL DEFAULT 0,
                    department TEXT,
                    designation TEXT,
                    refreshment_bill REAL NOT NULL DEFAULT 0,
                    bangla_name TEXT
                )
                """
            )
        )

        cur.execute(
            placeholders(
                """
                CREATE TABLE IF NOT EXISTS attendance (
                    id INTEGER PRIMARY KEY,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    present_days INTEGER DEFAULT 0,
                    absent_days INTEGER DEFAULT 0,
                    ot_hours REAL DEFAULT 0,
                    advance_deduction REAL DEFAULT 0
                )
                """
            )
        )

        cur.execute(
            placeholders(
                """
                CREATE TABLE IF NOT EXISTS daily_attendance (
                    id INTEGER PRIMARY KEY,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    day INTEGER NOT NULL,
                    status TEXT NOT NULL DEFAULT 'P'
                )
                """
            )
        )

        cur.execute(
            placeholders(
                """
                CREATE TABLE IF NOT EXISTS company_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
                """
            )
        )

        cur.execute(
            placeholders(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT,
                    full_name TEXT,
                    role TEXT DEFAULT 'Operator',
                    active INTEGER DEFAULT 1,
                    created_at TEXT,
                    last_login TEXT,
                    password TEXT
                )
                """
            )
        )

        cur.execute(
            placeholders(
                """
                CREATE TABLE IF NOT EXISTS activity_log (
                    id INTEGER PRIMARY KEY,
                    username TEXT,
                    action TEXT,
                    log_time TEXT NOT NULL
                )
                """
            )
        )

        conn.commit()

    finally:
        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)

    # Optional worker columns
    for col, typ in [
        ("phone", "TEXT"),
        ("address", "TEXT"),
        ("joining_date", "TEXT"),
        ("status", "TEXT")
    ]:
        try:
            add_column("workers", col, typ)
        except Exception:
            pass

    # User compatibility columns
    try:
        add_column("users", "password_hash", "TEXT")
    except Exception:
        pass

    try:
        add_column("users", "password", "TEXT")
    except Exception:
        pass

    # Default admin
    try:
        existing_admin = fetch_one(
            """
            SELECT id
            FROM users
            WHERE lower(username)=?
            LIMIT 1
            """,
            ("admin",)
        )

        if not existing_admin:

            cols = columns("users")

            names = []
            vals = []

            default_values = [
                ("username", "admin"),
                (
                    "password_hash",
                    hash_password("admin123")
                ),
                (
                    "password",
                    hash_password("admin123")
                ),
                (
                    "full_name",
                    "System Administrator"
                ),
                (
                    "role",
                    "Administrator"
                ),
                ("active", 1),
                ("created_at", nowstr())
            ]

            for name, value in default_values:
                if name in cols:
                    names.append(name)
                    vals.append(value)

            if names:
                new_id = next_id("users")

                if "id" in cols:
                    names.insert(0, "id")
                    vals.insert(0, new_id)

                execute(
                    "INSERT INTO users("
                    + ",".join(names)
                    + ") VALUES("
                    + ",".join(["?"] * len(vals))
                    + ")",
                    vals,
                    commit=True
                )

    except Exception as e:
        app.logger.warning(
            "Admin initialization error: %r",
            e
        )

    # Default company settings
    defaults = {
        "company_name": DEFAULT_COMPANY_NAME,
        "company_address": "",
        "company_phone": "",
        "company_email": "",
        "company_logo": ""
    }

    try:

        for key, value in defaults.items():

            if is_postgres():

                execute(
                    """
                    INSERT INTO company_settings(key,value)
                    VALUES(?,?)
                    ON CONFLICT(key)
                    DO NOTHING
                    """,
                    (key, value),
                    commit=True
                )

            else:

                execute(
                    """
                    INSERT OR IGNORE
                    INTO company_settings(key,value)
                    VALUES(?,?)
                    """,
                    (key, value),
                    commit=True
                )

    except Exception as e:
        app.logger.warning(
            "Settings initialization error: %r",
            e
        )


# =========================================================
# SETTINGS / ACTIVITY / USER
# =========================================================

def get_settings():
    try:
        return {
            r["key"]: r["value"]
            for r in fetch_all(
                "SELECT key,value FROM company_settings"
            )
        }

    except Exception:
        return {
            "company_name": DEFAULT_COMPANY_NAME
        }


def log_activity(username, action):
    try:

        new_id = next_id("activity_log")

        execute(
            """
            INSERT INTO activity_log
            (id,username,action,log_time)
            VALUES(?,?,?,?,?)
            """,
            (
                new_id,
                username,
                action,
                nowstr()
            ),
            commit=True
        )

    except Exception as e:
        app.logger.warning(
            "Activity log error: %r",
            e
        )


def current_user():
    uid = session.get("user_id")

    if not uid:
        return None

    try:
        return fetch_one(
            "SELECT * FROM users WHERE id=?",
            (uid,)
        )

    except Exception:
        return None


# =========================================================
# AUTH DECORATORS
# =========================================================

def login_required(f):

    @wraps(f)
    def wrapper(*args, **kwargs):

        if not session.get("user_id"):
            return redirect(
                url_for(
                    "login",
                    next=request.path
                )
            )

        return f(*args, **kwargs)

    return wrapper


def admin_required(f):

    @wraps(f)
    def wrapper(*args, **kwargs):

        u = current_user()

        if (
            not u
            or str(
                u.get("role", "")
            ).lower()
            not in (
                "administrator",
                "admin"
            )
        ):
            flash(
                "Administrator access required.",
                "danger"
            )

            return redirect(
                url_for("dashboard")
            )

        return f(*args, **kwargs)

    return wrapper


# =========================================================
# GENERAL HELPERS
# =========================================================

def month_name_year(month=None, year=None):

    month = (
        month
        or request.values.get("month")
        or MONTHS[
            datetime.date.today().month - 1
        ]
    )

    year = (
        year
        or request.values.get("year")
        or str(datetime.date.today().year)
    )

    if str(month) not in MONTHS:
        month = MONTHS[
            datetime.date.today().month - 1
        ]

    try:
        year = int(year)
    except Exception:
        year = datetime.date.today().year

    return f"{month} {year}"


def parse_num(value, default=0.0):
    try:
        return float(value or default)
    except Exception:
        return default


# =========================================================
# ADVANCE HELPERS
# =========================================================

def legacy_advance_source():

    if table_exists("worker_advances"):

        c = columns("worker_advances")

        if {
            "worker_id",
            "amount"
        }.issubset(c):
            return "worker_advances"

    if table_exists("advance_salary"):
        return "advance_salary"

    if table_exists("advances"):
        return "advances"

    return None


def month_name_from_date(value):

    try:

        d = str(value)[:10]

        y, m, _ = d.split("-")

        return (
            f"{MONTHS[int(m) - 1]} "
            f"{int(y)}"
        )

    except Exception:
        return ""


def advance_rows(
    month=None,
    worker_id=None
):

    table = legacy_advance_source()

    if not table:
        return []

    c = columns(table)

    # Current web table
    if table == "worker_advances":

        where = []
        params = []

        if month:
            where.append("month_year=?")
            params.append(month)

        if worker_id:
            where.append("worker_id=?")
            params.append(worker_id)

        sql = """
            SELECT
                id,
                worker_id,
                month_year,
                advance_date,
                amount,
                note
            FROM worker_advances
        """

        if where:
            sql += (
                " WHERE "
                + " AND ".join(where)
            )

        sql += " ORDER BY id DESC"

        return fetch_all(
            sql,
            params
        )

    # Legacy advance_salary
    if table == "advance_salary":

        datecol = (
            "date"
            if "date" in c
            else (
                "advance_date"
                if "advance_date" in c
                else None
            )
        )

        amountcol = (
            "amount"
            if "amount" in c
            else None
        )

        if not amountcol:
            return []

        sql = f"""
            SELECT
                id,
                worker_id,
                {datecol or 'NULL'} AS advance_date,
                amount,
                note
            FROM {table}
        """

        rows = fetch_all(sql)

        for r in rows:
            r["month_year"] = (
                month_name_from_date(
                    r.get("advance_date")
                )
            )

        if month:
            rows = [
                r for r in rows
                if r.get("month_year") == month
            ]

        if worker_id:
            rows = [
                r for r in rows
                if str(r.get("worker_id"))
                == str(worker_id)
            ]

        return rows

    # Legacy advances
    if table == "advances":

        datecol = (
            "date"
            if "date" in c
            else (
                "advance_date"
                if "advance_date" in c
                else None
            )
        )

        if not datecol:
            return []

        sql = f"""
            SELECT
                id,
                worker_id,
                {datecol} AS advance_date,
                amount,
                note
            FROM {table}
        """

        rows = fetch_all(sql)

        for r in rows:
            r["month_year"] = (
                month_name_from_date(
                    r.get("advance_date")
                )
            )

        if month:
            rows = [
                r for r in rows
                if r.get("month_year") == month
            ]

        if worker_id:
            rows = [
                r for r in rows
                if str(r.get("worker_id"))
                == str(worker_id)
            ]

        return rows

    return []


def save_advance_record(
    worker_id,
    month,
    adv_date,
    amount,
    note
):

    table = legacy_advance_source()

    if table == "worker_advances":

        execute(
            """
            INSERT INTO worker_advances
            (
                worker_id,
                month_year,
                advance_date,
                amount,
                note
            )
            VALUES(?,?,?,?,?)
            """,
            (
                worker_id,
                month,
                adv_date,
                amount,
                note
            ),
            commit=True
        )

        return

    if table == "advance_salary":

        c = columns(table)

        datecol = (
            "date"
            if "date" in c
            else "advance_date"
        )

        names = ["worker_id"]
        vals = [worker_id]

        if datecol:
            names.append(datecol)
            vals.append(adv_date)

        if "amount" in c:
            names.append("amount")
            vals.append(amount)

        if "note" in c:
            names.append("note")
            vals.append(note)

        execute(
            "INSERT INTO "
            + table
            + "("
            + ",".join(names)
            + ") VALUES("
            + ",".join(["?"] * len(vals))
            + ")",
            vals,
            commit=True
        )

        return

    if table == "advances":

        c = columns(table)

        names = []
        vals = []

        if "id" in c:
            names.append("id")
            vals.append(next_id(table))

        if "worker_id" in c:
            names.append("worker_id")
            vals.append(worker_id)

        if "date" in c:
            names.append("date")
            vals.append(adv_date)
        elif "advance_date" in c:
            names.append("advance_date")
            vals.append(adv_date)

        if "amount" in c:
            names.append("amount")
            vals.append(amount)

        if "note" in c:
            names.append("note")
            vals.append(note)

        execute(
            "INSERT INTO "
            + table
            + "("
            + ",".join(names)
            + ") VALUES("
            + ",".join(["?"] * len(vals))
            + ")",
            vals,
            commit=True
        )

        return

    # No advance table exists.
    if is_postgres():

        execute(
            """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount REAL NOT NULL,
                note TEXT
            )
            """,
            commit=True
        )

    else:

        execute(
            """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount REAL NOT NULL,
                note TEXT
            )
            """,
            commit=True
        )

    execute(
        """
        INSERT INTO worker_advances
        (
            worker_id,
            month_year,
            advance_date,
            amount,
            note
        )
        VALUES(?,?,?,?,?)
        """,
        (
            worker_id,
            month,
            adv_date,
            amount,
            note
        ),
        commit=True
    )


def delete_advance_record(aid):

    table = legacy_advance_source()

    if table:
        execute(
            f"DELETE FROM {table} WHERE id=?",
            (aid,),
            commit=True
        )


# =========================================================
# ATTENDANCE / SALARY CALCULATION
# =========================================================

def daily_map(worker_id, month):

    rows = fetch_all(
        """
        SELECT day,status
        FROM daily_attendance
        WHERE worker_id=?
        AND month_year=?
        """,
        (
            worker_id,
            month
        )
    )

    return {
        int(r["day"]):
            str(r["status"] or "P")
        for r in rows
    }


def calculate_salary_bulk(
    workers,
    month
):

    ids = [
        w["id"]
        for w in workers
    ]

    if not ids:
        return {}

    # Attendance
    att_rows = fetch_all(
        """
        SELECT *
        FROM attendance
        WHERE month_year=?
        ORDER BY id DESC
        """,
        (month,)
    )

    att_by = {}

    for r in att_rows:

        wid = r.get("worker_id")

        if wid not in att_by:
            att_by[wid] = r

    # Daily attendance
    daily_rows = fetch_all(
        """
        SELECT worker_id,day,status
        FROM daily_attendance
        WHERE month_year=?
        """,
        (month,)
    )

    daily_by = {}

    for r in daily_rows:

        try:

            daily_by.setdefault(
                r.get("worker_id"),
                {}
            )[
                int(r.get("day"))
            ] = str(
                r.get("status") or "P"
            )

        except Exception:
            pass

    # Advances
    advances_by = {}

    table = legacy_advance_source()

    if table:

        c = columns(table)

        try:

            if (
                table == "worker_advances"
                and {
                    "worker_id",
                    "amount"
                }.issubset(c)
            ):

                adv_rows = fetch_all(
                    """
                    SELECT
                        worker_id,
                        amount,
                        month_year
                    FROM worker_advances
                    WHERE month_year=?
                    """,
                    (month,)
                )

                for r in adv_rows:

                    wid = r.get("worker_id")

                    advances_by[wid] = (
                        advances_by.get(wid, 0)
                        + parse_num(
                            r.get("amount")
                        )
                    )

            elif (
                table in (
                    "advance_salary",
                    "advances"
                )
                and "worker_id" in c
                and "amount" in c
            ):

                datecol = (
                    "date"
                    if "date" in c
                    else (
                        "advance_date"
                        if "advance_date" in c
                        else None
                    )
                )

                if datecol:

                    adv_rows = fetch_all(
                        f"""
                        SELECT
                            worker_id,
                            {datecol} AS advance_date,
                            amount
                        FROM {table}
                        """
                    )

                    for r in adv_rows:

                        if (
                            month_name_from_date(
                                r.get("advance_date")
                            ) == month
                        ):

                            wid = r.get(
                                "worker_id"
                            )

                            advances_by[wid] = (
                                advances_by.get(
                                    wid,
                                    0
                                )
                                + parse_num(
                                    r.get(
                                        "amount"
                                    )
                                )
                            )

        except Exception as e:

            app.logger.warning(
                "Bulk advance load error: %r",
                e
            )

    # Calculate
    output = {}

    try:
        mon, ys = month.split()
        year = int(ys)
        mi = MONTHS.index(mon) + 1
        days = calendar.monthrange(
            year,
            mi
        )[1]

    except Exception:

        year = datetime.date.today().year
        mi = datetime.date.today().month
        days = 30

    for w in workers:

        att = att_by.get(
            w["id"],
            {}
        )

        present = int(
            att.get("present_days") or 0
        )

        absent = int(
            att.get("absent_days") or 0
        )

        ot = parse_num(
            att.get("ot_hours"),
            0
        )

        basic = parse_num(
            w.get("basic_salary")
        )

        ot_rate = parse_num(
            w.get("ot_rate")
        )

        nasta_rate = parse_num(
            w.get("refreshment_bill")
        )

        absent_cut = (
            basic / days
        ) * absent if days else 0

        earned_basic = max(
            0,
            basic - absent_cut
        )

        ot_amt = ot * ot_rate

        dm = daily_by.get(
            w["id"],
            {}
        )

        if dm:

            billable = sum(
                1
                for d in range(
                    1,
                    days + 1
                )
                if (
                    dm.get(d, "P") == "P"
                    and datetime.date(
                        year,
                        mi,
                        d
                    ).weekday() != 4
                )
            )

            nasta = (
                billable
                * nasta_rate
            )

        else:

            non_friday = sum(
                1
                for d in range(
                    1,
                    days + 1
                )
                if datetime.date(
                    year,
                    mi,
                    d
                ).weekday() != 4
            )

            nasta = (
                round(
                    present
                    * (
                        non_friday
                        / days
                    )
                    * nasta_rate,
                    2
                )
                if days
                else 0
            )

        advance = advances_by.get(
            w["id"],
            0
        )

        gross = (
            earned_basic
            + ot_amt
            + nasta
        )

        output[w["id"]] = {
            "present": present,
            "absent": absent,
            "ot": ot,
            "absent_cut": absent_cut,
            "earned_basic": earned_basic,
            "ot_amt": ot_amt,
            "nasta": nasta,
            "gross": gross,
            "advance": advance,
            "net": gross - advance
        }

    return output


def calculate_salary(
    worker,
    month,
    att=None
):

    if not worker:
        return {}

    if att is None:

        att = fetch_one(
            """
            SELECT *
            FROM attendance
            WHERE worker_id=?
            AND month_year=?
            ORDER BY id DESC
            LIMIT 1
            """,
            (
                worker["id"],
                month
            )
        ) or {}

    present = int(
        att.get("present_days") or 0
    )

    absent = int(
        att.get("absent_days") or 0
    )

    ot = parse_num(
        att.get("ot_hours"),
        0
    )

    try:

        mon, ys = month.split()

        year = int(ys)

        mi = MONTHS.index(mon) + 1

        days = calendar.monthrange(
            year,
            mi
        )[1]

    except Exception:

        year = datetime.date.today().year
        mi = datetime.date.today().month
        days = 30

    basic = parse_num(
        worker.get("basic_salary")
    )

    ot_rate = parse_num(
        worker.get("ot_rate")
    )

    nasta_rate = parse_num(
        worker.get("refreshment_bill")
    )

    absent_cut = (
        basic / days
    ) * absent if days else 0

    earned_basic = max(
        0,
        basic - absent_cut
    )

    ot_amt = (
        ot * ot_rate
    )

    dm = daily_map(
        worker["id"],
        month
    )

    if dm:

        billable = sum(
            1
            for d in range(
                1,
                days + 1
            )
            if (
                dm.get(d, "P") == "P"
                and datetime.date(
                    year,
                    mi,
                    d
                ).weekday() != 4
            )
        )

        nasta = (
            billable
            * nasta_rate
        )

    else:

        non_friday = sum(
            1
            for d in range(
                1,
                days + 1
            )
            if datetime.date(
                year,
                mi,
                d
            ).weekday() != 4
        )

        nasta = (
            round(
                present
                * (
                    non_friday
                    / days
                )
                * nasta_rate,
                2
            )
            if days
            else 0
        )

    advance = sum(
        parse_num(
            r.get("amount")
        )
        for r in advance_rows(
            month,
            worker["id"]
        )
    )

    gross = (
        earned_basic
        + ot_amt
        + nasta
    )

    return {
        "present": present,
        "absent": absent,
        "ot": ot,
        "absent_cut": absent_cut,
        "earned_basic": earned_basic,
        "ot_amt": ot_amt,
        "nasta": nasta,
        "gross": gross,
        "advance": advance,
        "net": gross - advance
    }


# =========================================================
# GLOBAL TEMPLATE VARIABLES
# =========================================================

@app.context_processor
def inject_globals():

    settings = get_settings()

    lang = session.get(
        "language",
        "en"
    )

    def tr(text):
        if lang == "bn":
            return LANG.get(
                text,
                text
            )

        return text

    def dept(text):

        if lang == "bn":
            return DEPT_BN.get(
                str(text),
                str(text)
            )

        return str(text or "")

    return {
        "current_user": current_user(),
        "settings": settings,
        "language": lang,
        "tr": tr,
        "display_dept": dept,
        "months": MONTHS,
        "years": list(
            range(
                2024,
                2032
            )
        )
    }


# =========================================================
# HEALTH
# =========================================================

@app.route("/health")
def health():

    try:

        workers_count = scalar(
            "SELECT COUNT(*) FROM workers",
            default=0
        )

        attendance_count = scalar(
            "SELECT COUNT(*) FROM attendance",
            default=0
        )

        daily_count = scalar(
            "SELECT COUNT(*) FROM daily_attendance",
            default=0
        )

        return jsonify({
            "status": "ok",
            "database": (
                "postgresql"
                if is_postgres()
                else "sqlite"
            ),
            "workers": workers_count,
            "attendance": attendance_count,
            "daily_attendance": daily_count
        })

    except Exception as e:

        return jsonify({
            "status": "error",
            "error": repr(e)
        }), 500


# =========================================================
# DATABASE CHECK
# =========================================================

@app.route("/db-check")
@login_required
def db_check():

    info = {}

    for table in [
        "workers",
        "attendance",
        "daily_attendance",
        "company_settings",
        "users",
        "activity_log",
        "worker_advances",
        "advance_salary",
        "advances"
    ]:

        try:

            exists = table_exists(table)

            info[table] = {
                "exists": exists,
                "columns": (
                    sorted(columns(table))
                    if exists
                    else []
                )
            }

        except Exception as e:

            info[table] = {
                "error": repr(e)
            }

    return render_template(
        "db_check.html",
        info=info
    )


# =========================================================
# LOGIN
# =========================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        user = fetch_one(
            """
            SELECT *
            FROM users
            WHERE username=?
            LIMIT 1
            """,
            (username,)
        )

        ok = False

        if user:

            active = user.get(
                "active"
            )

            if active is None:
                active = 1

            if int(active):

                stored = (
                    user.get(
                        "password_hash"
                    )
                    or user.get(
                        "password"
                    )
                    or ""
                )

                ok = (
                    stored
                    == hash_password(
                        password
                    )
                    or stored == password
                )

        if ok:

            session["user_id"] = user["id"]

            session["language"] = session.get(
                "language",
                "en"
            )

            try:

                execute(
                    """
                    UPDATE users
                    SET last_login=?
                    WHERE id=?
                    """,
                    (
                        nowstr(),
                        user["id"]
                    ),
                    commit=True
                )

            except Exception:
                pass

            log_activity(
                username,
                "Logged in"
            )

            return redirect(
                request.args.get(
                    "next"
                )
                or url_for(
                    "dashboard"
                )
            )

        flash(
            "Invalid User ID or Password.",
            "danger"
        )

    return render_template(
        "login.html"
    )


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    user = current_user()

    if user:
        log_activity(
            user.get("username"),
            "Logged out"
        )

    session.clear()

    return redirect(
        url_for("login")
    )


# =========================================================
# LANGUAGE
# =========================================================

@app.route("/language/<lang>")
def language(lang):

    session["language"] = (
        "bn"
        if lang == "bn"
        else "en"
    )

    return redirect(
        request.referrer
        or url_for("dashboard")
    )


# =========================================================
# HOME
# =========================================================

@app.route("/")
@login_required
def index():

    return redirect(
        url_for("dashboard")
    )


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
@login_required
def dashboard():

    month = month_name_year()

    workers_list = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    salaries = calculate_salary_bulk(
        workers_list,
        month
    )

    gross = sum(
        s["gross"]
        for s in salaries.values()
    )

    advance = sum(
        s["advance"]
        for s in salaries.values()
    )

    today = datetime.date.today()

    today_month = (
        f"{MONTHS[today.month - 1]} "
        f"{today.year}"
    )

    today_rows = fetch_all(
        """
        SELECT
            status,
            COUNT(*) AS c
        FROM daily_attendance
        WHERE month_year=?
        AND day=?
        GROUP BY status
        """,
        (
            today_month,
            today.day
        )
    )

    today_map = {
        r["status"]: r["c"]
        for r in today_rows
    }

    department_rows = fetch_all(
        """
        SELECT
            department,
            COUNT(*) AS worker_count,
            COALESCE(
                SUM(basic_salary),
                0
            ) AS total_basic
        FROM workers
        GROUP BY department
        ORDER BY department
        """
    )

    return render_template(
        "dashboard.html",
        workers=workers_list,
        month=month,
        month_name=month.split()[0],
        year=month.split()[1],
        gross=gross,
        advance=advance,
        today_present=today_map.get(
            "P",
            0
        ),
        today_absent=today_map.get(
            "A",
            0
        ),
        dept_rows=department_rows,
        total_workers=len(
            workers_list
        )
    )


# =========================================================
# WORKERS
# =========================================================

@app.route("/workers")
@login_required
def workers():

    q = request.args.get(
        "q",
        ""
    ).strip().lower()

    params = []

    sql = """
        SELECT *
        FROM workers
    """

    if q:

        like = f"%{q}%"

        sql += """
            WHERE
                CAST(id AS TEXT) LIKE ?
                OR lower(name) LIKE ?
                OR lower(
                    COALESCE(
                        bangla_name,
                        ''
                    )
                ) LIKE ?
                OR lower(
                    COALESCE(
                        department,
                        ''
                    )
                ) LIKE ?
                OR lower(
                    COALESCE(
                        designation,
                        ''
                    )
                ) LIKE ?
        """

        params = [
            like,
            like,
            like,
            like,
            like
        ]

    sql += " ORDER BY id DESC"

    return render_template(
        "workers.html",
        workers=fetch_all(
            sql,
            params
        ),
        q=q
    )


@app.route(
    "/workers/add",
    methods=["GET", "POST"]
)
@login_required
def add_worker():

    if request.method == "POST":

        form = request.form

        name = form.get(
            "name",
            ""
        ).strip()

        if not name:

            flash(
                "Worker name is required.",
                "danger"
            )

            return render_template(
                "worker_form.html",
                worker=form
            )

        new_id = next_id(
            "workers"
        )

        values = (
            new_id,
            name,
            parse_num(
                form.get(
                    "basic_salary"
                )
            ),
            parse_num(
                form.get(
                    "ot_rate"
                )
            ),
            form.get(
                "department",
                ""
            ).strip(),
            form.get(
                "designation",
                ""
            ).strip(),
            parse_num(
                form.get(
                    "refreshment_bill"
                )
            ),
            form.get(
                "bangla_name",
                ""
            ).strip()
        )

        execute(
            """
            INSERT INTO workers
            (
                id,
                name,
                basic_salary,
                ot_rate,
                department,
                designation,
                refreshment_bill,
                bangla_name
            )
            VALUES(?,?,?,?,?,?,?,?)
            """,
            values,
            commit=True
        )

        flash(
            "Worker saved.",
            "success"
        )

        return redirect(
            url_for("workers")
        )

    return render_template(
        "worker_form.html",
        worker=None
    )


@app.route(
    "/workers/edit/<int:worker_id>",
    methods=["GET", "POST"]
)
@login_required
def edit_worker(worker_id):

    worker = fetch_one(
        """
        SELECT *
        FROM workers
        WHERE id=?
        """,
        (worker_id,)
    )

    if not worker:
        abort(404)

    if request.method == "POST":

        form = request.form

        execute(
            """
            UPDATE workers
            SET
                name=?,
                basic_salary=?,
                ot_rate=?,
                department=?,
                designation=?,
                refreshment_bill=?,
                bangla_name=?
            WHERE id=?
            """,
            (
                form.get(
                    "name",
                    ""
                ),
                parse_num(
                    form.get(
                        "basic_salary"
                    )
                ),
                parse_num(
                    form.get(
                        "ot_rate"
                    )
                ),
                form.get(
                    "department",
                    ""
                ),
                form.get(
                    "designation",
                    ""
                ),
                parse_num(
                    form.get(
                        "refreshment_bill"
                    )
                ),
                form.get(
                    "bangla_name",
                    ""
                ),
                worker_id
            ),
            commit=True
        )

        flash(
            "Worker updated.",
            "success"
        )

        return redirect(
            url_for("workers")
        )

    return render_template(
        "worker_form.html",
        worker=worker
    )


@app.route(
    "/workers/delete/<int:worker_id>",
    methods=["GET", "POST"]
)
@login_required
def delete_worker(worker_id):

    execute(
        "DELETE FROM workers WHERE id=?",
        (worker_id,),
        commit=True
    )

    # Explicit user action only.
    for table in [
        "attendance",
        "daily_attendance",
        "worker_advances",
        "advance_salary",
        "advances"
    ]:

        if table_exists(table):

            try:

                execute(
                    f"""
                    DELETE FROM {table}
                    WHERE worker_id=?
                    """,
                    (worker_id,),
                    commit=True
                )

            except Exception:
                pass

    flash(
        "Worker deleted.",
        "success"
    )

    return redirect(
        url_for("workers")
    )


# =========================================================
# ATTENDANCE
# =========================================================

@app.route("/attendance")
@login_required
def attendance():

    month = month_name_year()

    worker_id = request.args.get(
        "worker_id",
        ""
    )

    worker = None
    days = []
    summary = {}

    if worker_id:

        worker = fetch_one(
            """
            SELECT *
            FROM workers
            WHERE id=?
            """,
            (worker_id,)
        )

        if worker:

            dm = daily_map(
                worker["id"],
                month
            )

            mon, ys = month.split()

            year = int(ys)

            mi = MONTHS.index(
                mon
            ) + 1

            number_days = calendar.monthrange(
                year,
                mi
            )[1]

            days = [
                {
                    "day": d,
                    "status": dm.get(
                        d,
                        "P"
                    ),
                    "date": datetime.date(
                        year,
                        mi,
                        d
                    )
                }
                for d in range(
                    1,
                    number_days + 1
                )
            ]

            summary = calculate_salary(
                worker,
                month
            )

    return render_template(
        "attendance.html",
        workers=fetch_all(
            """
            SELECT id,name,department
            FROM workers
            ORDER BY id
            """
        ),
        worker=worker,
        days=days,
        summary=summary,
        month=month
    )


@app.route(
    "/attendance/save",
    methods=["POST"]
)
@login_required
def save_attendance():

    form = request.form

    worker_id = int(
        form.get(
            "worker_id"
        )
    )

    if form.get("year"):

        month = month_name_year(
            form.get("month"),
            form.get("year")
        )

    else:

        month = (
            form.get("month_year")
            or month_name_year()
        )

    present = int(
        form.get(
            "present_days"
        )
        or 0
    )

    absent = int(
        form.get(
            "absent_days"
        )
        or 0
    )

    ot = parse_num(
        form.get(
            "ot_hours"
        )
    )

    # Daily statuses
    daily_status = {}

    raw = form.get(
        "statuses_json"
    )

    if raw:

        try:
            import json

            daily_status = json.loads(
                raw
            )

        except Exception:
            daily_status = {}

    # day_1=P format
    for key, value in form.items():

        if key.startswith("day_"):

            daily_status[
                key[4:]
            ] = value

    # Save daily attendance
    for day, status in daily_status.items():

        try:

            day = int(day)

            status = (
                "A"
                if str(
                    status
                ).upper().startswith("A")
                else "P"
            )

            existing = fetch_one(
                """
                SELECT id
                FROM daily_attendance
                WHERE worker_id=?
                AND month_year=?
                AND day=?
                ORDER BY id DESC
                LIMIT 1
                """,
                (
                    worker_id,
                    month,
                    day
                )
            )

            if existing:

                execute(
                    """
                    UPDATE daily_attendance
                    SET status=?
                    WHERE id=?
                    """,
                    (
                        status,
                        existing["id"]
                    ),
                    commit=True
                )

            else:

                new_id = next_id(
                    "daily_attendance"
                )

                execute(
                    """
                    INSERT INTO daily_attendance
                    (
                        id,
                        worker_id,
                        month_year,
                        day,
                        status
                    )
                    VALUES(?,?,?,?,?)
                    """,
                    (
                        new_id,
                        worker_id,
                        month,
                        day,
                        status
                    ),
                    commit=True
                )

        except Exception as e:

            app.logger.warning(
                "Daily attendance error: %r",
                e
            )

    # Monthly attendance
    existing = fetch_one(
        """
        SELECT id
        FROM attendance
        WHERE worker_id=?
        AND month_year=?
        ORDER BY id DESC
        LIMIT 1
        """,
        (
            worker_id,
            month
        )
    )

    if existing:

        execute(
            """
            UPDATE attendance
            SET
                present_days=?,
                absent_days=?,
                ot_hours=?
            WHERE id=?
            """,
            (
                present,
                absent,
                ot,
                existing["id"]
            ),
            commit=True
        )

    else:

        attendance_columns = columns(
            "attendance"
        )

        names = [
            "id",
            "worker_id",
            "month_year",
            "present_days",
            "absent_days",
            "ot_hours"
        ]

        values = [
            next_id("attendance"),
            worker_id,
            month,
            present,
            absent,
            ot
        ]

        if "advance_deduction" in attendance_columns:

            names.append(
                "advance_deduction"
            )

            values.append(0)

        execute(
            "INSERT INTO attendance("
            + ",".join(names)
            + ") VALUES("
            + ",".join(
                ["?"] * len(values)
            )
            + ")",
            values,
            commit=True
        )

    flash(
        "Attendance saved.",
        "success"
    )

    return redirect(
        url_for(
            "attendance",
            worker_id=worker_id,
            month=month.split()[0],
            year=month.split()[1]
        )
    )


# =========================================================
# PAYSLIP
# =========================================================

@app.route("/payslip")
@login_required
def payslip():

    month = month_name_year()

    worker_id = request.args.get(
        "worker_id"
    )

    worker = (
        fetch_one(
            """
            SELECT *
            FROM workers
            WHERE id=?
            """,
            (worker_id,)
        )
        if worker_id
        else None
    )

    summary = (
        calculate_salary(
            worker,
            month
        )
        if worker
        else {}
    )

    return render_template(
        "payslip.html",
        workers=fetch_all(
            """
            SELECT id,name,department
            FROM workers
            ORDER BY id
            """
        ),
        worker=worker,
        summary=summary,
        month=month
    )


# =========================================================
# DEPARTMENT
# =========================================================

@app.route("/department")
@login_required
def department():

    month = month_name_year()

    department_name = request.args.get(
        "department",
        ""
    )

    if department_name:

        worker_rows = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (department_name,)
        )

    else:

        worker_rows = fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    salary_map = calculate_salary_bulk(
        worker_rows,
        month
    )

    rows = [
        {
            **worker,
            **salary_map.get(
                worker["id"],
                {}
            )
        }
        for worker in worker_rows
    ]

    departments = [
        r["department"]
        for r in fetch_all(
            """
            SELECT DISTINCT department
            FROM workers
            WHERE department IS NOT NULL
            AND department<>''
            ORDER BY department
            """
        )
    ]

    return render_template(
        "department.html",
        rows=rows,
        month=month,
        department=department_name,
        departments=departments
    )


# =========================================================
# ADVANCE
# =========================================================

@app.route("/advance")
@app.route("/advances")
@login_required
def advance():

    month = month_name_year()

    worker_id = request.args.get(
        "worker_id"
    )

    worker_list = fetch_all(
        """
        SELECT id,name,department
        FROM workers
        ORDER BY id
        """
    )

    rows = advance_rows(
        month,
        worker_id
    )

    return render_template(
        "advance.html",
        workers=worker_list,
        rows=rows,
        month=month
    )


@app.route(
    "/advance/save",
    methods=["POST"]
)
@app.route(
    "/advances/save",
    methods=["POST"]
)
@login_required
def save_advance():

    form = request.form

    worker_id = int(
        form.get(
            "worker_id"
        )
    )

    month = (
        form.get("month_year")
        or month_name_year(
            form.get("month"),
            form.get("year")
        )
    )

    advance_date = (
        form.get(
            "advance_date"
        )
        or datetime.date.today().isoformat()
    )

    amount = parse_num(
        form.get(
            "amount"
        )
    )

    note = form.get(
        "note",
        ""
    )

    if amount <= 0:

        flash(
            "Amount must be greater than zero.",
            "danger"
        )

        return redirect(
            url_for(
                "advance"
            )
        )

    save_advance_record(
        worker_id,
        month,
        advance_date,
        amount,
        note
    )

    flash(
        "Advance saved.",
        "success"
    )

    return redirect(
        url_for(
            "advance"
        )
    )


@app.route(
    "/advance/delete/<int:aid>"
)
@app.route(
    "/advances/delete/<int:aid>"
)
@login_required
def delete_advance(aid):

    delete_advance_record(
        aid
    )

    flash(
        "Advance deleted.",
        "success"
    )

    return redirect(
        url_for("advance")
    )


# =========================================================
# SETTINGS
# =========================================================

@app.route(
    "/settings",
    methods=["GET", "POST"]
)
@admin_required
def settings():

    if request.method == "POST":

        for key in [
            "company_name",
            "company_address",
            "company_phone",
            "company_email",
            "company_logo"
        ]:

            value = request.form.get(
                key,
                ""
            )

            if is_postgres():

                execute(
                    """
                    INSERT INTO company_settings
                    (key,value)
                    VALUES(?,?)
                    ON CONFLICT(key)
                    DO UPDATE SET
                    value=EXCLUDED.value
                    """,
                    (
                        key,
                        value
                    ),
                    commit=True
                )

            else:

                execute(
                    """
                    INSERT OR REPLACE
                    INTO company_settings
                    (key,value)
                    VALUES(?,?)
                    """,
                    (
                        key,
                        value
                    ),
                    commit=True
                )

        flash(
            "Settings saved.",
            "success"
        )

    return render_template(
        "settings.html",
        settings=get_settings()
    )


# =========================================================
# USERS
# =========================================================

@app.route("/users")
@admin_required
def users():

    return render_template(
        "users.html",
        users=fetch_all(
            """
            SELECT
                id,
                username,
                full_name,
                role,
                active,
                created_at,
                last_login
            FROM users
            ORDER BY id
            """
        )
    )


@app.route(
    "/users/add",
    methods=["GET", "POST"]
)
@admin_required
def add_user():

    if request.method == "POST":

        form = request.form

        username = form.get(
            "username",
            ""
        ).strip()

        password = form.get(
            "password",
            ""
        )

        if not username or not password:

            flash(
                "Username and password are required.",
                "danger"
            )

            return render_template(
                "user_form.html",
                user=form
            )

        password_hash = hash_password(
            password
        )

        user_columns = columns(
            "users"
        )

        names = []
        values = []

        data = [
            (
                "id",
                next_id("users")
            ),
            (
                "username",
                username
            ),
            (
                "password_hash",
                password_hash
            ),
            (
                "password",
                password_hash
            ),
            (
                "full_name",
                form.get(
                    "full_name",
                    ""
                )
            ),
            (
                "role",
                form.get(
                    "role",
                    "Operator"
                )
            ),
            (
                "active",
                1
            ),
            (
                "created_at",
                nowstr()
            )
        ]

        for name, value in data:

            if name in user_columns:

                names.append(name)
                values.append(value)

        try:

            execute(
                "INSERT INTO users("
                + ",".join(names)
                + ") VALUES("
                + ",".join(
                    ["?"] * len(values)
                )
                + ")",
                values,
                commit=True
            )

            flash(
                "User created.",
                "success"
            )

            return redirect(
                url_for("users")
            )

        except Exception as e:

            flash(
                f"Could not create user: {e}",
                "danger"
            )

    return render_template(
        "user_form.html",
        user=None
    )


@app.route(
    "/users/delete/<int:user_id>"
)
@admin_required
def delete_user(user_id):

    user = current_user()

    if (
        user
        and user.get("id") == user_id
    ):

        flash(
            "You cannot delete the logged-in user.",
            "danger"
        )

    else:

        execute(
            "DELETE FROM users WHERE id=?",
            (user_id,),
            commit=True
        )

        flash(
            "User deleted.",
            "success"
        )

    return redirect(
        url_for("users")
    )


# =========================================================
# ACTIVITY
# =========================================================

@app.route("/activity")
@admin_required
def activity():

    return render_template(
        "activity.html",
        logs=fetch_all(
            """
            SELECT *
            FROM activity_log
            ORDER BY id DESC
            LIMIT 500
            """
        )
    )


# =========================================================
# EXPORT EXCEL
# =========================================================

def export_rows(
    rows,
    filename,
    title="Reedoy Payroll"
):

    if openpyxl is None:

        abort(
            503,
            description=(
                "openpyxl is not installed"
            )
        )

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Payroll"

    if rows:

        headers = list(
            rows[0].keys()
        )

        worksheet.append(
            headers
        )

        for row in rows:

            worksheet.append(
                [
                    row.get(header)
                    for header in headers
                ]
            )

    else:

        worksheet.append(
            ["No records"]
        )

    bio = io.BytesIO()

    workbook.save(
        bio
    )

    bio.seek(0)

    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )


@app.route(
    "/export/workers.xlsx"
)
@login_required
def export_workers():

    return export_rows(
        fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        ),
        "reedoy_workers.xlsx"
    )


@app.route(
    "/export/advances.xlsx"
)
@login_required
def export_advances():

    return export_rows(
        advance_rows(
            month_name_year()
        ),
        "reedoy_advances.xlsx"
    )


@app.route(
    "/export/department.xlsx"
)
@login_required
def export_department():

    month = month_name_year()

    department_name = request.args.get(
        "department",
        ""
    )

    if department_name:

        worker_rows = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (department_name,)
        )

    else:

        worker_rows = fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    salary_map = calculate_salary_bulk(
        worker_rows,
        month
    )

    rows = []

    for worker in worker_rows:

        salary = salary_map.get(
            worker["id"],
            {}
        )

        rows.append({
            "ID": worker["id"],
            "Worker Name": worker["name"],
            "Department": worker.get(
                "department"
            ),
            "Designation": worker.get(
                "designation"
            ),
            **salary
        })

    return export_rows(
        rows,
        "reedoy_department_salary.xlsx"
    )


# =========================================================
# PAYSLIP PDF
# =========================================================

@app.route(
    "/export/payslip.pdf"
)
@login_required
def export_payslip_pdf():

    if SimpleDocTemplate is None:

        abort(
            503,
            description=(
                "reportlab is not installed"
            )
        )

    worker_id = request.args.get(
        "worker_id"
    )

    month = month_name_year()

    worker = fetch_one(
        """
        SELECT *
        FROM workers
        WHERE id=?
        """,
        (worker_id,)
    )

    if not worker:
        abort(404)

    salary = calculate_salary(
        worker,
        month
    )

    bio = io.BytesIO()

    document = SimpleDocTemplate(
        bio,
        pagesize=A4,
        rightMargin=35,
        leftMargin=35,
        topMargin=35,
        bottomMargin=35
    )

    styles = getSampleStyleSheet()

    story = [
        Paragraph(
            DEFAULT_COMPANY_NAME,
            styles["Title"]
        ),
        Paragraph(
            "Payroll Payslip - "
            + month,
            styles["Heading2"]
        ),
        Spacer(1, 12)
    ]

    data = [
        [
            "Worker ID",
            worker["id"]
        ],
        [
            "Worker Name",
            worker["name"]
        ],
        [
            "Department",
            worker.get(
                "department"
            ) or ""
        ],
        [
            "Basic Salary",
            f"BDT {salary['earned_basic']:,.2f}"
        ],
        [
            "Absent Deduction",
            f"BDT {salary['absent_cut']:,.2f}"
        ],
        [
            "OT",
            f"BDT {salary['ot_amt']:,.2f}"
        ],
        [
            "Nasta",
            f"BDT {salary['nasta']:,.2f}"
        ],
        [
            "Gross",
            f"BDT {salary['gross']:,.2f}"
        ],
        [
            "Advance",
            f"BDT {salary['advance']:,.2f}"
        ],
        [
            "Net Payable",
            f"BDT {salary['net']:,.2f}"
        ]
    ]

    story.append(
        Table(
            data,
            colWidths=[
                180,
                280
            ],
            style=TableStyle([
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.grey
                ),
                (
                    "FONTNAME",
                    (0, 0),
                    (-1, -1),
                    "Helvetica"
                ),
                (
                    "PADDING",
                    (0, 0),
                    (-1, -1),
                    6
                )
            ])
        )
    )

    document.build(
        story
    )

    bio.seek(0)

    filename = (
        f"payslip_"
        f"{worker['id']}_"
        f"{month.replace(' ', '_')}.pdf"
    )

    return send_file(
        bio,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf"
    )


# =========================================================
# CSV EXPORT
# =========================================================

@app.route(
    "/export/workers.csv"
)
@login_required
def export_workers_csv():

    rows = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    text = io.StringIO()

    if rows:

        writer = csv.DictWriter(
            text,
            fieldnames=list(
                rows[0].keys()
            )
        )

        writer.writeheader()

        writer.writerows(rows)

    else:

        text.write(
            "No records\n"
        )

    return send_file(
        io.BytesIO(
            text.getvalue().encode(
                "utf-8-sig"
            )
        ),
        as_attachment=True,
        download_name=(
            "reedoy_workers.csv"
        ),
        mimetype="text/csv"
    )


# =========================================================
# =========================================================
# LEGACY TEMPLATE COMPATIBILITY ROUTES
# =========================================================
# এই অংশটি খুব গুরুত্বপূর্ণ।
# পুরোনো HTML template-এ অন্য endpoint name থাকলেও
# এখন Flask সেগুলো চিনবে।
# =========================================================


# Dashboard / report
app.add_url_rule(
    "/report",
    endpoint="report",
    view_func=department,
    methods=["GET"]
)


# Workers
app.add_url_rule(
    "/worker-management",
    endpoint="worker_management",
    view_func=workers,
    methods=["GET"]
)

app.add_url_rule(
    "/worker-list",
    endpoint="worker_list",
    view_func=workers,
    methods=["GET"]
)

app.add_url_rule(
    "/worker",
    endpoint="worker",
    view_func=workers,
    methods=["GET"]
)


# Attendance
app.add_url_rule(
    "/attendance-calendar",
    endpoint="attendance_calendar",
    view_func=attendance,
    methods=["GET"]
)

app.add_url_rule(
    "/attendance-calendar",
    endpoint="attendance_page",
    view_func=attendance,
    methods=["GET"]
)

app.add_url_rule(
    "/save-attendance",
    endpoint="save_attendance_legacy",
    view_func=save_attendance,
    methods=["POST"]
)


# Payslip
app.add_url_rule(
    "/single-payslip",
    endpoint="single_payslip",
    view_func=payslip,
    methods=["GET"]
)

app.add_url_rule(
    "/single-payslip",
    endpoint="single_payslip_page",
    view_func=payslip,
    methods=["GET"]
)


# Department salary
app.add_url_rule(
    "/department-salary-sheet",
    endpoint="department_salary_sheet",
    view_func=department,
    methods=["GET"]
)

app.add_url_rule(
    "/department-salary",
    endpoint="department_salary",
    view_func=department,
    methods=["GET"]
)


# Advance salary
app.add_url_rule(
    "/advance-salary",
    endpoint="advance_salary",
    view_func=advance,
    methods=["GET"]
)

app.add_url_rule(
    "/advance_salary",
    endpoint="advance_salary_page",
    view_func=advance,
    methods=["GET"]
)

app.add_url_rule(
    "/advance-salary/save",
    endpoint="advance_salary_save",
    view_func=save_advance,
    methods=["POST"]
)

app.add_url_rule(
    "/advance_salary/save",
    endpoint="advance_salary_save_page",
    view_func=save_advance,
    methods=["POST"]
)


# Worker add/edit compatibility
app.add_url_rule(
    "/worker/add",
    endpoint="worker_add",
    view_func=add_worker,
    methods=["GET", "POST"]
)

app.add_url_rule(
    "/worker/edit/<int:worker_id>",
    endpoint="worker_edit",
    view_func=edit_worker,
    methods=["GET", "POST"]
)

app.add_url_rule(
    "/worker/delete/<int:worker_id>",
    endpoint="worker_delete",
    view_func=delete_worker,
    methods=["GET", "POST"]
)


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(404)
def not_found(error):

    try:

        return (
            render_template(
                "404.html"
            ),
            404
        )

    except Exception:

        return (
            "404 - Page not found",
            404
        )


@app.errorhandler(500)
def server_error(error):

    app.logger.exception(
        "Unhandled application error"
    )

    try:

        return (
            render_template(
                "500.html",
                error=error
            ),
            500
        )

    except Exception:

        return (
            f"500 - Internal Server Error\n{error}",
            500
        )


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

try:

    init_db()

except Exception as e:

    app.logger.exception(
        "Database initialization error: %r",
        e
    )


# =========================================================
# LOCAL RUN
# =========================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000
            )
        ),
        debug=False
    )
