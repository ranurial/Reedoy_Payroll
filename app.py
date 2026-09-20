import os
import csv
import io
import sqlite3
import hashlib
import calendar
import datetime
from functools import wraps

from flask import (
    Flask,
    render_template,
    render_template_string,
    request,
    redirect,
    url_for,
    session,
    flash,
    send_file,
    abort,
    jsonify,
)

# ============================================================
# OPTIONAL EXPORT LIBRARIES
# ============================================================

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
        TableStyle,
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
    colors = None
    getSampleStyleSheet = None


# ============================================================
# FLASK APP
# ============================================================

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


# ============================================================
# MONTH / LANGUAGE DATA
# ============================================================

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
    "December",
]

MONTHS_BN = {
    "January": "জানুয়ারি",
    "February": "ফেব্রুয়ারি",
    "March": "মার্চ",
    "April": "এপ্রিল",
    "May": "মে",
    "June": "জুন",
    "July": "জুলাই",
    "August": "আগস্ট",
    "September": "সেপ্টেম্বর",
    "October": "অক্টোবর",
    "November": "নভেম্বর",
    "December": "ডিসেম্বর",
}

LANG = {
    "Dashboard": "ড্যাশবোর্ড",
    "Workers": "শ্রমিক",
    "Attendance": "উপস্থিতি",
    "Payslip": "বেতন স্লিপ",
    "Advance Salary": "অগ্রিম বেতন",
    "Department": "বিভাগ",
    "Settings": "সেটিংস",
    "Users": "ব্যবহারকারী",
    "Logout": "লগআউট",
    "Login": "লগইন",
    "Search": "অনুসন্ধান",
    "Add Worker": "শ্রমিক যোগ করুন",
    "Edit": "সম্পাদনা",
    "Delete": "মুছে ফেলুন",
    "Save": "সংরক্ষণ",
    "Cancel": "বাতিল",
    "Name": "নাম",
    "Department": "বিভাগ",
    "Designation": "পদবি",
    "Basic Salary": "মূল বেতন",
    "OT Rate": "ওটি রেট",
    "Present": "উপস্থিত",
    "Absent": "অনুপস্থিত",
    "OT Hours": "ওটি ঘণ্টা",
    "Net Salary": "নিট বেতন",
    "Gross Salary": "মোট বেতন",
    "Advance": "অগ্রিম",
    "Refreshment Bill": "নাস্তা বিল",
}

DEPT_BN = {
    "Dyeing": "ডাইং",
    "Printing": "প্রিন্টিং",
    "Finishing": "ফিনিশিং",
    "Production": "প্রোডাকশন",
    "Maintenance": "মেইনটেন্যান্স",
    "Electrical": "ইলেকট্রিক্যাল",
    "Store": "স্টোর",
    "Accounts": "অ্যাকাউন্টস",
    "HR": "এইচআর",
    "Security": "সিকিউরিটি",
    "Admin": "অ্যাডমিন",
}


# ============================================================
# DATABASE CONNECTION
# ============================================================

def is_postgres():
    return bool(DATABASE_URL)


def db_connect():
    """
    PostgreSQL on Render.
    SQLite locally.
    """

    if is_postgres():
        try:
            import psycopg2
            from psycopg2.extras import RealDictCursor

            conn = psycopg2.connect(
                DATABASE_URL,
                sslmode="require",
                cursor_factory=RealDictCursor,
            )
            return conn

        except Exception:
            raise

    conn = sqlite3.connect(
        DB_PATH,
        timeout=30,
        check_same_thread=False,
    )

    conn.row_factory = sqlite3.Row

    return conn


def placeholders(sql):
    """
    Convert SQLite ? placeholders to PostgreSQL %s.
    """

    if is_postgres():
        return sql.replace("?", "%s")

    return sql


def row_dict(row):
    if row is None:
        return None

    if isinstance(row, dict):
        return dict(row)

    try:
        return dict(row)
    except Exception:
        return row


def fetch_all(sql, params=()):
    conn = db_connect()

    try:
        cur = conn.cursor()

        cur.execute(
            placeholders(sql),
            tuple(params)
        )

        rows = cur.fetchall()

        return [
            row_dict(r)
            for r in rows
        ]

    finally:
        conn.close()


def fetch_one(sql, params=()):
    conn = db_connect()

    try:
        cur = conn.cursor()

        cur.execute(
            placeholders(sql),
            tuple(params)
        )

        row = cur.fetchone()

        return row_dict(row)

    finally:
        conn.close()


def scalar(sql, params=(), default=None):
    row = fetch_one(sql, params)

    if row is None:
        return default

    try:
        return next(iter(row.values()))
    except Exception:
        return default


def execute(sql, params=(), commit=False):
    conn = db_connect()

    try:
        cur = conn.cursor()

        cur.execute(
            placeholders(sql),
            tuple(params)
        )

        if commit:
            conn.commit()

        try:
            return cur.lastrowid
        except Exception:
            return None

    finally:
        conn.close()


# ============================================================
# DATABASE HELPERS
# ============================================================

def table_exists(table_name):
    try:
        if is_postgres():
            row = fetch_one(
                """
                SELECT EXISTS(
                    SELECT 1
                    FROM information_schema.tables
                    WHERE table_schema='public'
                    AND table_name=?
                ) AS exists
                """,
                (table_name,),
            )
            return bool(row and row.get("exists"))

        row = fetch_one(
            """
            SELECT name
            FROM sqlite_master
            WHERE type='table'
            AND name=?
            """,
            (table_name,),
        )

        return row is not None

    except Exception:
        return False


def columns(table_name):
    if not table_exists(table_name):
        return set()

    try:
        if is_postgres():
            rows = fetch_all(
                """
                SELECT column_name
                FROM information_schema.columns
                WHERE table_schema='public'
                AND table_name=?
                """,
                (table_name,),
            )

            return {
                str(r["column_name"])
                for r in rows
            }

        rows = fetch_all(
            f"PRAGMA table_info({table_name})"
        )

        return {
            str(r["name"])
            for r in rows
        }

    except Exception:
        return set()


def add_column(table_name, column_name, column_type):
    """
    Add only missing columns.
    Never delete or recreate legacy data.
    """

    if column_name in columns(table_name):
        return

    try:
        execute(
            f"""
            ALTER TABLE {table_name}
            ADD COLUMN {column_name} {column_type}
            """,
            commit=True,
        )
    except Exception as e:
        app.logger.warning(
            "Could not add column %s.%s: %r",
            table_name,
            column_name,
            e,
        )


def hash_password(password):
    return hashlib.sha256(
        str(password).encode("utf-8")
    ).hexdigest()


def nowstr():
    return datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_db():
    """
    Initialize only missing tables/columns.

    IMPORTANT:
    Existing migrated data is NEVER dropped,
    truncated, or re-imported.
    """

    if is_postgres():

        execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE,
                password_hash TEXT,
                password TEXT,
                full_name TEXT,
                role TEXT DEFAULT 'Operator',
                active INTEGER DEFAULT 1,
                last_login TEXT
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS company_settings (
                id SERIAL PRIMARY KEY,
                key TEXT UNIQUE,
                value TEXT
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS workers (
                id SERIAL PRIMARY KEY,
                name TEXT,
                basic_salary REAL DEFAULT 0,
                ot_rate REAL DEFAULT 0,
                department TEXT,
                designation TEXT,
                refreshment_bill REAL DEFAULT 0,
                bangla_name TEXT
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS attendance (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT,
                present_days INTEGER DEFAULT 0,
                absent_days INTEGER DEFAULT 0,
                ot_hours REAL DEFAULT 0
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS daily_attendance (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT,
                day INTEGER,
                status TEXT DEFAULT 'P'
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS activity_log (
                id SERIAL PRIMARY KEY,
                username TEXT,
                action TEXT,
                created_at TEXT
            )
            """,
            commit=True,
        )

    else:

        execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE,
                password_hash TEXT,
                password TEXT,
                full_name TEXT,
                role TEXT DEFAULT 'Operator',
                active INTEGER DEFAULT 1,
                last_login TEXT
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS company_settings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                key TEXT UNIQUE,
                value TEXT
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS workers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                basic_salary REAL DEFAULT 0,
                ot_rate REAL DEFAULT 0,
                department TEXT,
                designation TEXT,
                refreshment_bill REAL DEFAULT 0,
                bangla_name TEXT
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER,
                month_year TEXT,
                present_days INTEGER DEFAULT 0,
                absent_days INTEGER DEFAULT 0,
                ot_hours REAL DEFAULT 0
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS daily_attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER,
                month_year TEXT,
                day INTEGER,
                status TEXT DEFAULT 'P'
            )
            """,
            commit=True,
        )

        execute(
            """
            CREATE TABLE IF NOT EXISTS activity_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                action TEXT,
                created_at TEXT
            )
            """,
            commit=True,
        )

    # --------------------------------------------------------
    # Add missing columns only
    # --------------------------------------------------------

    if table_exists("users"):
        add_column("users", "password_hash", "TEXT")
        add_column("users", "password", "TEXT")
        add_column("users", "full_name", "TEXT")
        add_column("users", "role", "TEXT")
        add_column("users", "active", "INTEGER")
        add_column("users", "last_login", "TEXT")

    if table_exists("workers"):
        add_column("workers", "bangla_name", "TEXT")
        add_column("workers", "refreshment_bill", "REAL")
        add_column("workers", "designation", "TEXT")
        add_column("workers", "department", "TEXT")
        add_column("workers", "ot_rate", "REAL")
        add_column("workers", "basic_salary", "REAL")

    if table_exists("attendance"):
        add_column("attendance", "worker_id", "INTEGER")
        add_column("attendance", "month_year", "TEXT")
        add_column("attendance", "present_days", "INTEGER")
        add_column("attendance", "absent_days", "INTEGER")
        add_column("attendance", "ot_hours", "REAL")

    if table_exists("daily_attendance"):
        add_column("daily_attendance", "worker_id", "INTEGER")
        add_column("daily_attendance", "month_year", "TEXT")
        add_column("daily_attendance", "day", "INTEGER")
        add_column("daily_attendance", "status", "TEXT")

    # --------------------------------------------------------
    # Company setting
    # --------------------------------------------------------

    try:
        existing = fetch_one(
            """
            SELECT *
            FROM company_settings
            WHERE key=?
            LIMIT 1
            """,
            ("company_name",),
        )

        if not existing:
            execute(
                """
                INSERT INTO company_settings(key,value)
                VALUES(?,?)
                """,
                (
                    "company_name",
                    DEFAULT_COMPANY_NAME,
                ),
                commit=True,
            )

    except Exception:
        pass

    # --------------------------------------------------------
    # Create default admin only if no user exists
    # --------------------------------------------------------

    try:
        count = scalar(
            "SELECT COUNT(*) FROM users",
            default=0,
        )

        if int(count or 0) == 0:

            execute(
                """
                INSERT INTO users
                (username,password_hash,full_name,role,active)
                VALUES(?,?,?,?,?)
                """,
                (
                    "admin",
                    hash_password("admin123"),
                    "Administrator",
                    "Administrator",
                    1,
                ),
                commit=True,
            )

    except Exception as e:
        app.logger.warning(
            "Admin initialization warning: %r",
            e,
        )


# ============================================================
# SETTINGS / USER
# ============================================================

def get_settings():
    settings = {
        "company_name": DEFAULT_COMPANY_NAME
    }

    try:
        rows = fetch_all(
            "SELECT key,value FROM company_settings"
        )

        for r in rows:
            settings[
                str(r.get("key"))
            ] = r.get("value")

    except Exception:
        pass

    return settings


def current_user():
    uid = session.get("user_id")

    if not uid:
        return None

    try:
        return fetch_one(
            "SELECT * FROM users WHERE id=?",
            (uid,),
        )
    except Exception:
        return None


def log_activity(username, action):
    try:
        execute(
            """
            INSERT INTO activity_log
            (username,action,created_at)
            VALUES(?,?,?)
            """,
            (
                username,
                action,
                nowstr(),
            ),
            commit=True,
        )
    except Exception:
        pass


# ============================================================
# AUTH DECORATORS
# ============================================================

def login_required(f):

    @wraps(f)
    def wrapper(*args, **kwargs):

        if not session.get("user_id"):
            return redirect(
                url_for(
                    "login",
                    next=request.path,
                )
            )

        return f(*args, **kwargs)

    return wrapper


def admin_required(f):

    @wraps(f)
    def wrapper(*args, **kwargs):

        u = current_user()

        if not u or str(
            u.get("role", "")
        ).lower() not in (
            "administrator",
            "admin",
        ):

            flash(
                "Administrator access required.",
                "danger",
            )

            return redirect(
                url_for("dashboard")
            )

        return f(*args, **kwargs)

    return wrapper


# ============================================================
# GENERAL HELPERS
# ============================================================

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


def parse_num(v, default=0.0):

    try:
        return float(v or default)

    except Exception:
        return default


# ============================================================
# ADVANCE SALARY HELPERS
# ============================================================

def legacy_advance_source():

    if table_exists("worker_advances"):

        c = columns("worker_advances")

        if {
            "worker_id",
            "amount",
        }.issubset(c):

            return "worker_advances"

    if table_exists("advance_salary"):
        return "advance_salary"

    if table_exists("advances"):
        return "advances"

    return None


def month_name_from_date(s):

    try:

        d = str(s)[:10]

        y, m, _ = d.split("-")

        return (
            f"{MONTHS[int(m) - 1]} "
            f"{int(y)}"
        )

    except Exception:

        return ""


def advance_rows(month=None, worker_id=None):

    t = legacy_advance_source()

    if not t:
        return []

    c = columns(t)

    # --------------------------------------------------------
    # New worker_advances table
    # --------------------------------------------------------

    if t == "worker_advances":

        where = []
        p = []

        if month:

            where.append(
                "month_year=?"
            )

            p.append(month)

        if worker_id:

            where.append(
                "worker_id=?"
            )

            p.append(worker_id)

        q = """
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

            q += (
                " WHERE "
                + " AND ".join(where)
            )

        q += " ORDER BY id DESC"

        return fetch_all(q, p)

    # --------------------------------------------------------
    # Legacy advance_salary
    # --------------------------------------------------------

    if t == "advance_salary":

        datecol = (
            "date"
            if "date" in c
            else (
                "advance_date"
                if "advance_date" in c
                else None
            )
        )

        q = f"""
            SELECT
                id,
                worker_id,
                {datecol or 'NULL'} AS advance_date,
                amount,
                note
            FROM {t}
        """

        rows = fetch_all(q)

        for r in rows:

            r["month_year"] = (
                month_name_from_date(
                    r.get("advance_date")
                )
            )

        if month:

            rows = [
                r
                for r in rows
                if r.get("month_year") == month
            ]

        if worker_id:

            rows = [
                r
                for r in rows
                if str(
                    r.get("worker_id")
                ) == str(worker_id)
            ]

        return rows

    # --------------------------------------------------------
    # Legacy advances
    # --------------------------------------------------------

    q = f"""
        SELECT
            id,
            worker_id,
            date AS advance_date,
            amount,
            note
        FROM {t}
    """

    rows = fetch_all(q)

    for r in rows:

        r["month_year"] = (
            month_name_from_date(
                r.get("advance_date")
            )
        )

    if month:

        rows = [
            r
            for r in rows
            if r.get("month_year") == month
        ]

    if worker_id:

        rows = [
            r
            for r in rows
            if str(
                r.get("worker_id")
            ) == str(worker_id)
        ]

    return rows


def save_advance_record(
    worker_id,
    month,
    adv_date,
    amount,
    note,
):

    t = legacy_advance_source()

    if t == "worker_advances":

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
                note,
            ),
            commit=True,
        )

    elif t == "advance_salary":

        c = columns(t)

        datecol = (
            "date"
            if "date" in c
            else "advance_date"
        )

        execute(
            f"""
            INSERT INTO {t}
            (
                worker_id,
                {datecol},
                amount,
                note
            )
            VALUES(?,?,?,?)
            """,
            (
                worker_id,
                adv_date,
                amount,
                note,
            ),
            commit=True,
        )

    elif t == "advances":

        execute(
            """
            INSERT INTO advances
            (
                worker_id,
                date,
                amount,
                note
            )
            VALUES(?,?,?,?)
            """,
            (
                str(worker_id),
                adv_date,
                amount,
                note,
            ),
            commit=True,
        )

    else:

        # Create a new table only when no advance table exists.
        # This does not touch migrated tables.

        if is_postgres():

            execute(
                """
                CREATE TABLE IF NOT EXISTS
                worker_advances (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    advance_date TEXT NOT NULL,
                    amount REAL NOT NULL,
                    note TEXT
                )
                """,
                commit=True,
            )

        else:

            execute(
                """
                CREATE TABLE IF NOT EXISTS
                worker_advances (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    advance_date TEXT NOT NULL,
                    amount REAL NOT NULL,
                    note TEXT
                )
                """,
                commit=True,
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
                note,
            ),
            commit=True,
        )


def delete_advance_record(aid):

    t = legacy_advance_source()

    if t:

        execute(
            f"""
            DELETE FROM {t}
            WHERE id=?
            """,
            (aid,),
            commit=True,
        )


# ============================================================
# ATTENDANCE / SALARY
# ============================================================

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
            month,
        ),
    )

    return {
        int(r["day"]):
        str(r["status"] or "P")
        for r in rows
    }


def calculate_salary_bulk(workers, month):
    """
    Calculate salary for multiple workers.

    Formula:
        absent cut = basic / month days * absent days
        earned basic = basic - absent cut
        OT = OT hours * OT rate
        Nasta = eligible present days * refreshment bill
        gross = earned basic + OT + Nasta
        net = gross - advance
    """

    ids = [
        w["id"]
        for w in workers
    ]

    if not ids:
        return {}

    # --------------------------------------------------------
    # Monthly attendance
    # --------------------------------------------------------

    att_rows = fetch_all(
        """
        SELECT *
        FROM attendance
        WHERE month_year=?
        ORDER BY id DESC
        """,
        (month,),
    )

    att_by = {}

    for r in att_rows:

        wid = r.get("worker_id")

        if wid not in att_by:
            att_by[wid] = r

    # --------------------------------------------------------
    # Daily attendance
    # --------------------------------------------------------

    daily_rows = fetch_all(
        """
        SELECT
            worker_id,
            day,
            status
        FROM daily_attendance
        WHERE month_year=?
        """,
        (month,),
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

    # --------------------------------------------------------
    # Advances
    # --------------------------------------------------------

    advances_by = {}

    t = legacy_advance_source()

    if t:

        c = columns(t)

        try:

            if (
                t == "worker_advances"
                and {
                    "worker_id",
                    "amount",
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
                    (month,),
                )

                for r in adv_rows:

                    wid = r.get("worker_id")

                    advances_by[wid] = (
                        advances_by.get(
                            wid,
                            0
                        )
                        + parse_num(
                            r.get("amount")
                        )
                    )

            elif (
                t in (
                    "advance_salary",
                    "advances",
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
                        FROM {t}
                        """
                    )

                    for r in adv_rows:

                        if (
                            month_name_from_date(
                                r.get(
                                    "advance_date"
                                )
                            )
                            == month
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
                                    r.get("amount")
                                )
                            )

        except Exception as e:

            app.logger.warning(
                "Bulk advance load error: %r",
                e,
            )

    # --------------------------------------------------------
    # Calculate
    # --------------------------------------------------------

    out = {}

    for w in workers:

        att = att_by.get(
            w["id"],
            {}
        )

        present = int(
            att.get("present_days")
            or 0
        )

        absent = int(
            att.get("absent_days")
            or 0
        )

        ot = parse_num(
            att.get("ot_hours"),
            0,
        )

        try:

            mon, ys = month.split()

            y = int(ys)

            mi = MONTHS.index(mon) + 1

            days = calendar.monthrange(
                y,
                mi
            )[1]

        except Exception:

            days = 30

            y = datetime.date.today().year

            mi = datetime.date.today().month

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
            (basic / days) * absent
            if days
            else 0
        )

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
                        y,
                        mi,
                        d
                    ).weekday() != 4
                )
            )

            nasta = (
                billable * nasta_rate
            )

        else:

            non_friday = sum(
                1
                for d in range(
                    1,
                    days + 1
                )
                if datetime.date(
                    y,
                    mi,
                    d
                ).weekday() != 4
            )

            nasta = (
                round(
                    present
                    * (
                        non_friday / days
                    )
                    * nasta_rate,
                    2,
                )
                if days
                else 0
            )

        advances = advances_by.get(
            w["id"],
            0
        )

        gross = (
            earned_basic
            + ot_amt
            + nasta
        )

        out[w["id"]] = {
            "present": present,
            "absent": absent,
            "ot": ot,
            "absent_cut": absent_cut,
            "earned_basic": earned_basic,
            "ot_amt": ot_amt,
            "nasta": nasta,
            "gross": gross,
            "advance": advances,
            "net": gross - advances,
        }

    return out


def calculate_salary(
    worker,
    month,
    att=None,
):

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
                month,
            ),
        ) or {}

    present = int(
        att.get("present_days")
        or 0
    )

    absent = int(
        att.get("absent_days")
        or 0
    )

    ot = parse_num(
        att.get("ot_hours"),
        0,
    )

    try:

        mon, ys = month.split()

        y = int(ys)

        mi = MONTHS.index(mon) + 1

        days = calendar.monthrange(
            y,
            mi
        )[1]

    except Exception:

        days = 30

        y = datetime.date.today().year

        mi = datetime.date.today().month

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
        (basic / days) * absent
        if days
        else 0
    )

    earned_basic = max(
        0,
        basic - absent_cut
    )

    ot_amt = (
        ot * ot_rate
    )

    dm = daily_map(
        worker["id"],
        month,
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
                    y,
                    mi,
                    d
                ).weekday() != 4
            )
        )

        nasta = (
            billable * nasta_rate
        )

    else:

        non_friday = sum(
            1
            for d in range(
                1,
                days + 1
            )
            if datetime.date(
                y,
                mi,
                d
            ).weekday() != 4
        )

        nasta = (
            round(
                present
                * (
                    non_friday / days
                )
                * nasta_rate,
                2,
            )
            if days
            else 0
        )

    advances = sum(
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
        "advance": advances,
        "net": gross - advances,
    }


# ============================================================
# TEMPLATE GLOBALS
# ============================================================

@app.context_processor
def inject_globals():

    settings = get_settings()

    lang = session.get(
        "language",
        "en",
    )

    def tr(s):

        return (
            LANG.get(s, s)
            if lang == "bn"
            else s
        )

    def dept(s):

        value = str(s or "")

        return (
            DEPT_BN.get(
                value,
                value,
            )
            if lang == "bn"
            else value
        )

    return {
        "current_user": current_user(),
        "settings": settings,
        "language": lang,
        "tr": tr,

        # IMPORTANT:
        # workers.html and department.html use dept()
        "dept": dept,

        # Keep backward compatibility
        "display_dept": dept,

        "months": MONTHS,
        "years": list(
            range(
                2024,
                2032
            )
        ),
    }


# ============================================================
# HEALTH
# ============================================================

@app.route("/health")
def health():

    try:

        w = scalar(
            "SELECT COUNT(*) FROM workers",
            default=0,
        )

        a = scalar(
            "SELECT COUNT(*) FROM attendance",
            default=0,
        )

        d = scalar(
            "SELECT COUNT(*) FROM daily_attendance",
            default=0,
        )

        return jsonify(
            {
                "status": "ok",
                "database": (
                    "postgresql"
                    if is_postgres()
                    else "sqlite"
                ),
                "workers": w,
                "attendance": a,
                "daily_attendance": d,
            }
        )

    except Exception as e:

        return jsonify(
            {
                "status": "error",
                "error": repr(e),
            }
        ), 500


# ============================================================
# DATABASE CHECK
# ============================================================

@app.route("/db-check")
@login_required
def db_check():

    info = {}

    for t in [
        "workers",
        "attendance",
        "daily_attendance",
        "company_settings",
        "users",
        "activity_log",
        "worker_advances",
        "advance_salary",
        "advances",
    ]:

        try:

            exists = table_exists(t)

            info[t] = {
                "exists": exists,
                "columns": (
                    sorted(columns(t))
                    if exists
                    else []
                ),
            }

        except Exception as e:

            info[t] = {
                "error": repr(e)
            }

    return render_template_string(
        """
        <!doctype html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Database Check</title>
        </head>

        <body>

        <h1>Database Check</h1>

        <pre>{{ info|tojson(indent=2) }}</pre>

        <p>
            <a href="{{ url_for('dashboard') }}">
                Back
            </a>
        </p>

        </body>
        </html>
        """,
        info=info,
    )


# ============================================================
# LOGIN
# ============================================================

@app.route(
    "/login",
    methods=["GET", "POST"]
)
def login():

    if request.method == "POST":

        u = request.form.get(
            "username",
            ""
        ).strip()

        p = request.form.get(
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
            (u,),
        )

        ok = False

        if user and int(
            user.get("active")
            if user.get("active") is not None
            else 1
        ):

            ph = (
                user.get("password_hash")
                or user.get("password")
                or ""
            )

            ok = ph in (
                hash_password(p),
                p,
            )

        if ok:

            session["user_id"] = user["id"]

            session["language"] = session.get(
                "language",
                "en",
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
                        user["id"],
                    ),
                    commit=True,
                )

            except Exception:
                pass

            log_activity(
                u,
                "Logged in"
            )

            return redirect(
                request.args.get(
                    "next"
                )
                or url_for("dashboard")
            )

        flash(
            "Invalid User ID or Password.",
            "danger",
        )

    return render_template(
        "login.html"
    )


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():

    u = current_user()

    if u:

        log_activity(
            u.get("username"),
            "Logged out",
        )

    session.clear()

    return redirect(
        url_for("login")
    )


# ============================================================
# LANGUAGE
# ============================================================

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


# ============================================================
# HOME
# ============================================================

@app.route("/")
@login_required
def index():

    return redirect(
        url_for("dashboard")
    )


# ============================================================
# DASHBOARD
# ============================================================

@app.route("/dashboard")
@login_required
def dashboard():

    month = month_name_year()

    workers = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    salaries = calculate_salary_bulk(
        workers,
        month,
    )

    gross = sum(
        s["gross"]
        for s in salaries.values()
    )

    adv = sum(
        s["advance"]
        for s in salaries.values()
    )

    today = datetime.date.today()

    tm = (
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
            tm,
            today.day,
        ),
    )

    mp = {
        r["status"]: r["c"]
        for r in today_rows
    }

    depts = fetch_all(
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
        workers=workers,
        month=month,
        month_name=month.split()[0],
        year=month.split()[1],
        gross=gross,
        advance=adv,
        today_present=mp.get(
            "P",
            0,
        ),
        today_absent=mp.get(
            "A",
            0,
        ),
        dept_rows=depts,
        total_workers=len(workers),
    )


# ============================================================
# WORKERS
# ============================================================

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

        l = f"%{q}%"

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
            l,
            l,
            l,
            l,
            l,
        ]

    sql += """
        ORDER BY id DESC
    """

    return render_template(
        "workers.html",
        workers=fetch_all(
            sql,
            params,
        ),
        q=q,
    )


# ============================================================
# WORKER FORM
# ============================================================

WORKER_FORM_HTML = """
<!doctype html>

<html lang="en">

<head>

<meta charset="utf-8">

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Worker</title>

<style>

body{
    font:16px Arial;
    background:#f3f6fb;
    margin:30px;
}

form{
    max-width:650px;
    background:white;
    padding:24px;
    border-radius:12px;
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:12px;
}

label{
    display:flex;
    flex-direction:column;
    gap:5px;
}

input{
    padding:10px;
    border:1px solid #ccd5e1;
    border-radius:6px;
}

button{
    padding:12px;
    background:#1764c0;
    color:white;
    border:0;
    border-radius:6px;
    cursor:pointer;
}

.wide{
    grid-column:1/-1;
}

</style>

</head>

<body>

<h2>
{{ 'Edit Worker' if worker else 'Add Worker' }}
</h2>

<form method="post">

<label>
Name

<input
name="name"
required
value="{{ worker.get('name','') if worker else '' }}"
>

</label>


<label>
Bangla Name

<input
name="bangla_name"
value="{{ worker.get('bangla_name','') if worker else '' }}"
>

</label>


<label>
Basic Salary

<input
type="number"
step="any"
name="basic_salary"
value="{{ worker.get('basic_salary',0) if worker else 0 }}"
>

</label>


<label>
OT Rate

<input
type="number"
step="any"
name="ot_rate"
value="{{ worker.get('ot_rate',0) if worker else 0 }}"
>

</label>


<label>
Department

<input
name="department"
value="{{ worker.get('department','') if worker else '' }}"
>

</label>


<label>
Designation

<input
name="designation"
value="{{ worker.get('designation','') if worker else '' }}"
>

</label>


<label>
Refreshment/Nasta Bill

<input
type="number"
step="any"
name="refreshment_bill"
value="{{ worker.get('refreshment_bill',0) if worker else 0 }}"
>

</label>


<label>
Phone

<input
name="phone"
value="{{ worker.get('phone','') if worker else '' }}"
>

</label>


<button
class="wide"
type="submit"
>
Save
</button>


<a
class="wide"
href="{{ url_for('workers') }}"
>
Cancel / Back to Workers
</a>

</form>

</body>

</html>
"""


# ============================================================
# USER FORM
# ============================================================

USER_FORM_HTML = """
<!doctype html>

<html>

<head>

<meta charset="utf-8">

<title>User</title>

</head>

<body>

<h2>Add User</h2>

<form method="post">

<p>
Username

<input
name="username"
required
value="{{ user.get('username','') if user else '' }}"
>
</p>

<p>
Full name

<input
name="full_name"
value="{{ user.get('full_name','') if user else '' }}"
>
</p>

<p>
Password

<input
type="password"
name="password"
required
>
</p>

<p>

Role

<select name="role">

<option>
Operator
</option>

<option>
Administrator
</option>

</select>

</p>

<button type="submit">
Create
</button>

</form>

<p>

<a href="{{ url_for('users') }}">
Back
</a>

</p>

</body>

</html>
"""


# ============================================================
# ADD WORKER
# ============================================================

@app.route(
    "/workers/add",
    methods=["GET", "POST"]
)
@login_required
def add_worker():

    if request.method == "POST":

        f = request.form

        name = f.get(
            "name",
            ""
        ).strip()

        if not name:

            flash(
                "Worker name is required.",
                "danger",
            )

            return render_template_string(
                WORKER_FORM_HTML,
                worker=f,
            )

        vals = (
            name,
            parse_num(
                f.get("basic_salary")
            ),
            parse_num(
                f.get("ot_rate")
            ),
            f.get(
                "department",
                ""
            ).strip(),
            f.get(
                "designation",
                ""
            ).strip(),
            parse_num(
                f.get("refreshment_bill")
            ),
            f.get(
                "bangla_name",
                ""
            ).strip(),
        )

        execute(
            """
            INSERT INTO workers
            (
                name,
                basic_salary,
                ot_rate,
                department,
                designation,
                refreshment_bill,
                bangla_name
            )
            VALUES(?,?,?,?,?,?,?)
            """,
            vals,
            commit=True,
        )

        flash(
            "Worker saved.",
            "success",
        )

        return redirect(
            url_for("workers")
        )

    return render_template_string(
        WORKER_FORM_HTML,
        worker=None,
    )


# ============================================================
# EDIT WORKER
# ============================================================

@app.route(
    "/workers/edit/<int:worker_id>",
    methods=["GET", "POST"]
)
@login_required
def edit_worker(worker_id):

    w = fetch_one(
        """
        SELECT *
        FROM workers
        WHERE id=?
        """,
        (worker_id,),
    )

    if not w:
        abort(404)

    if request.method == "POST":

        f = request.form

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
                f.get("name", ""),
                parse_num(
                    f.get("basic_salary")
                ),
                parse_num(
                    f.get("ot_rate")
                ),
                f.get(
                    "department",
                    ""
                ),
                f.get(
                    "designation",
                    ""
                ),
                parse_num(
                    f.get(
                        "refreshment_bill"
                    )
                ),
                f.get(
                    "bangla_name",
                    ""
                ),
                worker_id,
            ),
            commit=True,
        )

        flash(
            "Worker updated.",
            "success",
        )

        return redirect(
            url_for("workers")
        )

    return render_template_string(
        WORKER_FORM_HTML,
        worker=w,
    )


# ============================================================
# DELETE WORKER
# ============================================================

@app.route(
    "/workers/delete/<int:worker_id>",
    methods=["POST", "GET"]
)
@login_required
def delete_worker(worker_id):

    w = fetch_one(
        """
        SELECT *
        FROM workers
        WHERE id=?
        """,
        (worker_id,),
    )

    if not w:
        abort(404)

    execute(
        """
        DELETE FROM workers
        WHERE id=?
        """,
        (worker_id,),
        commit=True,
    )

    flash(
        "Worker deleted.",
        "success",
    )

    return redirect(
        url_for("workers")
    )


# ============================================================
# ATTENDANCE
# ============================================================

@app.route(
    "/attendance",
    methods=["GET", "POST"]
)
@login_required
def attendance():

    month = month_name_year()

    workers = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    if request.method == "POST":

        month = month_name_year(
            request.form.get("month"),
            request.form.get("year"),
        )

        for w in workers:

            wid = w["id"]

            present = int(
                parse_num(
                    request.form.get(
                        f"present_{wid}",
                        0,
                    )
                )
            )

            absent = int(
                parse_num(
                    request.form.get(
                        f"absent_{wid}",
                        0,
                    )
                )
            )

            ot = parse_num(
                request.form.get(
                    f"ot_{wid}",
                    0,
                )
            )

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
                    wid,
                    month,
                ),
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
                        existing["id"],
                    ),
                    commit=True,
                )

            else:

                execute(
                    """
                    INSERT INTO attendance
                    (
                        worker_id,
                        month_year,
                        present_days,
                        absent_days,
                        ot_hours
                    )
                    VALUES(?,?,?,?,?)
                    """,
                    (
                        wid,
                        month,
                        present,
                        absent,
                        ot,
                    ),
                    commit=True,
                )

        flash(
            "Attendance saved successfully.",
            "success",
        )

        return redirect(
            url_for(
                "attendance",
                month=month.split()[0],
                year=month.split()[1],
            )
        )

    att_rows = fetch_all(
        """
        SELECT *
        FROM attendance
        WHERE month_year=?
        ORDER BY worker_id
        """,
        (month,),
    )

    att_map = {}

    for r in att_rows:

        att_map[
            r["worker_id"]
        ] = r

    return render_template(
        "attendance.html",
        workers=workers,
        month=month,
        att_map=att_map,
    )


# ============================================================
# DAILY ATTENDANCE
# ============================================================

@app.route(
    "/daily-attendance",
    methods=["GET", "POST"]
)
@login_required
def daily_attendance():

    month = month_name_year()

    workers = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    try:

        mon, ys = month.split()

        year = int(ys)

        month_num = (
            MONTHS.index(mon) + 1
        )

        days = calendar.monthrange(
            year,
            month_num,
        )[1]

    except Exception:

        today = datetime.date.today()

        year = today.year

        month_num = today.month

        days = calendar.monthrange(
            year,
            month_num,
        )[1]

    if request.method == "POST":

        month = month_name_year(
            request.form.get("month"),
            request.form.get("year"),
        )

        for w in workers:

            wid = w["id"]

            for day in range(
                1,
                days + 1,
            ):

                status = request.form.get(
                    f"status_{wid}_{day}"
                )

                if not status:
                    continue

                existing = fetch_one(
                    """
                    SELECT id
                    FROM daily_attendance
                    WHERE worker_id=?
                    AND month_year=?
                    AND day=?
                    LIMIT 1
                    """,
                    (
                        wid,
                        month,
                        day,
                    ),
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
                            existing["id"],
                        ),
                        commit=True,
                    )

                else:

                    execute(
                        """
                        INSERT INTO daily_attendance
                        (
                            worker_id,
                            month_year,
                            day,
                            status
                        )
                        VALUES(?,?,?,?)
                        """,
                        (
                            wid,
                            month,
                            day,
                            status,
                        ),
                        commit=True,
                    )

        flash(
            "Daily attendance saved.",
            "success",
        )

        return redirect(
            url_for(
                "daily_attendance"
            )
        )

    rows = fetch_all(
        """
        SELECT *
        FROM daily_attendance
        WHERE month_year=?
        ORDER BY worker_id,day
        """,
        (month,),
    )

    daily = {}

    for r in rows:

        daily.setdefault(
            r["worker_id"],
            {}
        )[int(r["day"])] = (
            r["status"]
        )

    return render_template(
        "attendance.html",
        workers=workers,
        month=month,
        daily=daily,
        days=days,
        daily_mode=True,
    )


# ============================================================
# PAYSLIP
# ============================================================

@app.route("/payslip")
@login_required
def payslip():

    month = month_name_year()

    wid = request.args.get(
        "worker_id"
    )

    # Compatibility with old links
    if not wid:
        wid = request.args.get(
            "wid"
        )

    w = (
        fetch_one(
            """
            SELECT *
            FROM workers
            WHERE id=?
            """,
            (wid,),
        )
        if wid
        else None
    )

    salary = (
        calculate_salary(
            w,
            month,
        )
        if w
        else {}
    )

    return render_template(
        "payslip.html",
        workers=fetch_all(
            """
            SELECT
                id,
                name,
                department
            FROM workers
            ORDER BY id
            """
        ),
        worker=w,
        summary=salary,
        month=month,
    )


# ============================================================
# DEPARTMENT
# ============================================================

@app.route("/department")
@login_required
def department():

    month = month_name_year()

    # IMPORTANT:
    # Do NOT call this variable "dept".
    # The template uses dept() as a function.
    selected_dept = request.args.get(
        "department",
        "",
    ).strip()

    if selected_dept:

        ws = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (selected_dept,),
        )

    else:

        ws = fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    salary_map = calculate_salary_bulk(
        ws,
        month,
    )

    rows = [
        {
            **w,
            **salary_map.get(
                w["id"],
                {}
            ),
        }
        for w in ws
    ]

    departments = [
        r["department"]
        for r in fetch_all(
            """
            SELECT DISTINCT department
            FROM workers
            WHERE department IS NOT NULL
            AND department <> ''
            ORDER BY department
            """
        )
    ]

    return render_template(
        "department.html",
        rows=rows,
        month=month,
        department=selected_dept,
        selected_department=selected_dept,
        departments=departments,
    )


# ============================================================
# ADVANCE
# ============================================================

@app.route("/advance")
@app.route("/advances")
@login_required
def advance():

    month = month_name_year()

    wid = request.args.get(
        "worker_id"
    )

    return render_template(
        "advances.html",
        workers=fetch_all(
            """
            SELECT
                id,
                name,
                department
            FROM workers
            ORDER BY id
            """
        ),
        rows=advance_rows(
            month,
            wid,
        ),
        month=month,
    )


# ============================================================
# ADD ADVANCE
# ============================================================

@app.route(
    "/advance/add",
    methods=["POST"]
)
@login_required
def add_advance():

    worker_id = request.form.get(
        "worker_id"
    )

    month = month_name_year(
        request.form.get("month"),
        request.form.get("year"),
    )

    adv_date = request.form.get(
        "advance_date"
    ) or datetime.date.today().isoformat()

    amount = parse_num(
        request.form.get("amount")
    )

    note = request.form.get(
        "note",
        "",
    ).strip()

    if not worker_id:

        flash(
            "Please select a worker.",
            "danger",
        )

        return redirect(
            url_for("advance")
        )

    if amount <= 0:

        flash(
            "Advance amount must be greater than zero.",
            "danger",
        )

        return redirect(
            url_for("advance")
        )

    save_advance_record(
        worker_id,
        month,
        adv_date,
        amount,
        note,
    )

    flash(
        "Advance salary saved.",
        "success",
    )

    return redirect(
        url_for(
            "advance",
            worker_id=worker_id,
        )
    )


# ============================================================
# DELETE ADVANCE
# ============================================================

@app.route(
    "/advance/delete/<int:aid>",
    methods=["POST", "GET"]
)
@login_required
def delete_advance(aid):

    delete_advance_record(aid)

    flash(
        "Advance deleted.",
        "success",
    )

    return redirect(
        url_for("advance")
    )


# ============================================================
# SETTINGS
# ============================================================

@app.route(
    "/settings",
    methods=["GET", "POST"]
)
@login_required
def settings():

    if request.method == "POST":

        company_name = request.form.get(
            "company_name",
            DEFAULT_COMPANY_NAME,
        ).strip()

        if not company_name:
            company_name = DEFAULT_COMPANY_NAME

        existing = fetch_one(
            """
            SELECT *
            FROM company_settings
            WHERE key=?
            LIMIT 1
            """,
            ("company_name",),
        )

        if existing:

            execute(
                """
                UPDATE company_settings
                SET value=?
                WHERE key=?
                """,
                (
                    company_name,
                    "company_name",
                ),
                commit=True,
            )

        else:

            execute(
                """
                INSERT INTO company_settings
                (key,value)
                VALUES(?,?)
                """,
                (
                    "company_name",
                    company_name,
                ),
                commit=True,
            )

        flash(
            "Settings saved.",
            "success",
        )

        return redirect(
            url_for("settings")
        )

    return render_template(
        "settings.html",
        settings=get_settings(),
    )


# ============================================================
# USERS
# ============================================================

@app.route("/users")
@login_required
@admin_required
def users():

    return render_template(
        "users.html",
        users=fetch_all(
            """
            SELECT *
            FROM users
            ORDER BY id
            """
        ),
    )


@app.route(
    "/users/add",
    methods=["GET", "POST"]
)
@login_required
@admin_required
def add_user():

    if request.method == "POST":

        f = request.form

        username = f.get(
            "username",
            "",
        ).strip()

        password = f.get(
            "password",
            "",
        )

        full_name = f.get(
            "full_name",
            "",
        ).strip()

        role = f.get(
            "role",
            "Operator",
        )

        if not username or not password:

            flash(
                "Username and password are required.",
                "danger",
            )

            return render_template_string(
                USER_FORM_HTML,
                user=f,
            )

        try:

            execute(
                """
                INSERT INTO users
                (
                    username,
                    password_hash,
                    full_name,
                    role,
                    active
                )
                VALUES(?,?,?,?,?)
                """,
                (
                    username,
                    hash_password(
                        password
                    ),
                    full_name,
                    role,
                    1,
                ),
                commit=True,
            )

            flash(
                "User created.",
                "success",
            )

            return redirect(
                url_for("users")
            )

        except Exception as e:

            app.logger.exception(e)

            flash(
                "Could not create user. Username may already exist.",
                "danger",
            )

    return render_template_string(
        USER_FORM_HTML,
        user=None,
    )


# ============================================================
# ACTIVITY
# ============================================================

@app.route("/activity")
@login_required
def activity():

    rows = []

    if table_exists(
        "activity_log"
    ):

        rows = fetch_all(
            """
            SELECT *
            FROM activity_log
            ORDER BY id DESC
            LIMIT 500
            """
        )

    return render_template(
        "activity.html",
        rows=rows,
    )


# ============================================================
# REPORT
# ============================================================

@app.route("/report")
@login_required
def report():

    month = month_name_year()

    workers = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    salaries = calculate_salary_bulk(
        workers,
        month,
    )

    rows = []

    for w in workers:

        rows.append(
            {
                **w,
                **salaries.get(
                    w["id"],
                    {},
                ),
            }
        )

    return render_template(
        "report.html",
        rows=rows,
        month=month,
    )


# ============================================================
# EXPORT WORKERS CSV
# ============================================================

@app.route("/workers/export")
@login_required
def export_workers():

    workers = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    output = io.StringIO()

    writer = csv.writer(
        output
    )

    writer.writerow(
        [
            "ID",
            "Name",
            "Bangla Name",
            "Department",
            "Designation",
            "Basic Salary",
            "OT Rate",
            "Refreshment Bill",
        ]
    )

    for w in workers:

        writer.writerow(
            [
                w.get("id"),
                w.get("name"),
                w.get("bangla_name"),
                w.get("department"),
                w.get("designation"),
                w.get("basic_salary"),
                w.get("ot_rate"),
                w.get("refreshment_bill"),
            ]
        )

    data = io.BytesIO(
        output.getvalue().encode(
            "utf-8-sig"
        )
    )

    return send_file(
        data,
        mimetype="text/csv",
        as_attachment=True,
        download_name="workers.csv",
    )


# ============================================================
# EXPORT DEPARTMENT CSV
# ============================================================

@app.route("/department/export")
@login_required
def export_department():

    month = month_name_year()

    selected_dept = request.args.get(
        "department",
        "",
    ).strip()

    if selected_dept:

        workers = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (selected_dept,),
        )

    else:

        workers = fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    salary_map = calculate_salary_bulk(
        workers,
        month,
    )

    output = io.StringIO()

    writer = csv.writer(
        output
    )

    writer.writerow(
        [
            "ID",
            "Name",
            "Department",
            "Present",
            "Absent",
            "OT Hours",
            "Basic Earned",
            "OT Amount",
            "Nasta",
            "Gross",
            "Advance",
            "Net",
        ]
    )

    for w in workers:

        s = salary_map.get(
            w["id"],
            {},
        )

        writer.writerow(
            [
                w.get("id"),
                w.get("name"),
                w.get("department"),
                s.get("present", 0),
                s.get("absent", 0),
                s.get("ot", 0),
                s.get("earned_basic", 0),
                s.get("ot_amt", 0),
                s.get("nasta", 0),
                s.get("gross", 0),
                s.get("advance", 0),
                s.get("net", 0),
            ]
        )

    data = io.BytesIO(
        output.getvalue().encode(
            "utf-8-sig"
        )
    )

    return send_file(
        data,
        mimetype="text/csv",
        as_attachment=True,
        download_name="department_salary.csv",
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

@app.route("/export/excel")
@login_required
def export_excel():

    if openpyxl is None:

        return (
            "openpyxl is not installed.",
            500,
        )

    month = month_name_year()

    workers = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    salaries = calculate_salary_bulk(
        workers,
        month,
    )

    wb = Workbook()

    ws = wb.active

    ws.title = "Salary"

    ws.append(
        [
            "ID",
            "Name",
            "Department",
            "Present",
            "Absent",
            "OT Hours",
            "Basic Earned",
            "OT Amount",
            "Nasta",
            "Gross",
            "Advance",
            "Net",
        ]
    )

    for w in workers:

        s = salaries.get(
            w["id"],
            {},
        )

        ws.append(
            [
                w.get("id"),
                w.get("name"),
                w.get("department"),
                s.get("present", 0),
                s.get("absent", 0),
                s.get("ot", 0),
                s.get("earned_basic", 0),
                s.get("ot_amt", 0),
                s.get("nasta", 0),
                s.get("gross", 0),
                s.get("advance", 0),
                s.get("net", 0),
            ]
        )

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name="salary.xlsx",
    )


# ============================================================
# ERROR HANDLERS
# ============================================================

@app.errorhandler(404)
def page_not_found(e):

    try:

        return render_template(
            "404.html"
        ), 404

    except Exception:

        return (
            "404 - Page Not Found",
            404,
        )


@app.errorhandler(500)
def server_error(e):

    app.logger.exception(
        "Unhandled application error"
    )

    try:

        return render_template(
            "500.html",
            error=e,
        ), 500

    except Exception:

        return (
            f"500 - Internal Server Error\n{e}",
            500,
        )


# ============================================================
# INITIALIZE DATABASE
# ============================================================

try:

    init_db()

except Exception as e:

    app.logger.exception(
        "Database initialization error: %r",
        e,
    )


# ============================================================
# LOCAL RUN
# ============================================================

if __name__ == "__main__":

    app.run(
        host="0.0.0.0",
        port=int(
            os.environ.get(
                "PORT",
                5000,
            )
        ),
        debug=False,
    )
