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

try:
    from reportlab.lib.pagesizes import A4
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


# ============================================================
# APP CONFIG
# ============================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "CHANGE-ME-IN-RENDER"
)

DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    ""
).strip()

DB_PATH = os.environ.get(
    "SQLITE_DB",
    "reedoy_payroll.db"
)

DEFAULT_COMPANY_NAME = (
    "REEDOY TEXTILE DYEING PRINTING & FINISHING"
)


# ============================================================
# MONTHS
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


# ============================================================
# DEPARTMENT BANGLA
# ============================================================

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
    "Utility": "ইউটিলিটি",
}


# ============================================================
# TRANSLATIONS
# ============================================================

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
    "Export PDF": "PDF রপ্তানি",
    "Workers": "কর্মী",
    "Total Basic Salary": "মোট মূল বেতন",
    "No department data available": "কোনো বিভাগীয় তথ্য নেই",
}


# ============================================================
# POSTGRES CONNECTION POOL
# ============================================================

_PG_POOL = None


def is_postgres():
    return bool(
        DATABASE_URL
        and not DATABASE_URL.startswith("sqlite://")
    )


def _get_pg_pool():

    global _PG_POOL

    if _PG_POOL is None:

        from psycopg2.pool import ThreadedConnectionPool

        minconn = int(
            os.environ.get(
                "PG_POOL_MIN",
                "1"
            )
        )

        maxconn = int(
            os.environ.get(
                "PG_POOL_MAX",
                "4"
            )
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
            keepalives_count=3,
        )

    return _PG_POOL


# ============================================================
# DATABASE CONNECTION
# ============================================================

def db_connect():

    if is_postgres():

        return _get_pg_pool().getconn()

    conn = sqlite3.connect(
        DB_PATH
    )

    conn.row_factory = sqlite3.Row

    return conn


def db_release(conn):

    if conn is None:
        return

    if is_postgres():

        try:

            if getattr(conn, "status", None) != 1:
                conn.rollback()

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

        return sql.replace(
            "?",
            "%s"
        )

    return sql


# ============================================================
# DATABASE HELPERS
# ============================================================

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

        sql = placeholders(sql)

        if many:

            cur.executemany(
                sql,
                params
            )

        else:

            cur.execute(
                sql,
                params
            )

        rows = (
            cur.fetchall()
            if fetch
            else None
        )

        if commit:

            conn.commit()

        return rows

    except Exception:

        if commit:

            try:
                conn.rollback()
            except Exception:
                pass

        raise

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
        for i, d in enumerate(
            cur.description
        )
    }


def fetch_all(
    sql,
    params=()
):

    conn = db_connect()

    cur = conn.cursor()

    try:

        cur.execute(
            placeholders(sql),
            params
        )

        rows = cur.fetchall()

        return [
            row_dict(cur, row)
            for row in rows
        ]

    finally:

        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)


def fetch_one(
    sql,
    params=()
):

    rows = fetch_all(
        sql,
        params
    )

    if rows:

        return rows[0]

    return None


def scalar(
    sql,
    params=(),
    default=0
):

    row = fetch_one(
        sql,
        params
    )

    if not row:

        return default

    return next(
        iter(row.values())
    )


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
                False,
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
            0,
        )
    )


def columns(name):

    if not table_exists(name):

        return set()

    if is_postgres():

        rows = fetch_all(
            """
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema='public'
            AND table_name=?
            """,
            (name,),
        )

        return {
            r["column_name"]
            for r in rows
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

        db_release(conn)


def add_column(
    name,
    col,
    typ
):

    if col in columns(name):

        return

    execute(
        f"ALTER TABLE {name} ADD COLUMN {col} {typ}",
        commit=True
    )


# ============================================================
# GENERAL HELPERS
# ============================================================

def hash_password(password):

    return hashlib.sha256(
        str(password).encode("utf-8")
    ).hexdigest()


def nowstr():

    return datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


def parse_num(
    value,
    default=0.0
):

    try:

        return float(
            value or default
        )

    except Exception:

        return default


def month_name_year(
    month=None,
    year=None
):

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
        or str(
            datetime.date.today().year
        )
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


def month_name_from_date(value):

    try:

        year, month, day = (
            str(value)[:10].split("-")
        )

        return (
            f"{MONTHS[int(month)-1]} "
            f"{int(year)}"
        )

    except Exception:

        return ""


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_db():

    """
    Create only missing structures.

    IMPORTANT:
    Existing migrated data is never dropped,
    truncated, or replaced.
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

        db_release(conn)

    # --------------------------------------------------------
    # OPTIONAL WORKER COLUMNS
    # --------------------------------------------------------

    for col, typ in [
        ("phone", "TEXT"),
        ("address", "TEXT"),
        ("joining_date", "TEXT"),
        ("status", "TEXT"),
    ]:

        try:

            add_column(
                "workers",
                col,
                typ
            )

        except Exception:

            pass

    # --------------------------------------------------------
    # USER COLUMNS
    # --------------------------------------------------------

    try:

        add_column(
            "users",
            "password_hash",
            "TEXT"
        )

    except Exception:

        pass

    try:

        add_column(
            "users",
            "password",
            "TEXT"
        )

    except Exception:

        pass

    # --------------------------------------------------------
    # DEFAULT ADMIN
    # --------------------------------------------------------

    try:

        existing_admin = fetch_one(
            """
            SELECT id
            FROM users
            WHERE lower(username)=?
            LIMIT 1
            """,
            ("admin",),
        )

        if not existing_admin:

            user_columns = columns(
                "users"
            )

            data = [
                (
                    "username",
                    "admin"
                ),
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
                (
                    "active",
                    1
                ),
                (
                    "created_at",
                    nowstr()
                ),
            ]

            names = []
            values = []

            for name, value in data:

                if name in user_columns:

                    names.append(name)
                    values.append(value)

            if names:

                execute(
                    "INSERT INTO users("
                    + ",".join(names)
                    + ") VALUES("
                    + ",".join(["?"] * len(values))
                    + ")",
                    values,
                    commit=True,
                )

    except Exception as e:

        app.logger.warning(
            "Admin initialization: %r",
            e
        )

    # --------------------------------------------------------
    # DEFAULT COMPANY SETTINGS
    # --------------------------------------------------------

    defaults = {
        "company_name": DEFAULT_COMPANY_NAME,
        "company_address": "",
        "company_phone": "",
        "company_email": "",
        "company_logo": "",
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
                    commit=True,
                )

            else:

                execute(
                    """
                    INSERT OR IGNORE INTO
                    company_settings(key,value)
                    VALUES(?,?)
                    """,
                    (key, value),
                    commit=True,
                )

    except Exception as e:

        app.logger.warning(
            "Settings initialization: %r",
            e
        )


# ============================================================
# SETTINGS
# ============================================================

def get_settings():

    try:

        rows = fetch_all(
            """
            SELECT key,value
            FROM company_settings
            """
        )

        return {
            r["key"]: r["value"]
            for r in rows
        }

    except Exception:

        return {
            "company_name":
                DEFAULT_COMPANY_NAME
        }


# ============================================================
# ACTIVITY LOG
# ============================================================

def log_activity(
    username,
    action
):

    try:

        execute(
            """
            INSERT INTO activity_log(
                username,
                action,
                log_time
            )
            VALUES(?,?,?)
            """,
            (
                username,
                action,
                nowstr()
            ),
            commit=True,
        )

    except Exception as e:

        app.logger.warning(
            "Activity log error: %r",
            e
        )


# ============================================================
# USER
# ============================================================

def current_user():

    user_id = session.get(
        "user_id"
    )

    if not user_id:

        return None

    try:

        return fetch_one(
            """
            SELECT *
            FROM users
            WHERE id=?
            """,
            (user_id,),
        )

    except Exception:

        return None


# ============================================================
# LOGIN DECORATORS
# ============================================================

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

        return f(
            *args,
            **kwargs
        )

    return wrapper


def admin_required(f):

    @wraps(f)
    def wrapper(*args, **kwargs):

        user = current_user()

        if (
            not user
            or str(
                user.get(
                    "role",
                    ""
                )
            ).lower()
            not in (
                "administrator",
                "admin",
            )
        ):

            flash(
                "Administrator access required.",
                "danger"
            )

            return redirect(
                url_for(
                    "dashboard"
                )
            )

        return f(
            *args,
            **kwargs
        )

    return wrapper


# ============================================================
# ADVANCE SOURCE DETECTION
# ============================================================

def legacy_advance_source():

    for table_name in [
        "worker_advances",
        "advance_salary",
        "advances",
    ]:

        if not table_exists(
            table_name
        ):

            continue

        table_columns = columns(
            table_name
        )

        if {
            "worker_id",
            "amount",
        }.issubset(
            table_columns
        ):

            return table_name

    return None


# ============================================================
# ADVANCE LIST
# ============================================================

def advance_rows(
    month=None,
    worker_id=None
):

    table_name = (
        legacy_advance_source()
    )

    if not table_name:

        return []

    table_columns = columns(
        table_name
    )

    # --------------------------------------------------------
    # NEW WEB ADVANCE TABLE
    # --------------------------------------------------------

    if table_name == "worker_advances":

        month_col = (
            "month_year"
            if "month_year"
            in table_columns
            else None
        )

        date_col = (
            "advance_date"
            if "advance_date"
            in table_columns
            else None
        )

        note_col = (
            "note"
            if "note"
            in table_columns
            else None
        )

        select_columns = [
            "id",
            "worker_id",
            "amount",
        ]

        if month_col:

            select_columns.append(
                "month_year"
            )

        else:

            select_columns.append(
                "NULL AS month_year"
            )

        if date_col:

            select_columns.append(
                "advance_date"
            )

        else:

            select_columns.append(
                "NULL AS advance_date"
            )

        if note_col:

            select_columns.append(
                "note"
            )

        else:

            select_columns.append(
                "'' AS note"
            )

        conditions = []
        params = []

        if month and month_col:

            conditions.append(
                "month_year=?"
            )

            params.append(
                month
            )

        if worker_id:

            conditions.append(
                "worker_id=?"
            )

            params.append(
                worker_id
            )

        sql = (
            "SELECT "
            + ",".join(select_columns)
            + " FROM "
            + table_name
        )

        if conditions:

            sql += (
                " WHERE "
                + " AND ".join(
                    conditions
                )
            )

        sql += " ORDER BY id DESC"

        rows = fetch_all(
            sql,
            params
        )

        if month and not month_col:

            rows = [
                row
                for row in rows
                if month_name_from_date(
                    row.get(
                        "advance_date"
                    )
                )
                == month
            ]

        return rows

    # --------------------------------------------------------
    # LEGACY ADVANCE TABLE
    # --------------------------------------------------------

    if "date" in table_columns:

        date_column = "date"

    elif "advance_date" in table_columns:

        date_column = "advance_date"

    else:

        date_column = None

    note_column = (
        "note"
        if "note" in table_columns
        else None
    )

    select_columns = [
        "id",
        "worker_id",
        "amount",
    ]

    if date_column:

        select_columns.append(
            f"{date_column} AS advance_date"
        )

    else:

        select_columns.append(
            "NULL AS advance_date"
        )

    if note_column:

        select_columns.append(
            "note"
        )

    else:

        select_columns.append(
            "'' AS note"
        )

    conditions = []
    params = []

    if worker_id:

        conditions.append(
            "worker_id=?"
        )

        params.append(
            worker_id
        )

    sql = (
        "SELECT "
        + ",".join(select_columns)
        + " FROM "
        + table_name
    )

    if conditions:

        sql += (
            " WHERE "
            + " AND ".join(
                conditions
            )
        )

    sql += " ORDER BY id DESC"

    rows = fetch_all(
        sql,
        params
    )

    for row in rows:

        row["month_year"] = (
            month_name_from_date(
                row.get(
                    "advance_date"
                )
            )
        )

    if month:

        rows = [
            row
            for row in rows
            if row.get(
                "month_year"
            )
            == month
        ]

    return rows


# ============================================================
# SAVE ADVANCE
# ============================================================

def save_advance_record(
    worker_id,
    month,
    adv_date,
    amount,
    note
):

    table_name = (
        legacy_advance_source()
    )

    # --------------------------------------------------------
    # EXISTING WORKER ADVANCES
    # --------------------------------------------------------

    if table_name == "worker_advances":

        table_columns = columns(
            table_name
        )

        names = [
            "worker_id",
            "amount",
        ]

        values = [
            worker_id,
            amount,
        ]

        if "month_year" in table_columns:

            names.append(
                "month_year"
            )

            values.append(
                month
            )

        if "advance_date" in table_columns:

            names.append(
                "advance_date"
            )

            values.append(
                adv_date
            )

        if "note" in table_columns:

            names.append(
                "note"
            )

            values.append(
                note
            )

        execute(
            "INSERT INTO "
            + table_name
            + "("
            + ",".join(names)
            + ") VALUES("
            + ",".join(
                ["?"] * len(values)
            )
            + ")",
            values,
            commit=True,
        )

        return

    # --------------------------------------------------------
    # OLD ADVANCE TABLE
    # --------------------------------------------------------

    if table_name in (
        "advance_salary",
        "advances",
    ):

        table_columns = columns(
            table_name
        )

        if "date" in table_columns:

            date_column = "date"

        elif "advance_date" in table_columns:

            date_column = "advance_date"

        else:

            raise RuntimeError(
                "Advance table has no date column."
            )

        names = [
            "worker_id",
            date_column,
            "amount",
        ]

        values = [
            worker_id,
            adv_date,
            amount,
        ]

        if "note" in table_columns:

            names.append(
                "note"
            )

            values.append(
                note
            )

        execute(
            "INSERT INTO "
            + table_name
            + "("
            + ",".join(names)
            + ") VALUES("
            + ",".join(
                ["?"] * len(values)
            )
            + ")",
            values,
            commit=True,
        )

        return

    # --------------------------------------------------------
    # CREATE NEW ADVANCE TABLE
    # --------------------------------------------------------

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
        INSERT INTO worker_advances(
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


# ============================================================
# DELETE ADVANCE
# ============================================================

def delete_advance_record(
    advance_id
):

    table_name = (
        legacy_advance_source()
    )

    if table_name:

        execute(
            "DELETE FROM "
            + table_name
            + " WHERE id=?",
            (advance_id,),
            commit=True,
        )


# ============================================================
# DAILY ATTENDANCE
# ============================================================

def daily_map(
    worker_id,
    month
):

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
        int(row["day"]):
        str(
            row["status"]
            or "P"
        )
        for row in rows
    }


# ============================================================
# SALARY CALCULATION - BULK
# ============================================================

def calculate_salary_bulk(
    workers_list,
    month
):

    if not workers_list:

        return {}

    attendance_rows = fetch_all(
        """
        SELECT *
        FROM attendance
        WHERE month_year=?
        ORDER BY id DESC
        """,
        (month,),
    )

    attendance_by_worker = {}

    for row in attendance_rows:

        worker_id = row.get(
            "worker_id"
        )

        if worker_id not in attendance_by_worker:

            attendance_by_worker[
                worker_id
            ] = row

    daily_rows = fetch_all(
        """
        SELECT worker_id,day,status
        FROM daily_attendance
        WHERE month_year=?
        """,
        (month,),
    )

    daily_by_worker = {}

    for row in daily_rows:

        try:

            daily_by_worker.setdefault(
                row["worker_id"],
                {}
            )[int(row["day"])] = (
                str(
                    row["status"]
                    or "P"
                )
            )

        except Exception:

            pass

    advance_by_worker = {}

    try:

        all_advances = advance_rows(
            month
        )

        for row in all_advances:

            worker_id = row.get(
                "worker_id"
            )

            advance_by_worker[
                worker_id
            ] = (
                advance_by_worker.get(
                    worker_id,
                    0
                )
                + parse_num(
                    row.get(
                        "amount"
                    )
                )
            )

    except Exception as e:

        app.logger.warning(
            "Bulk advance load error: %r",
            e
        )

    try:

        month_name, year_string = (
            month.split()
        )

        year = int(
            year_string
        )

        month_number = (
            MONTHS.index(
                month_name
            ) + 1
        )

        days_in_month = (
            calendar.monthrange(
                year,
                month_number
            )[1]
        )

    except Exception:

        today = datetime.date.today()

        year = today.year
        month_number = today.month
        days_in_month = 30

    result = {}

    for worker in workers_list:

        worker_id = worker["id"]

        attendance = (
            attendance_by_worker.get(
                worker_id,
                {}
            )
        )

        present = int(
            attendance.get(
                "present_days"
            )
            or 0
        )

        absent = int(
            attendance.get(
                "absent_days"
            )
            or 0
        )

        ot_hours = parse_num(
            attendance.get(
                "ot_hours"
            )
        )

        basic_salary = parse_num(
            worker.get(
                "basic_salary"
            )
        )

        ot_rate = parse_num(
            worker.get(
                "ot_rate"
            )
        )

        nasta_rate = parse_num(
            worker.get(
                "refreshment_bill"
            )
        )

        absent_deduction = (
            basic_salary
            / days_in_month
            * absent
            if days_in_month
            else 0
        )

        earned_basic = max(
            0,
            basic_salary
            - absent_deduction
        )

        ot_amount = (
            ot_hours
            * ot_rate
        )

        daily_status = (
            daily_by_worker.get(
                worker_id,
                {}
            )
        )

        if daily_status:

            billable_days = sum(
                1
                for day in range(
                    1,
                    days_in_month + 1
                )
                if (
                    daily_status.get(
                        day,
                        "P"
                    )
                    == "P"
                    and datetime.date(
                        year,
                        month_number,
                        day
                    ).weekday()
                    != 4
                )
            )

            nasta = (
                billable_days
                * nasta_rate
            )

        else:

            non_friday_days = sum(
                1
                for day in range(
                    1,
                    days_in_month + 1
                )
                if datetime.date(
                    year,
                    month_number,
                    day
                ).weekday()
                != 4
            )

            nasta = (
                round(
                    present
                    * (
                        non_friday_days
                        / days_in_month
                    )
                    * nasta_rate,
                    2,
                )
                if days_in_month
                else 0
            )

        advance = advance_by_worker.get(
            worker_id,
            0
        )

        gross = (
            earned_basic
            + ot_amount
            + nasta
        )

        net = (
            gross
            - advance
        )

        result[worker_id] = {
            "present": present,
            "absent": absent,
            "ot": ot_hours,
            "absent_cut": absent_deduction,
            "earned_basic": earned_basic,
            "ot_amt": ot_amount,
            "nasta": nasta,
            "gross": gross,
            "advance": advance,
            "net": net,
        }

    return result


# ============================================================
# SALARY CALCULATION - SINGLE
# ============================================================

def calculate_salary(
    worker,
    month,
    attendance=None
):

    if not worker:

        return {}

    if attendance is None:

        attendance = fetch_one(
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
        attendance.get(
            "present_days"
        )
        or 0
    )

    absent = int(
        attendance.get(
            "absent_days"
        )
        or 0
    )

    ot_hours = parse_num(
        attendance.get(
            "ot_hours"
        )
    )

    try:

        month_name, year_string = (
            month.split()
        )

        year = int(
            year_string
        )

        month_number = (
            MONTHS.index(
                month_name
            ) + 1
        )

        days_in_month = (
            calendar.monthrange(
                year,
                month_number
            )[1]
        )

    except Exception:

        today = datetime.date.today()

        year = today.year
        month_number = today.month
        days_in_month = 30

    basic_salary = parse_num(
        worker.get(
            "basic_salary"
        )
    )

    ot_rate = parse_num(
        worker.get(
            "ot_rate"
        )
    )

    nasta_rate = parse_num(
        worker.get(
            "refreshment_bill"
        )
    )

    absent_deduction = (
        basic_salary
        / days_in_month
        * absent
        if days_in_month
        else 0
    )

    earned_basic = max(
        0,
        basic_salary
        - absent_deduction
    )

    ot_amount = (
        ot_hours
        * ot_rate
    )

    daily_status = daily_map(
        worker["id"],
        month
    )

    if daily_status:

        billable_days = sum(
            1
            for day in range(
                1,
                days_in_month + 1
            )
            if (
                daily_status.get(
                    day,
                    "P"
                )
                == "P"
                and datetime.date(
                    year,
                    month_number,
                    day
                ).weekday()
                != 4
            )
        )

        nasta = (
            billable_days
            * nasta_rate
        )

    else:

        non_friday_days = sum(
            1
            for day in range(
                1,
                days_in_month + 1
            )
            if datetime.date(
                year,
                month_number,
                day
            ).weekday()
            != 4
        )

        nasta = (
            round(
                present
                * (
                    non_friday_days
                    / days_in_month
                )
                * nasta_rate,
                2,
            )
            if days_in_month
            else 0
        )

    advance = sum(
        parse_num(
            row.get(
                "amount"
            )
        )
        for row in advance_rows(
            month,
            worker["id"]
        )
    )

    gross = (
        earned_basic
        + ot_amount
        + nasta
    )

    return {
        "present": present,
        "absent": absent,
        "ot": ot_hours,
        "absent_cut": absent_deduction,
        "earned_basic": earned_basic,
        "ot_amt": ot_amount,
        "nasta": nasta,
        "gross": gross,
        "advance": advance,
        "net": gross - advance,
    }


# ============================================================
# GLOBAL TEMPLATE VARIABLES
# ============================================================

@app.context_processor
def inject_globals():

    settings = get_settings()

    language = session.get(
        "language",
        "en"
    )

    def tr(text):

        if language == "bn":

            return LANG.get(
                text,
                text
            )

        return text

    def display_dept(
        department
    ):

        if language == "bn":

            return DEPT_BN.get(
                str(department),
                str(department)
            )

        return str(
            department
            or ""
        )

    return {
        "current_user":
            current_user(),

        "settings":
            settings,

        "language":
            language,

        "tr":
            tr,

        "display_dept":
            display_dept,

        "months":
            MONTHS,

        "years":
            list(
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

        return jsonify(
            {
                "status": "ok",
                "database":
                    (
                        "postgresql"
                        if is_postgres()
                        else "sqlite"
                    ),
                "workers":
                    workers_count,
                "attendance":
                    attendance_count,
                "daily_attendance":
                    daily_count,
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

    table_list = [
        "workers",
        "attendance",
        "daily_attendance",
        "company_settings",
        "users",
        "activity_log",
        "worker_advances",
        "advance_salary",
        "advances",
    ]

    for table_name in table_list:

        try:

            exists = table_exists(
                table_name
            )

            info[table_name] = {
                "exists":
                    exists,
                "columns":
                    (
                        sorted(
                            columns(
                                table_name
                            )
                        )
                        if exists
                        else []
                    ),
            }

        except Exception as e:

            info[table_name] = {
                "error":
                    repr(e)
            }

    return render_template(
        "db_check.html",
        info=info
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

        username = (
            request.form
            .get(
                "username",
                ""
            )
            .strip()
        )

        password = (
            request.form
            .get(
                "password",
                ""
            )
        )

        user = fetch_one(
            """
            SELECT *
            FROM users
            WHERE username=?
            LIMIT 1
            """,
            (username,),
        )

        valid = False

        if user:

            active = user.get(
                "active"
            )

            if active is None:

                active = 1

            if int(active):

                stored_password = (
                    user.get(
                        "password_hash"
                    )
                    or user.get(
                        "password"
                    )
                    or ""
                )

                valid = (
                    stored_password
                    in (
                        hash_password(
                            password
                        ),
                        password,
                    )
                )

        if valid:

            session[
                "user_id"
            ] = user["id"]

            session[
                "language"
            ] = session.get(
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
                        user["id"],
                    ),
                    commit=True,
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


# ============================================================
# LOGOUT
# ============================================================

@app.route("/logout")
def logout():

    user = current_user()

    if user:

        log_activity(
            user.get(
                "username"
            ),
            "Logged out"
        )

    session.clear()

    return redirect(
        url_for(
            "login"
        )
    )


# ============================================================
# LANGUAGE
# ============================================================

@app.route(
    "/language/<lang>"
)
def language(lang):

    session[
        "language"
    ] = (
        "bn"
        if lang == "bn"
        else "en"
    )

    return redirect(
        request.referrer
        or url_for(
            "dashboard"
        )
    )


# ============================================================
# HOME
# ============================================================

@app.route("/")
@login_required
def index():

    return redirect(
        url_for(
            "dashboard"
        )
    )


# ============================================================
# DASHBOARD
# ============================================================

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

    salary_map = (
        calculate_salary_bulk(
            workers_list,
            month
        )
    )

    gross = sum(
        row["gross"]
        for row in salary_map.values()
    )

    advance_total = sum(
        row["advance"]
        for row in salary_map.values()
    )

    today = datetime.date.today()

    today_month = (
        f"{MONTHS[today.month - 1]} "
        f"{today.year}"
    )

    today_rows = fetch_all(
        """
        SELECT status,COUNT(*) AS c
        FROM daily_attendance
        WHERE month_year=?
        AND day=?
        GROUP BY status
        """,
        (
            today_month,
            today.day,
        ),
    )

    today_status = {
        row["status"]:
            row["c"]
        for row in today_rows
    }

    departments = fetch_all(
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
        advance=advance_total,
        today_present=
            today_status.get(
                "P",
                0
            ),
        today_absent=
            today_status.get(
                "A",
                0
            ),
        dept_rows=departments,
        total_workers=
            len(
                workers_list
            ),
    )


# ============================================================
# WORKERS
# ============================================================

@app.route("/workers")
@login_required
def workers():

    search = (
        request.args
        .get(
            "q",
            ""
        )
        .strip()
        .lower()
    )

    params = []

    sql = """
        SELECT *
        FROM workers
    """

    if search:

        like = (
            f"%{search}%"
        )

        sql += """
            WHERE CAST(id AS TEXT) LIKE ?
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
            like,
        ]

    sql += """
        ORDER BY id DESC
    """

    return render_template(
        "workers.html",
        workers=fetch_all(
            sql,
            params
        ),
        q=search,
    )


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

        form = request.form

        name = (
            form
            .get(
                "name",
                ""
            )
            .strip()
        )

        if not name:

            flash(
                "Worker name is required.",
                "danger"
            )

            return render_template(
                "worker_form.html",
                worker=form
            )

        values = (
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
            ).strip(),
        )

        try:

            execute(
                """
                INSERT INTO workers(
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
                values,
                commit=True,
            )

            flash(
                "Worker saved.",
                "success"
            )

            return redirect(
                url_for(
                    "workers"
                )
            )

        except Exception as e:

            app.logger.exception(
                "Add worker error"
            )

            flash(
                f"Could not save worker: {e}",
                "danger"
            )

    return render_template(
        "worker_form.html",
        worker=None
    )


# ============================================================
# EDIT WORKER
# ============================================================

@app.route(
    "/workers/edit/<int:worker_id>",
    methods=["GET", "POST"]
)
@login_required
def edit_worker(
    worker_id
):

    worker = fetch_one(
        """
        SELECT *
        FROM workers
        WHERE id=?
        """,
        (worker_id,),
    )

    if not worker:

        abort(404)

    if request.method == "POST":

        form = request.form

        try:

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
                    worker_id,
                ),
                commit=True,
            )

            flash(
                "Worker updated.",
                "success"
            )

            return redirect(
                url_for(
                    "workers"
                )
            )

        except Exception as e:

            flash(
                f"Could not update worker: {e}",
                "danger"
            )

    return render_template(
        "worker_form.html",
        worker=worker
    )


# ============================================================
# DELETE WORKER
# ============================================================

@app.route(
    "/workers/delete/<int:worker_id>",
    methods=["POST", "GET"]
)
@login_required
def delete_worker(
    worker_id
):

    execute(
        """
        DELETE FROM workers
        WHERE id=?
        """,
        (worker_id,),
        commit=True,
    )

    related_tables = [
        "attendance",
        "daily_attendance",
        "worker_advances",
        "advance_salary",
        "advances",
    ]

    for table_name in related_tables:

        if table_exists(
            table_name
        ):

            try:

                execute(
                    "DELETE FROM "
                    + table_name
                    + " WHERE worker_id=?",
                    (worker_id,),
                    commit=True,
                )

            except Exception:

                pass

    flash(
        "Worker deleted.",
        "success"
    )

    return redirect(
        url_for(
            "workers"
        )
    )


# ============================================================
# ATTENDANCE
# ============================================================

@app.route("/attendance")
@login_required
def attendance():

    month = month_name_year()

    worker_id = request.args.get(
        "worker_id"
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
            (worker_id,),
        )

        if worker:

            daily_status = daily_map(
                worker["id"],
                month
            )

            month_name, year_string = (
                month.split()
            )

            year = int(
                year_string
            )

            month_number = (
                MONTHS.index(
                    month_name
                ) + 1
            )

            number_of_days = (
                calendar.monthrange(
                    year,
                    month_number
                )[1]
            )

            days = [
                {
                    "day": day,
                    "status":
                        daily_status.get(
                            day,
                            "P"
                        ),
                    "date":
                        datetime.date(
                            year,
                            month_number,
                            day
                        ),
                }
                for day in range(
                    1,
                    number_of_days + 1
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
        month=month,
    )


# ============================================================
# SAVE ATTENDANCE
# ============================================================

@app.route(
    "/attendance/save",
    methods=["POST"]
)
@login_required
def save_attendance():

    form = request.form

    try:

        worker_id = int(
            form.get(
                "worker_id"
            )
        )

    except Exception:

        flash(
            "Invalid worker.",
            "danger"
        )

        return redirect(
            url_for(
                "attendance"
            )
        )

    if form.get("year"):

        month = month_name_year(
            form.get("month"),
            form.get("year")
        )

    else:

        month = (
            form.get(
                "month_year"
            )
            or month_name_year()
        )

    present_days = int(
        form.get(
            "present_days"
        )
        or 0
    )

    absent_days = int(
        form.get(
            "absent_days"
        )
        or 0
    )

    ot_hours = parse_num(
        form.get(
            "ot_hours"
        )
    )

    daily_statuses = {}

    raw_json = form.get(
        "statuses_json"
    )

    if raw_json:

        try:

            import json

            daily_statuses = json.loads(
                raw_json
            )

        except Exception:

            daily_statuses = {}

    for key, value in form.items():

        if key.startswith(
            "day_"
        ):

            daily_statuses[
                key[4:]
            ] = value

    for day_value, status in (
        daily_statuses.items()
    ):

        try:

            day = int(
                day_value
            )

            status = (
                "A"
                if str(
                    status
                ).upper().startswith(
                    "A"
                )
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
                    INSERT INTO daily_attendance(
                        worker_id,
                        month_year,
                        day,
                        status
                    )
                    VALUES(?,?,?,?)
                    """,
                    (
                        worker_id,
                        month,
                        day,
                        status,
                    ),
                    commit=True,
                )

        except Exception:

            pass

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
                present_days,
                absent_days,
                ot_hours,
                existing["id"],
            ),
            commit=True,
        )

    else:

        attendance_columns = columns(
            "attendance"
        )

        names = [
            "worker_id",
            "month_year",
            "present_days",
            "absent_days",
            "ot_hours",
        ]

        values = [
            worker_id,
            month,
            present_days,
            absent_days,
            ot_hours,
        ]

        if (
            "advance_deduction"
            in attendance_columns
        ):

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
            commit=True,
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
            year=month.split()[1],
        )
    )


# ============================================================
# PAYSLIP
# ============================================================

@app.route("/payslip")
@login_required
def payslip():

    month = month_name_year()

    worker_id = request.args.get(
        "worker_id"
    )

    worker = None

    if worker_id:

        worker = fetch_one(
            """
            SELECT *
            FROM workers
            WHERE id=?
            """,
            (worker_id,),
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
        month=month,
    )


# ============================================================
# DEPARTMENT SALARY
# ============================================================

@app.route("/department")
@login_required
def department():

    month = month_name_year()

    department_name = request.args.get(
        "department",
        ""
    )

    sql = """
        SELECT *
        FROM workers
    """

    params = ()

    if department_name:

        sql += """
            WHERE department=?
        """

        params = (
            department_name,
        )

    sql += """
        ORDER BY id
    """

    workers_list = fetch_all(
        sql,
        params
    )

    salary_map = (
        calculate_salary_bulk(
            workers_list,
            month
        )
    )

    rows = [
        {
            **worker,
            **salary_map.get(
                worker["id"],
                {}
            ),
        }
        for worker in workers_list
    ]

    department_rows = fetch_all(
        """
        SELECT DISTINCT department
        FROM workers
        WHERE department IS NOT NULL
        AND department<>''
        ORDER BY department
        """
    )

    departments = [
        row["department"]
        for row in department_rows
    ]

    return render_template(
        "department.html",
        rows=rows,
        month=month,
        department=department_name,
        departments=departments,
    )


# ============================================================
# ADVANCE SALARY
# ============================================================

@app.route(
    "/advance",
    endpoint="advance"
)
@app.route(
    "/advances",
    endpoint="advances"
)
@app.route(
    "/advance-salary",
    endpoint="advance_salary"
)
@app.route(
    "/advance_salary",
    endpoint="advance_salary_page"
)
@login_required
def advance():

    month = month_name_year()

    worker_id = request.args.get(
        "worker_id"
    )

    workers_list = fetch_all(
        """
        SELECT id,name,department
        FROM workers
        ORDER BY id
        """
    )

    try:

        rows = advance_rows(
            month,
            worker_id
        )

    except Exception as e:

        app.logger.exception(
            "Advance rows error"
        )

        rows = []

        flash(
            f"Advance data could not be loaded: {e}",
            "danger"
        )

    return render_template(
        "advance.html",
        workers=workers_list,
        rows=rows,
        month=month,
    )


# ============================================================
# SAVE ADVANCE
# ============================================================

@app.route(
    "/advance/save",
    methods=["POST"],
    endpoint="save_advance"
)
@app.route(
    "/advances/save",
    methods=["POST"],
    endpoint="save_advance_legacy"
)
@login_required
def save_advance():

    form = request.form

    try:

        worker_id = int(
            form.get(
                "worker_id"
            )
        )

    except Exception:

        flash(
            "Invalid worker selected.",
            "danger"
        )

        return redirect(
            url_for(
                "advances"
            )
        )

    month = (
        form.get(
            "month_year"
        )
        or month_name_year(
            form.get(
                "month"
            ),
            form.get(
                "year"
            ),
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

    note = (
        form.get(
            "note",
            ""
        )
        .strip()
    )

    if amount <= 0:

        flash(
            "Amount must be greater than zero.",
            "danger"
        )

        return redirect(
            url_for(
                "advances",
                worker_id=worker_id
            )
        )

    try:

        save_advance_record(
            worker_id,
            month,
            advance_date,
            amount,
            note,
        )

        flash(
            "Advance saved successfully.",
            "success"
        )

    except Exception as e:

        app.logger.exception(
            "Save advance error"
        )

        flash(
            f"Could not save advance: {e}",
            "danger"
        )

    return redirect(
        url_for(
            "advances",
            worker_id=worker_id,
            month=month.split()[0],
            year=month.split()[1],
        )
    )


# ============================================================
# DELETE ADVANCE
# ============================================================

@app.route(
    "/advance/delete/<int:aid>",
    endpoint="delete_advance"
)
@app.route(
    "/advances/delete/<int:aid>",
    endpoint="delete_advance_legacy"
)
@login_required
def delete_advance(
    aid
):

    try:

        delete_advance_record(
            aid
        )

        flash(
            "Advance deleted successfully.",
            "success"
        )

    except Exception as e:

        app.logger.exception(
            "Delete advance error"
        )

        flash(
            f"Could not delete advance: {e}",
            "danger"
        )

    return redirect(
        url_for(
            "advances"
        )
    )


# ============================================================
# LEGACY ENDPOINT COMPATIBILITY
# ============================================================

app.add_url_rule(
    "/worker-management",
    endpoint="worker_management",
    view_func=workers,
    methods=["GET"],
)

app.add_url_rule(
    "/worker-list",
    endpoint="worker_list",
    view_func=workers,
    methods=["GET"],
)

app.add_url_rule(
    "/attendance-calendar",
    endpoint="attendance_calendar",
    view_func=attendance,
    methods=["GET"],
)

app.add_url_rule(
    "/single-payslip",
    endpoint="single_payslip",
    view_func=payslip,
    methods=["GET"],
)

app.add_url_rule(
    "/department-salary-sheet",
    endpoint="department_salary_sheet",
    view_func=department,
    methods=["GET"],
)

app.add_url_rule(
    "/report",
    endpoint="report",
    view_func=department,
    methods=["GET"],
)


# ============================================================
# SETTINGS
# ============================================================

@app.route(
    "/settings",
    methods=["GET", "POST"]
)
@admin_required
def settings():

    if request.method == "POST":

        setting_keys = [
            "company_name",
            "company_address",
            "company_phone",
            "company_email",
            "company_logo",
        ]

        for key in setting_keys:

            value = request.form.get(
                key,
                ""
            )

            if is_postgres():

                execute(
                    """
                    INSERT INTO company_settings(
                        key,
                        value
                    )
                    VALUES(?,?)
                    ON CONFLICT(key)
                    DO UPDATE SET
                        value=EXCLUDED.value
                    """,
                    (
                        key,
                        value,
                    ),
                    commit=True,
                )

            else:

                execute(
                    """
                    INSERT OR REPLACE INTO
                    company_settings(
                        key,
                        value
                    )
                    VALUES(?,?)
                    """,
                    (
                        key,
                        value,
                    ),
                    commit=True,
                )

        flash(
            "Settings saved.",
            "success"
        )

    return render_template(
        "settings.html",
        settings=get_settings()
    )


# ============================================================
# USERS
# ============================================================

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
        ),
    )


# ============================================================
# ADD USER
# ============================================================

@app.route(
    "/users/add",
    methods=["GET", "POST"]
)
@admin_required
def add_user():

    if request.method == "POST":

        form = request.form

        username = (
            form.get(
                "username",
                ""
            )
            .strip()
        )

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

        password_hash = (
            hash_password(
                password
            )
        )

        user_columns = columns(
            "users"
        )

        data = [
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
            ),
        ]

        names = []
        values = []

        for name, value in data:

            if name in user_columns:

                names.append(
                    name
                )

                values.append(
                    value
                )

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
                commit=True,
            )

            flash(
                "User created.",
                "success"
            )

            return redirect(
                url_for(
                    "users"
                )
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


# ============================================================
# DELETE USER
# ============================================================

@app.route(
    "/users/delete/<int:user_id>"
)
@admin_required
def delete_user(
    user_id
):

    user = current_user()

    if (
        user
        and user.get(
            "id"
        )
        == user_id
    ):

        flash(
            "You cannot delete the logged-in user.",
            "danger"
        )

    else:

        execute(
            """
            DELETE FROM users
            WHERE id=?
            """,
            (
                user_id,
            ),
            commit=True,
        )

        flash(
            "User deleted.",
            "success"
        )

    return redirect(
        url_for(
            "users"
        )
    )


# ============================================================
# ACTIVITY
# ============================================================

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
        ),
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

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
            ),
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
                    row.get(
                        header
                    )
                    for header in headers
                ]
            )

    else:

        worksheet.append(
            ["No records"]
        )

    output = io.BytesIO()

    workbook.save(
        output
    )

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# ============================================================
# WORKER EXCEL
# ============================================================

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
        "reedoy_workers.xlsx",
    )


# ============================================================
# ADVANCE EXCEL
# ============================================================

@app.route(
    "/export/advances.xlsx"
)
@login_required
def export_advances():

    return export_rows(
        advance_rows(
            month_name_year()
        ),
        "reedoy_advances.xlsx",
    )


# ============================================================
# DEPARTMENT EXCEL
# ============================================================

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

    sql = """
        SELECT *
        FROM workers
    """

    params = ()

    if department_name:

        sql += """
            WHERE department=?
        """

        params = (
            department_name,
        )

    sql += """
        ORDER BY id
    """

    workers_list = fetch_all(
        sql,
        params
    )

    salary_map = (
        calculate_salary_bulk(
            workers_list,
            month
        )
    )

    rows = []

    for worker in workers_list:

        row = {
            "ID":
                worker["id"],

            "Worker Name":
                worker["name"],

            "Department":
                worker.get(
                    "department"
                ),

            "Designation":
                worker.get(
                    "designation"
                ),
        }

        row.update(
            salary_map.get(
                worker["id"],
                {}
            )
        )

        rows.append(
            row
        )

    return export_rows(
        rows,
        "reedoy_department_salary.xlsx",
    )


# ============================================================
# PAYSLIP PDF
# ============================================================

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
            ),
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
        (
            worker_id,
        ),
    )

    if not worker:

        abort(404)

    salary = calculate_salary(
        worker,
        month
    )

    output = io.BytesIO()

    document = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=35,
        leftMargin=35,
        topMargin=35,
        bottomMargin=35,
    )

    styles = (
        getSampleStyleSheet()
    )

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
        Spacer(
            1,
            12
        ),
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
        ],
    ]

    story.append(
        Table(
            data,
            colWidths=[
                180,
                280
            ],
            style=TableStyle(
                [
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.5,
                        colors.grey,
                    ),
                    (
                        "FONTNAME",
                        (0, 0),
                        (-1, -1),
                        "Helvetica",
                    ),
                    (
                        "PADDING",
                        (0, 0),
                        (-1, -1),
                        6,
                    ),
                ]
            ),
        )
    )

    document.build(
        story
    )

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=(
            f"payslip_"
            f"{worker['id']}_"
            f"{month.replace(' ', '_')}.pdf"
        ),
        mimetype="application/pdf",
    )


# ============================================================
# WORKERS CSV
# ============================================================

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

    output = io.StringIO()

    if rows:

        writer = csv.DictWriter(
            output,
            fieldnames=list(
                rows[0].keys()
            )
        )

        writer.writeheader()

        writer.writerows(
            rows
        )

    return send_file(
        io.BytesIO(
            output
            .getvalue()
            .encode(
                "utf-8-sig"
            )
        ),
        as_attachment=True,
        download_name=(
            "reedoy_workers.csv"
        ),
        mimetype="text/csv",
    )


# ============================================================
# ERROR HANDLERS
# ============================================================

@app.errorhandler(404)
def not_found(error):

    try:

        return (
            render_template(
                "404.html"
            ),
            404,
        )

    except Exception:

        return (
            "404 - Page not found",
            404,
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
            500,
        )

    except Exception:

        return (
            f"500 - Internal Server Error\n{error}",
            500,
        )


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

try:

    init_db()

except Exception as e:

    app.logger.exception(
        "Database initialization error: %r",
        e
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
                5000
            )
        ),
        debug=False,
    )
