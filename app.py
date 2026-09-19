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
# OPTIONAL LIBRARIES
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
# APP CONFIG
# ============================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "REEDOY-PAYROLL-CHANGE-THIS-SECRET"
)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
DB_PATH = os.environ.get("SQLITE_DB", "reedoy_payroll.db")

DEFAULT_COMPANY_NAME = "REEDOY TEXTILE DYEING PRINTING & FINISHING"


# ============================================================
# CONSTANTS
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

LANG = {
    "Dashboard": "ড্যাশবোর্ড",
    "Workers Management": "কর্মী ব্যবস্থাপনা",
    "Attendance & Calendar": "উপস্থিতি ও ক্যালেন্ডার",
    "Single Payslip": "একক পে-স্লিপ",
    "Advance Salary": "অগ্রিম বেতন",
    "Department Salary Sheet": "বিভাগভিত্তিক বেতন শীট",
    "Settings": "সেটিংস",
    "Users": "ইউজার",
    "Activity Log": "কার্যক্রম",
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
}


# ============================================================
# DATABASE
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

        minconn = int(os.environ.get("PG_POOL_MIN", "1"))
        maxconn = int(os.environ.get("PG_POOL_MAX", "4"))

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


def db_connect():
    if is_postgres():
        return _get_pg_pool().getconn()

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
    commit=False,
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

    except Exception:
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
        for i, d in enumerate(cur.description)
    }


def fetch_all(sql, params=()):
    conn = db_connect()
    cur = conn.cursor()

    try:
        cur.execute(placeholders(sql), params)
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


def fetch_one(sql, params=()):
    rows = fetch_all(sql, params)
    return rows[0] if rows else None


def scalar(sql, params=(), default=0):
    row = fetch_one(sql, params)

    if not row:
        return default

    return next(iter(row.values()))


# ============================================================
# DATABASE HELPERS
# ============================================================

def table_exists(name):
    try:
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

    except Exception:
        return False


def columns(name):
    if not table_exists(name):
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

    except Exception:
        return set()


def add_column(name, column, column_type):
    try:
        if column not in columns(name):
            execute(
                f"ALTER TABLE {name} "
                f"ADD COLUMN {column} {column_type}",
                commit=True,
            )
    except Exception as e:
        app.logger.warning(
            "Could not add column %s.%s: %r",
            name,
            column,
            e,
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


def parse_num(value, default=0.0):
    try:
        return float(value or default)
    except Exception:
        return default


def month_name_year(month=None, year=None):
    month = (
        month
        or request.values.get("month")
        or MONTHS[datetime.date.today().month - 1]
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


def month_name_from_date(value):
    try:
        text = str(value)[:10]
        y, m, d = text.split("-")

        return (
            f"{MONTHS[int(m) - 1]} "
            f"{int(y)}"
        )

    except Exception:
        return ""


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_db():
    """
    IMPORTANT:
    This function NEVER drops existing tables,
    NEVER truncates data and NEVER imports legacy data.
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

    # Additional worker fields
    for col, typ in [
        ("phone", "TEXT"),
        ("address", "TEXT"),
        ("joining_date", "TEXT"),
        ("status", "TEXT"),
    ]:
        add_column("workers", col, typ)

    add_column("users", "password_hash", "TEXT")
    add_column("users", "password", "TEXT")

    # Default admin
    try:
        existing = fetch_one(
            """
            SELECT id
            FROM users
            WHERE lower(username)=?
            LIMIT 1
            """,
            ("admin",),
        )

        if not existing:
            cols = columns("users")

            names = []
            values = []

            defaults = [
                (
                    "username",
                    "admin",
                ),
                (
                    "password_hash",
                    hash_password("admin123"),
                ),
                (
                    "password",
                    hash_password("admin123"),
                ),
                (
                    "full_name",
                    "System Administrator",
                ),
                (
                    "role",
                    "Administrator",
                ),
                (
                    "active",
                    1,
                ),
                (
                    "created_at",
                    nowstr(),
                ),
            ]

            for name, value in defaults:
                if name in cols:
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
            "Admin initialization error: %r",
            e,
        )

    # Default company settings
    defaults = {
        "company_name": DEFAULT_COMPANY_NAME,
        "company_address": "",
        "company_phone": "",
        "company_email": "",
        "company_logo": "",
    }

    for key, value in defaults.items():
        try:
            if is_postgres():
                execute(
                    """
                    INSERT INTO company_settings(key,value)
                    VALUES(?,?)
                    ON CONFLICT(key) DO NOTHING
                    """,
                    (key, value),
                    commit=True,
                )
            else:
                execute(
                    """
                    INSERT OR IGNORE
                    INTO company_settings(key,value)
                    VALUES(?,?)
                    """,
                    (key, value),
                    commit=True,
                )

        except Exception as e:
            app.logger.warning(
                "Settings initialization error: %r",
                e,
            )


def get_settings():
    try:
        return {
            row["key"]: row["value"]
            for row in fetch_all(
                "SELECT key,value FROM company_settings"
            )
        }

    except Exception:
        return {
            "company_name":
                DEFAULT_COMPANY_NAME
        }


# ============================================================
# USER / LOGIN
# ============================================================

def current_user():
    user_id = session.get("user_id")

    if not user_id:
        return None

    try:
        return fetch_one(
            "SELECT * FROM users WHERE id=?",
            (user_id,),
        )
    except Exception:
        return None


def login_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        if not session.get("user_id"):
            return redirect(
                url_for(
                    "login",
                    next=request.path,
                )
            )

        return function(*args, **kwargs)

    return wrapper


def admin_required(function):
    @wraps(function)
    def wrapper(*args, **kwargs):
        user = current_user()

        if not user:
            return redirect(
                url_for("login")
            )

        role = str(
            user.get("role", "")
        ).lower()

        if role not in (
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

        return function(*args, **kwargs)

    return wrapper


def log_activity(username, action):
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
                nowstr(),
            ),
            commit=True,
        )

    except Exception as e:
        app.logger.warning(
            "Activity log error: %r",
            e,
        )


# ============================================================
# ADVANCE TABLE DETECTION
# ============================================================

def legacy_advance_source():
    """
    Detect existing advance table.
    """

    candidates = [
        "worker_advances",
        "advance_salary",
        "advances",
    ]

    for table in candidates:
        try:
            if not table_exists(table):
                continue

            cols = columns(table)

            if "worker_id" in cols and "amount" in cols:
                return table

        except Exception:
            continue

    return None


def advance_rows(month=None, worker_id=None):
    table = legacy_advance_source()

    if not table:
        return []

    cols = columns(table)

    try:
        # ----------------------------------------------------
        # worker_advances
        # ----------------------------------------------------
        if table == "worker_advances":

            date_column = (
                "advance_date"
                if "advance_date" in cols
                else (
                    "date"
                    if "date" in cols
                    else None
                )
            )

            month_column = (
                "month_year"
                if "month_year" in cols
                else None
            )

            note_column = (
                "note"
                if "note" in cols
                else None
            )

            select_date = (
                f"{date_column} AS advance_date"
                if date_column
                else "NULL AS advance_date"
            )

            select_month = (
                month_column
                if month_column
                else (
                    f"NULL AS month_year"
                )
            )

            select_note = (
                note_column
                if note_column
                else "'' AS note"
            )

            sql = f"""
                SELECT
                    id,
                    worker_id,
                    {select_month},
                    {select_date},
                    amount,
                    {select_note}
                FROM {table}
            """

            conditions = []
            params = []

            if month and month_column:
                conditions.append(
                    f"{month_column}=?"
                )
                params.append(month)

            if worker_id:
                conditions.append(
                    "worker_id=?"
                )
                params.append(worker_id)

            if conditions:
                sql += (
                    " WHERE "
                    + " AND ".join(conditions)
                )

            sql += " ORDER BY id DESC"

            rows = fetch_all(sql, params)

            # If old rows don't have month_year,
            # derive it from date.
            if month:
                rows = [
                    r for r in rows
                    if (
                        not r.get("month_year")
                        or r.get("month_year") == month
                        or month_name_from_date(
                            r.get("advance_date")
                        ) == month
                    )
                ]

            return rows

        # ----------------------------------------------------
        # advance_salary / advances
        # ----------------------------------------------------

        date_column = None

        for candidate in [
            "advance_date",
            "date",
            "created_at",
        ]:
            if candidate in cols:
                date_column = candidate
                break

        note_column = (
            "note"
            if "note" in cols
            else None
        )

        select_date = (
            f"{date_column} AS advance_date"
            if date_column
            else "NULL AS advance_date"
        )

        select_note = (
            note_column
            if note_column
            else "'' AS note"
        )

        sql = f"""
            SELECT
                id,
                worker_id,
                {select_date},
                amount,
                {select_note}
            FROM {table}
        """

        params = []

        if worker_id:
            sql += " WHERE worker_id=?"
            params.append(worker_id)

        sql += " ORDER BY id DESC"

        rows = fetch_all(
            sql,
            params,
        )

        for row in rows:
            row["month_year"] = (
                month_name_from_date(
                    row.get("advance_date")
                )
            )

        if month:
            rows = [
                row for row in rows
                if row.get("month_year") == month
            ]

        return rows

    except Exception as e:
        app.logger.exception(
            "Advance loading error: %r",
            e,
        )
        return []


def create_worker_advances_table():
    if table_exists("worker_advances"):
        return

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
            commit=True,
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
            commit=True,
        )


def save_advance_record(
    worker_id,
    month,
    advance_date,
    amount,
    note,
):
    table = legacy_advance_source()

    if not table:
        create_worker_advances_table()
        table = "worker_advances"

    cols = columns(table)

    names = []
    values = []

    if "worker_id" in cols:
        names.append("worker_id")
        values.append(worker_id)

    if "month_year" in cols:
        names.append("month_year")
        values.append(month)

    if "advance_date" in cols:
        names.append("advance_date")
        values.append(advance_date)
    elif "date" in cols:
        names.append("date")
        values.append(advance_date)

    if "amount" in cols:
        names.append("amount")
        values.append(amount)

    if "note" in cols:
        names.append("note")
        values.append(note)

    if not names:
        raise RuntimeError(
            "Advance table has no usable columns."
        )

    execute(
        "INSERT INTO "
        + table
        + "("
        + ",".join(names)
        + ") VALUES("
        + ",".join(["?"] * len(values))
        + ")",
        values,
        commit=True,
    )


def delete_advance_record(advance_id):
    table = legacy_advance_source()

    if not table:
        return

    if "id" not in columns(table):
        return

    execute(
        f"DELETE FROM {table} WHERE id=?",
        (advance_id,),
        commit=True,
    )


# ============================================================
# SALARY CALCULATION
# ============================================================

def daily_map(worker_id, month):
    if not table_exists("daily_attendance"):
        return {}

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

    result = {}

    for row in rows:
        try:
            result[
                int(row["day"])
            ] = str(
                row["status"] or "P"
            )
        except Exception:
            pass

    return result


def days_in_month(month):
    try:
        month_name, year_text = month.split()

        year = int(year_text)
        month_number = (
            MONTHS.index(month_name) + 1
        )

        return (
            year,
            month_number,
            calendar.monthrange(
                year,
                month_number,
            )[1],
        )

    except Exception:
        today = datetime.date.today()

        return (
            today.year,
            today.month,
            calendar.monthrange(
                today.year,
                today.month,
            )[1],
        )


def advance_total_for_worker(
    worker_id,
    month,
):
    rows = advance_rows(
        month,
        worker_id,
    )

    return sum(
        parse_num(
            row.get("amount")
        )
        for row in rows
    )


def calculate_salary(
    worker,
    month,
    attendance=None,
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
            "present_days",
            0,
        ) or 0
    )

    absent = int(
        attendance.get(
            "absent_days",
            0,
        ) or 0
    )

    ot_hours = parse_num(
        attendance.get(
            "ot_hours",
            0,
        )
    )

    year, month_number, total_days = (
        days_in_month(month)
    )

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
        (basic / total_days) * absent
        if total_days
        else 0
    )

    earned_basic = max(
        0,
        basic - absent_cut,
    )

    ot_amount = (
        ot_hours * ot_rate
    )

    daily = daily_map(
        worker["id"],
        month,
    )

    if daily:
        nasta_days = 0

        for day in range(
            1,
            total_days + 1,
        ):
            status = daily.get(
                day,
                "P",
            )

            date_obj = datetime.date(
                year,
                month_number,
                day,
            )

            # Friday = 4
            if (
                status == "P"
                and date_obj.weekday() != 4
            ):
                nasta_days += 1

        nasta = (
            nasta_days * nasta_rate
        )

    else:
        non_friday_days = sum(
            1
            for day in range(
                1,
                total_days + 1,
            )
            if datetime.date(
                year,
                month_number,
                day,
            ).weekday() != 4
        )

        nasta = (
            present
            * (
                non_friday_days
                / total_days
            )
            * nasta_rate
            if total_days
            else 0
        )

    advance = advance_total_for_worker(
        worker["id"],
        month,
    )

    gross = (
        earned_basic
        + ot_amount
        + nasta
    )

    net = gross - advance

    return {
        "present": present,
        "absent": absent,
        "ot": ot_hours,
        "absent_cut": absent_cut,
        "earned_basic": earned_basic,
        "ot_amt": ot_amount,
        "nasta": nasta,
        "gross": gross,
        "advance": advance,
        "net": net,
    }


def calculate_salary_bulk(
    workers,
    month,
):
    result = {}

    for worker in workers:
        try:
            result[worker["id"]] = (
                calculate_salary(
                    worker,
                    month,
                )
            )
        except Exception as e:
            app.logger.exception(
                "Salary calculation error "
                "for worker %s: %r",
                worker.get("id"),
                e,
            )

            result[worker["id"]] = {
                "present": 0,
                "absent": 0,
                "ot": 0,
                "absent_cut": 0,
                "earned_basic":
                    parse_num(
                        worker.get(
                            "basic_salary"
                        )
                    ),
                "ot_amt": 0,
                "nasta": 0,
                "gross":
                    parse_num(
                        worker.get(
                            "basic_salary"
                        )
                    ),
                "advance": 0,
                "net":
                    parse_num(
                        worker.get(
                            "basic_salary"
                        )
                    ),
            }

    return result


# ============================================================
# TEMPLATE GLOBALS
# ============================================================

@app.context_processor
def inject_globals():
    settings = get_settings()

    language = session.get(
        "language",
        "en",
    )

    def tr(text):
        if language == "bn":
            return LANG.get(
                text,
                text,
            )

        return text

    def display_department(value):
        if language == "bn":
            return DEPT_BN.get(
                str(value),
                str(value or ""),
            )

        return str(value or "")

    return {
        "current_user": current_user(),
        "settings": settings,
        "language": language,
        "tr": tr,
        "display_dept": display_department,
        "months": MONTHS,
        "years": list(
            range(
                2024,
                2032,
            )
        ),
    }


# ============================================================
# HEALTH CHECK
# ============================================================

@app.route("/health")
def health():
    try:
        workers_count = scalar(
            "SELECT COUNT(*) FROM workers",
            default=0,
        )

        attendance_count = scalar(
            "SELECT COUNT(*) FROM attendance",
            default=0,
        )

        daily_count = scalar(
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
# DB CHECK
# ============================================================

@app.route("/db-check")
@login_required
def db_check():
    information = {}

    tables = [
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

    for table in tables:
        try:
            exists = table_exists(table)

            information[table] = {
                "exists": exists,
                "columns": (
                    sorted(
                        columns(table)
                    )
                    if exists
                    else []
                ),
            }

        except Exception as e:
            information[table] = {
                "error": repr(e)
            }

    return render_template(
        "db_check.html",
        info=information,
    )


# ============================================================
# LOGIN
# ============================================================

@app.route(
    "/login",
    methods=["GET", "POST"],
)
def login():
    if request.method == "POST":

        username = request.form.get(
            "username",
            "",
        ).strip()

        password = request.form.get(
            "password",
            "",
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
                "active",
                1,
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

                valid = (
                    stored
                    == hash_password(
                        password
                    )
                    or stored
                    == password
                )

        if valid:
            session["user_id"] = user["id"]

            if "language" not in session:
                session["language"] = "en"

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
                "Logged in",
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


@app.route("/logout")
def logout():
    user = current_user()

    if user:
        log_activity(
            user.get("username"),
            "Logged out",
        )

    session.clear()

    return redirect(
        url_for("login")
    )


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
# DASHBOARD
# ============================================================

@app.route("/")
@login_required
def index():
    return redirect(
        url_for("dashboard")
    )


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

    salary_map = calculate_salary_bulk(
        workers_list,
        month,
    )

    gross = sum(
        salary["gross"]
        for salary in salary_map.values()
    )

    advance_total = sum(
        salary["advance"]
        for salary in salary_map.values()
    )

    today = datetime.date.today()

    today_month = (
        f"{MONTHS[today.month - 1]} "
        f"{today.year}"
    )

    today_rows = []

    try:
        today_rows = fetch_all(
            """
            SELECT status, COUNT(*) AS c
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
    except Exception:
        pass

    today_map = {
        row["status"]: row["c"]
        for row in today_rows
    }

    department_rows = []

    try:
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
    except Exception:
        department_rows = []

    return render_template(
        "dashboard.html",
        workers=workers_list,
        month=month,
        month_name=month.split()[0],
        year=month.split()[1],
        gross=gross,
        advance=advance_total,
        today_present=today_map.get(
            "P",
            0,
        ),
        today_absent=today_map.get(
            "A",
            0,
        ),
        dept_rows=department_rows,
        total_workers=len(
            workers_list
        ),
    )


# ============================================================
# WORKERS
# ============================================================

@app.route("/workers")
@login_required
def workers():
    query = request.args.get(
        "q",
        "",
    ).strip()

    params = []

    sql = """
        SELECT *
        FROM workers
    """

    if query:
        like = f"%{query.lower()}%"

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
            like,
        ]

    sql += " ORDER BY id DESC"

    worker_rows = fetch_all(
        sql,
        params,
    )

    return render_template(
        "workers.html",
        workers=worker_rows,
        q=query,
    )


@app.route(
    "/workers/add",
    methods=["GET", "POST"],
)
@login_required
def add_worker():
    if request.method == "POST":

        form = request.form

        name = form.get(
            "name",
            "",
        ).strip()

        if not name:
            flash(
                "Worker name is required.",
                "danger",
            )

            return render_template(
                "worker_form.html",
                worker=form,
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
                "",
            ).strip(),
            form.get(
                "designation",
                "",
            ).strip(),
            parse_num(
                form.get(
                    "refreshment_bill"
                )
            ),
            form.get(
                "bangla_name",
                "",
            ).strip(),
        )

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
            "success",
        )

        return redirect(
            url_for("workers")
        )

    return render_template(
        "worker_form.html",
        worker=None,
    )


@app.route(
    "/workers/edit/<int:worker_id>",
    methods=["GET", "POST"],
)
@login_required
def edit_worker(worker_id):
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
                    "",
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
                    "",
                ),
                form.get(
                    "designation",
                    "",
                ),
                parse_num(
                    form.get(
                        "refreshment_bill"
                    )
                ),
                form.get(
                    "bangla_name",
                    "",
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

    return render_template(
        "worker_form.html",
        worker=worker,
    )


@app.route(
    "/workers/delete/<int:worker_id>",
    methods=["GET", "POST"],
)
@login_required
def delete_worker(worker_id):
    execute(
        """
        DELETE FROM workers
        WHERE id=?
        """,
        (worker_id,),
        commit=True,
    )

    for table in [
        "attendance",
        "daily_attendance",
        "worker_advances",
        "advance_salary",
        "advances",
    ]:
        try:
            if table_exists(table):
                if "worker_id" in columns(table):
                    execute(
                        f"""
                        DELETE FROM {table}
                        WHERE worker_id=?
                        """,
                        (worker_id,),
                        commit=True,
                    )
        except Exception as e:
            app.logger.warning(
                "Delete related data error: %r",
                e,
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

@app.route("/attendance")
@login_required
def attendance():
    month = month_name_year()

    worker_id = request.args.get(
        "worker_id",
        "",
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
            year, month_number, total_days = (
                days_in_month(month)
            )

            daily = daily_map(
                worker["id"],
                month,
            )

            days = [
                {
                    "day": day,
                    "status":
                        daily.get(
                            day,
                            "P",
                        ),
                    "date":
                        datetime.date(
                            year,
                            month_number,
                            day,
                        ),
                }
                for day in range(
                    1,
                    total_days + 1,
                )
            ]

            summary = calculate_salary(
                worker,
                month,
            )

    worker_list = fetch_all(
        """
        SELECT
            id,
            name,
            department
        FROM workers
        ORDER BY id
        """
    )

    return render_template(
        "attendance.html",
        workers=worker_list,
        worker=worker,
        days=days,
        summary=summary,
        month=month,
    )


@app.route(
    "/attendance/save",
    methods=["POST"],
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
            "danger",
        )

        return redirect(
            url_for("attendance")
        )

    if form.get("year"):
        month = month_name_year(
            form.get("month"),
            form.get("year"),
        )
    else:
        month = (
            form.get("month_year")
            or month_name_year()
        )

    present = int(
        form.get(
            "present_days",
            0,
        )
        or 0
    )

    absent = int(
        form.get(
            "absent_days",
            0,
        )
        or 0
    )

    ot_hours = parse_num(
        form.get(
            "ot_hours",
            0,
        )
    )

    # Daily statuses
    status_map = {}

    raw_json = form.get(
        "statuses_json"
    )

    if raw_json:
        try:
            import json

            status_map = json.loads(
                raw_json
            )
        except Exception:
            status_map = {}

    for key, value in form.items():
        if key.startswith("day_"):
            status_map[
                key[4:]
            ] = value

    # Save daily attendance
    for day_text, status in (
        status_map.items()
    ):
        try:
            day = int(day_text)

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

        except Exception as e:
            app.logger.warning(
                "Daily attendance save error: %r",
                e,
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
            present,
            absent,
            ot_hours,
        ]

        if "advance_deduction" in (
            attendance_columns
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
        "success",
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
            summary = calculate_salary(
                worker,
                month,
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
# DEPARTMENT
# ============================================================

@app.route("/department")
@login_required
def department():
    month = month_name_year()

    selected_department = request.args.get(
        "department",
        "",
    ).strip()

    try:
        if selected_department:
            workers_list = fetch_all(
                """
                SELECT *
                FROM workers
                WHERE department=?
                ORDER BY id
                """,
                (
                    selected_department,
                ),
            )
        else:
            workers_list = fetch_all(
                """
                SELECT *
                FROM workers
                ORDER BY id
                """
            )

        salary_map = calculate_salary_bulk(
            workers_list,
            month,
        )

        rows = []

        for worker in workers_list:
            salary = salary_map.get(
                worker["id"],
                {},
            )

            combined = dict(worker)
            combined.update(salary)

            rows.append(combined)

        department_rows = fetch_all(
            """
            SELECT DISTINCT department
            FROM workers
            WHERE department IS NOT NULL
            AND department <> ''
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
            department=selected_department,
            departments=departments,
        )

    except Exception as e:
        app.logger.exception(
            "Department page error: %r",
            e,
        )

        flash(
            "Department page error: "
            + str(e),
            "danger",
        )

        return render_template(
            "department.html",
            rows=[],
            month=month,
            department=selected_department,
            departments=[],
        )


# ============================================================
# ADVANCE SALARY
#
# IMPORTANT:
# Your repository has advances.html,
# NOT advance.html.
# ============================================================

@app.route("/advance")
@login_required
def advance():
    month = month_name_year()

    worker_id = request.args.get(
        "worker_id"
    )

    try:
        worker_list = fetch_all(
            """
            SELECT
                id,
                name,
                department
            FROM workers
            ORDER BY id
            """
        )

        rows = advance_rows(
            month,
            worker_id,
        )

        return render_template(
            "advances.html",
            workers=worker_list,
            rows=rows,
            month=month,
        )

    except Exception as e:
        app.logger.exception(
            "Advance page error: %r",
            e,
        )

        return render_template(
            "advances.html",
            workers=[],
            rows=[],
            month=month,
        )


@app.route(
    "/advance/save",
    methods=["POST"],
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
            "Invalid worker.",
            "danger",
        )

        return redirect(
            url_for("advance")
        )

    month = (
        form.get("month_year")
        or month_name_year(
            form.get("month"),
            form.get("year"),
        )
    )

    advance_date = (
        form.get("advance_date")
        or datetime.date.today().isoformat()
    )

    amount = parse_num(
        form.get("amount")
    )

    note = form.get(
        "note",
        "",
    )

    if amount <= 0:
        flash(
            "Amount must be greater than zero.",
            "danger",
        )

        return redirect(
            url_for(
                "advance",
                month=month,
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
            "Advance saved.",
            "success",
        )

    except Exception as e:
        app.logger.exception(
            "Advance save error: %r",
            e,
        )

        flash(
            "Advance save error: "
            + str(e),
            "danger",
        )

    return redirect(
        url_for(
            "advance",
            month=month,
        )
    )


@app.route(
    "/advance/delete/<int:advance_id>",
    methods=["GET", "POST"],
)
@login_required
def delete_advance(advance_id):
    try:
        delete_advance_record(
            advance_id
        )

        flash(
            "Advance deleted.",
            "success",
        )

    except Exception as e:
        app.logger.exception(
            "Advance delete error: %r",
            e,
        )

        flash(
            "Could not delete advance.",
            "danger",
        )

    return redirect(
        url_for("advance")
    )


# ============================================================
# SETTINGS
# ============================================================

@app.route(
    "/settings",
    methods=["GET", "POST"],
)
@admin_required
def settings():
    if request.method == "POST":

        fields = [
            "company_name",
            "company_address",
            "company_phone",
            "company_email",
            "company_logo",
        ]

        for key in fields:
            value = request.form.get(
                key,
                "",
            )

            if is_postgres():
                execute(
                    """
                    INSERT INTO company_settings(
                        key,value
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
                    INSERT OR REPLACE
                    INTO company_settings(
                        key,value
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
            "success",
        )

    return render_template(
        "settings.html",
        settings=get_settings(),
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


@app.route(
    "/users/add",
    methods=["GET", "POST"],
)
@admin_required
def add_user():
    if request.method == "POST":

        form = request.form

        username = form.get(
            "username",
            "",
        ).strip()

        password = form.get(
            "password",
            "",
        )

        if not username or not password:
            flash(
                "Username and password are required.",
                "danger",
            )

            return render_template(
                "user_form.html",
                user=form,
            )

        password_hash = hash_password(
            password
        )

        cols = columns("users")

        names = []
        values = []

        values_map = {
            "username": username,
            "password_hash":
                password_hash,
            "password":
                password_hash,
            "full_name":
                form.get(
                    "full_name",
                    "",
                ),
            "role":
                form.get(
                    "role",
                    "Operator",
                ),
            "active": 1,
            "created_at": nowstr(),
        }

        for name, value in (
            values_map.items()
        ):
            if name in cols:
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
            flash(
                "Could not create user: "
                + str(e),
                "danger",
            )

    return render_template(
        "user_form.html",
        user=None,
    )


@app.route(
    "/users/delete/<int:user_id>",
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
            "danger",
        )

    else:
        execute(
            """
            DELETE FROM users
            WHERE id=?
            """,
            (user_id,),
            commit=True,
        )

        flash(
            "User deleted.",
            "success",
        )

    return redirect(
        url_for("users")
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
# REPORT
# ============================================================

@app.route("/report")
@login_required
def report():
    return department()


# ============================================================
# EXCEL EXPORT
# ============================================================

def export_rows(
    rows,
    filename,
    title="Reedoy Payroll",
):
    if Workbook is None:
        abort(
            503,
            description=
            "openpyxl is not installed.",
        )

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Payroll"

    if rows:
        headers = list(
            rows[0].keys()
        )

        worksheet.append(headers)

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

    output = io.BytesIO()

    workbook.save(output)

    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


@app.route(
    "/export/workers.xlsx"
)
@login_required
def export_workers():
    rows = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    return export_rows(
        rows,
        "reedoy_workers.xlsx",
    )


@app.route(
    "/export/advances.xlsx"
)
@login_required
def export_advances():
    rows = advance_rows(
        month_name_year()
    )

    return export_rows(
        rows,
        "reedoy_advances.xlsx",
    )


@app.route(
    "/export/department.xlsx"
)
@login_required
def export_department():
    month = month_name_year()

    selected_department = request.args.get(
        "department",
        "",
    )

    if selected_department:
        worker_rows = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (
                selected_department,
            ),
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
        month,
    )

    rows = []

    for worker in worker_rows:
        salary = salary_map.get(
            worker["id"],
            {},
        )

        rows.append(
            {
                "ID":
                    worker["id"],
                "Worker Name":
                    worker.get("name"),
                "Department":
                    worker.get(
                        "department"
                    ),
                "Designation":
                    worker.get(
                        "designation"
                    ),
                **salary,
            }
        )

    return export_rows(
        rows,
        "reedoy_department_salary.xlsx",
    )


# ============================================================
# CSV EXPORT
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
            ),
        )

        writer.writeheader()
        writer.writerows(rows)

    return send_file(
        io.BytesIO(
            output.getvalue()
            .encode("utf-8-sig")
        ),
        as_attachment=True,
        download_name=
            "reedoy_workers.csv",
        mimetype="text/csv",
    )


# ============================================================
# PDF PAYSLIP
# ============================================================

@app.route(
    "/export/payslip.pdf"
)
@login_required
def export_payslip_pdf():
    if SimpleDocTemplate is None:
        abort(
            503,
            description=
            "reportlab is not installed.",
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
        (worker_id,),
    )

    if not worker:
        abort(404)

    salary = calculate_salary(
        worker,
        month,
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

    styles = getSampleStyleSheet()

    story = [
        Paragraph(
            get_settings().get(
                "company_name",
                DEFAULT_COMPANY_NAME,
            ),
            styles["Title"],
        ),
        Paragraph(
            "Payroll Payslip - "
            + month,
            styles["Heading2"],
        ),
        Spacer(1, 12),
    ]

    data = [
        [
            "Worker ID",
            worker["id"],
        ],
        [
            "Worker Name",
            worker.get("name", ""),
        ],
        [
            "Department",
            worker.get(
                "department",
                "",
            ),
        ],
        [
            "Basic Salary",
            "BDT "
            + f"{salary['earned_basic']:,.2f}",
        ],
        [
            "Absent Deduction",
            "BDT "
            + f"{salary['absent_cut']:,.2f}",
        ],
        [
            "OT",
            "BDT "
            + f"{salary['ot_amt']:,.2f}",
        ],
        [
            "Nasta",
            "BDT "
            + f"{salary['nasta']:,.2f}",
        ],
        [
            "Gross",
            "BDT "
            + f"{salary['gross']:,.2f}",
        ],
        [
            "Advance",
            "BDT "
            + f"{salary['advance']:,.2f}",
        ],
        [
            "Net Payable",
            "BDT "
            + f"{salary['net']:,.2f}",
        ],
    ]

    story.append(
        Table(
            data,
            colWidths=[
                180,
                280,
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

    document.build(story)

    output.seek(0)

    filename = (
        f"payslip_{worker['id']}_"
        f"{month.replace(' ', '_')}.pdf"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/pdf",
    )


# ============================================================
# COMPATIBILITY ROUTES
#
# These are IMPORTANT because older HTML files may use
# old endpoint names.
# ============================================================

# Advance Salary
app.add_url_rule(
    "/advances",
    endpoint="advances",
    view_func=advance,
    methods=["GET"],
)

app.add_url_rule(
    "/advance-salary",
    endpoint="advance_salary",
    view_func=advance,
    methods=["GET"],
)

app.add_url_rule(
    "/advance_salary",
    endpoint="advance_salary_page",
    view_func=advance,
    methods=["GET"],
)


# Workers
app.add_url_rule(
    "/worker-management",
    endpoint="worker_management",
    view_func=workers,
    methods=["GET"],
)


# Attendance
app.add_url_rule(
    "/attendance-calendar",
    endpoint="attendance_calendar",
    view_func=attendance,
    methods=["GET"],
)


# Payslip
app.add_url_rule(
    "/single-payslip",
    endpoint="single_payslip",
    view_func=payslip,
    methods=["GET"],
)


# Department
app.add_url_rule(
    "/department-salary-sheet",
    endpoint="department_salary_sheet",
    view_func=department,
    methods=["GET"],
)


# Old report endpoint
app.add_url_rule(
    "/reports",
    endpoint="reports",
    view_func=department,
    methods=["GET"],
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
                error=error,
            ),
            500,
        )
    except Exception:
        return (
            "500 - Internal Server Error\n"
            + str(error),
            500,
        )


# ============================================================
# STARTUP
# ============================================================

try:
    init_db()

except Exception as error:
    app.logger.exception(
        "Database initialization error: %r",
        error,
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
