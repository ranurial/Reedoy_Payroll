import os
import csv
import io
import sqlite3
import hashlib
import calendar
import datetime
import json
import zipfile
import base64
from decimal import Decimal
from functools import wraps

from jinja2 import TemplateNotFound

from flask import (
    Flask,
    render_template,
    render_template_string,
    request,
    redirect,
    url_for,
    session,
    flash,
    get_flashed_messages,
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
    colors = None
    getSampleStyleSheet = None
    A4 = None
    landscape = None


# ============================================================
# APPLICATION
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


# Local/offline SQLite database: always use the final source-of-truth
# factory_payroll.db beside this application file unless SQLITE_DB is
# explicitly supplied. This prevents the app from silently opening a new
# empty reedoy_payroll.db and showing 0 workers.
_DEFAULT_SQLITE_DB = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "factory_payroll.db",
)
DB_PATH = os.environ.get(
    "SQLITE_DB",
    _DEFAULT_SQLITE_DB,
).strip() or _DEFAULT_SQLITE_DB


DEFAULT_COMPANY_NAME = (
    "REEDOY TEXTILE DYEING PRINTING & FINISHING"
)


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
# FIXED DEPARTMENTS
# ============================================================

DEPARTMENTS = [
    "All Departments",
    "General",
    "Printing",
    "Jigar",
    "Wash",
    "Loop",
    "Stanter",
    "Electrical",
    "Accounts",
    "Design",
]


DEPT_BN = {
    "All Departments": "সকল বিভাগ",
    "General": "সাধারণ",
    "Printing": "প্রিন্টিং",
    "Jigar": "জিগার",
    "Wash": "ওয়াশ",
    "Loop": "লুপ",
    "Stanter": "স্ট্যান্টার",
    "Electrical": "ইলেকট্রিক্যাল",
    "Accounts": "অ্যাকাউন্টস",
    "Design": "ডিজাইন",
}


def get_worker_departments():
    """
    Return fixed department list.

    This function does not modify the database.
    """
    return DEPARTMENTS.copy()


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
    "Attendance": "উপস্থিতি",
    "Advances": "অগ্রিম",
    "Payroll Report": "বেতন প্রতিবেদন",
    "Users": "ব্যবহারকারী",
    "Activity Log": "কার্যক্রমের লগ",
    "Company Settings": "কোম্পানির সেটিংস",
    "Bangla": "বাংলা",
    "English": "ইংরেজি",
    "Logout": "লগআউট",
    "Quick Links": "দ্রুত লিংক",
    "Salary Report": "বেতন প্রতিবেদন",
    "Add Advance": "অগ্রিম যোগ করুন",
    "Worker ID": "কর্মী আইডি",
    "English Name": "ইংরেজি নাম",
    "Bangla Name": "বাংলা নাম",
    "Name": "নাম",
    "Basic Salary (BDT)": "মূল বেতন (টাকা)",
    "Refreshment Bill": "নাস্তার বিল",
    "Edit": "সম্পাদনা",
    "Actions": "কার্যক্রম",
    "Month": "মাস",
    "Year": "বছর",
    "Select Worker": "কর্মী নির্বাচন করুন",
    "Select Month": "মাস নির্বাচন করুন",
    "Select Year": "বছর নির্বাচন করুন",
    "Status": "অবস্থা",
    "OT Hours": "ওটি ঘণ্টা",
    "Total Present": "মোট উপস্থিত",
    "Total Absent": "মোট অনুপস্থিত",
    "Salary Summary": "বেতনের সারসংক্ষেপ",
    "No workers found.": "কোনো কর্মী পাওয়া যায়নি।",
    "Save Worker": "কর্মী সংরক্ষণ করুন",
    "Add Worker": "কর্মী যোগ করুন",
    "Update Worker": "কর্মীর তথ্য আপডেট করুন",
    "Worker name is required.": "কর্মীর নাম আবশ্যক।",
    "Worker saved successfully.": "কর্মীর তথ্য সফলভাবে সংরক্ষিত হয়েছে।",
}


# ============================================================
# DATABASE TYPE
# ============================================================

def is_postgres():
    return bool(
        DATABASE_URL
        and not DATABASE_URL.startswith("sqlite://")
    )


# ============================================================
# POSTGRES CONNECTION POOL
# ============================================================

_PG_POOL = None


def _get_pg_pool():
    global _PG_POOL

    if _PG_POOL is None:

        try:
            import psycopg2
            from psycopg2.pool import ThreadedConnectionPool
        except Exception as e:
            raise RuntimeError(
                "psycopg2 is required for PostgreSQL. "
                "Add psycopg2-binary to requirements.txt."
            ) from e

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
        DB_PATH,
        timeout=30,
        check_same_thread=False,
    )

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


# ============================================================
# SQL PLACEHOLDER
# ============================================================

def placeholders(sql):

    if is_postgres():
        return sql.replace("?", "%s")

    return sql


# ============================================================
# GENERIC EXECUTE
# ============================================================

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
            cur.executemany(
                sql2,
                params
            )
        else:
            cur.execute(
                sql2,
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


# ============================================================
# ROW CONVERSION
# ============================================================

def row_dict(cur, row):

    if row is None:
        return None

    if hasattr(row, "keys"):
        return dict(row)

    return {
        description[0]: row[index]
        for index, description
        in enumerate(cur.description)
    }


# ============================================================
# FETCH ALL
# ============================================================

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
            row_dict(cur, row)
            for row in rows
        ]

    finally:

        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)


# ============================================================
# FETCH ONE
# ============================================================

def fetch_one(sql, params=()):

    rows = fetch_all(
        sql,
        params
    )

    return rows[0] if rows else None


# ============================================================
# SCALAR
# ============================================================

def scalar(
    sql,
    params=(),
    default=0,
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


# ============================================================
# TABLE EXISTS
# ============================================================

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


# ============================================================
# TABLE COLUMNS
# ============================================================

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
            row["column_name"]
            for row in rows
        }

    conn = db_connect()
    cur = conn.cursor()

    try:

        cur.execute(
            "PRAGMA table_info(" + name + ")"
        )

        return {
            row[1]
            for row in cur.fetchall()
        }

    finally:

        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)


# ============================================================
# ADD COLUMN
# ============================================================

def add_column(
    name,
    col,
    typ,
):

    if not table_exists(name):
        return

    if col in columns(name):
        return

    execute(
        f"ALTER TABLE {name} ADD COLUMN {col} {typ}",
        commit=True,
    )


# ============================================================
# PASSWORD
# ============================================================

def hash_password(password):

    return hashlib.sha256(
        str(password).encode("utf-8")
    ).hexdigest()


# ============================================================
# CURRENT DATETIME
# ============================================================

def nowstr():

    return datetime.datetime.now().strftime(
        "%Y-%m-%d %H:%M:%S"
    )


# ============================================================
# DATABASE INITIALIZATION
# ============================================================

def init_db():
    """
    Create only missing structures.

    IMPORTANT:
    - Never DROP tables.
    - Never TRUNCATE tables.
    - Never delete migrated workers.
    - Never re-import worker records.
    - Existing worker IDs remain untouched.
    """

    conn = db_connect()
    cur = conn.cursor()

    try:

        # ----------------------------------------------------
        # Workers
        # ----------------------------------------------------

        if is_postgres():

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS workers (
                    id SERIAL PRIMARY KEY,
                    name TEXT NOT NULL,
                    basic_salary DOUBLE PRECISION NOT NULL DEFAULT 0,
                    ot_rate DOUBLE PRECISION NOT NULL DEFAULT 0,
                    department TEXT,
                    designation TEXT,
                    refreshment_bill DOUBLE PRECISION NOT NULL DEFAULT 0,
                    bangla_name TEXT
                )
                """
            )

        else:

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS workers (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
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

        # ----------------------------------------------------
        # Attendance
        # ----------------------------------------------------

        if is_postgres():

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    present_days INTEGER DEFAULT 0,
                    absent_days INTEGER DEFAULT 0,
                    ot_hours DOUBLE PRECISION DEFAULT 0,
                    advance_deduction DOUBLE PRECISION DEFAULT 0
                )
                """
            )

        else:

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS attendance (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    present_days INTEGER DEFAULT 0,
                    absent_days INTEGER DEFAULT 0,
                    ot_hours REAL DEFAULT 0,
                    advance_deduction REAL DEFAULT 0
                )
                """
            )

        # ----------------------------------------------------
        # Daily Attendance
        # ----------------------------------------------------

        if is_postgres():

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS daily_attendance (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    day INTEGER NOT NULL,
                    status TEXT NOT NULL DEFAULT 'P'
                )
                """
            )

        else:

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS daily_attendance (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER,
                    month_year TEXT NOT NULL,
                    day INTEGER NOT NULL,
                    status TEXT NOT NULL DEFAULT 'P'
                )
                """
            )

        # ----------------------------------------------------
        # Company Settings
        # ----------------------------------------------------

        cur.execute(
            """
            CREATE TABLE IF NOT EXISTS company_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """
        )

        # ----------------------------------------------------
        # Users
        # ----------------------------------------------------

        if is_postgres():

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
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

        else:

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
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

        # ----------------------------------------------------
        # Payroll Payments
        # ----------------------------------------------------

        if is_postgres():
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS payroll_payments (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER NOT NULL,
                    month_year TEXT NOT NULL,
                    net_payable DOUBLE PRECISION DEFAULT 0,
                    paid_amount DOUBLE PRECISION DEFAULT 0,
                    payment_date TEXT,
                    payment_method TEXT DEFAULT 'Cash',
                    status TEXT DEFAULT 'Unpaid',
                    note TEXT,
                    created_at TEXT,
                    updated_at TEXT,
                    UNIQUE(worker_id, month_year)
                )
                """
            )
        else:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS payroll_payments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER NOT NULL,
                    month_year TEXT NOT NULL,
                    net_payable REAL DEFAULT 0,
                    paid_amount REAL DEFAULT 0,
                    payment_date TEXT,
                    payment_method TEXT DEFAULT 'Cash',
                    status TEXT DEFAULT 'Unpaid',
                    note TEXT,
                    created_at TEXT,
                    updated_at TEXT,
                    UNIQUE(worker_id, month_year)
                )
                """
            )

        # ----------------------------------------------------
        # Payment Transactions / History
        # ----------------------------------------------------

        if is_postgres():
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS payment_transactions (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER NOT NULL,
                    month_year TEXT NOT NULL,
                    amount DOUBLE PRECISION DEFAULT 0,
                    payment_date TEXT,
                    payment_method TEXT DEFAULT 'Cash',
                    note TEXT,
                    created_at TEXT
                )
                """
            )
        else:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS payment_transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER NOT NULL,
                    month_year TEXT NOT NULL,
                    amount REAL DEFAULT 0,
                    payment_date TEXT,
                    payment_method TEXT DEFAULT 'Cash',
                    note TEXT,
                    created_at TEXT
                )
                """
            )

        # ----------------------------------------------------
        # Payroll Locks / Monthly Closing
        # ----------------------------------------------------
        cur.execute(
            """CREATE TABLE IF NOT EXISTS payroll_locks (
                month_year TEXT PRIMARY KEY,
                locked_at TEXT,
                locked_by TEXT,
                note TEXT
            )"""
        )

        # ----------------------------------------------------
        # Activity Log
        # ----------------------------------------------------

        if is_postgres():

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS activity_log (
                    id SERIAL PRIMARY KEY,
                    username TEXT,
                    action TEXT,
                    log_time TEXT NOT NULL
                )
                """
            )

        else:

            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS activity_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT,
                    action TEXT,
                    log_time TEXT NOT NULL
                )
                """
            )

        conn.commit()

    finally:

        try:
            cur.close()
        except Exception:
            pass

        db_release(conn)

    # ========================================================
    # ADDITIVE MIGRATION ONLY
    # ========================================================

    worker_columns = [
        ("phone", "TEXT"),
        ("address", "TEXT"),
        ("joining_date", "TEXT"),
        ("status", "TEXT"),
        ("bangla_name", "TEXT"),
    ]

    for col, typ in worker_columns:

        try:
            add_column(
                "workers",
                col,
                typ
            )
        except Exception as e:
            app.logger.warning(
                "Worker column migration %s: %r",
                col,
                e
            )

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

    try:
        add_column(
            "users",
            "last_login",
            "TEXT"
        )
    except Exception:
        pass

    # ========================================================
    # ADMIN USER
    # ========================================================

    try:

        admin = fetch_one(
            """
            SELECT id
            FROM users
            WHERE LOWER(username)=?
            LIMIT 1
            """,
            ("admin",),
        )

        if not admin:

            user_columns = columns(
                "users"
            )

            names = []
            values = []

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

            for name, value in data:

                if name in user_columns:

                    names.append(name)
                    values.append(value)

            if names:

                placeholders_list = ",".join(
                    ["?"] * len(values)
                )

                execute(
                    """
                    INSERT INTO users
                    (""" + ",".join(names) + """)
                    VALUES
                    (""" + placeholders_list + """)
                    """,
                    values,
                    commit=True,
                )

    except Exception as e:

        app.logger.warning(
            "Admin initialization error: %r",
            e
        )

    # ========================================================
    # DEFAULT COMPANY SETTINGS
    # ========================================================

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
                    INSERT INTO company_settings
                    (key,value)
                    VALUES (?,?)
                    ON CONFLICT(key)
                    DO NOTHING
                    """,
                    (key, value),
                    commit=True,
                )

            else:

                execute(
                    """
                    INSERT OR IGNORE INTO company_settings
                    (key,value)
                    VALUES (?,?)
                    """,
                    (key, value),
                    commit=True,
                )

    except Exception as e:

        app.logger.warning(
            "Settings initialization error: %r",
            e
        )



# ============================================================
# LOGIN REQUIRED
# ============================================================
#
# IMPORTANT:
# This decorator must be defined before any route uses
# @login_required. The v48 sync build had it defined later
# in the file, which caused a NameError during startup.
# ============================================================

def login_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        if not session.get("user_id"):

            return redirect(
                url_for(
                    "login",
                    next=request.path
                )
            )

        return function(
            *args,
            **kwargs
        )

    return wrapper


# ============================================================
# ADVANCE SALARY EXPORTS
# ============================================================

@app.route("/advance/export")
@app.route("/advances/export")
@login_required
def advance_export():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    worker_id = request.args.get("worker_id") or ""
    rows = advance_rows(month, worker_id)
    if openpyxl is None or Workbook is None:
        flash("Excel export requires openpyxl.", "danger")
        return redirect(url_for("advance", month=month))

    wb = Workbook()
    ws = wb.active
    ws.title = "Advance Salary"
    ws.append(["REEDOY TEXTILE DYEING PRINTING & FINISHING"])
    ws.append([f"Advance Salary Report - {month}"])
    ws.append([])
    ws.append(["ID", "Worker ID", "English Name", "Bangla Name", "Department", "Date", "Amount (BDT)", "Note"])
    total = 0.0
    for r in rows:
        amount = parse_num(r.get("amount"), 0)
        total += amount
        ws.append([r.get("id"), r.get("worker_id"), r.get("worker_name") or "", r.get("bangla_name") or "", r.get("department") or "", r.get("advance_date") or "", amount, r.get("note") or ""])
    ws.append([])
    ws.append(["", "", "", "", "", "TOTAL", total, ""])
    for col in ws.columns:
        max_len = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 12), 35)
    ws.freeze_panes = "A5"
    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    safe_month = month.replace(" ", "_")
    return send_file(bio, as_attachment=True, download_name=f"Advance_Salary_{safe_month}.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/advance/pdf")
@app.route("/advances/pdf")
@login_required
def advance_pdf():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    worker_id = request.args.get("worker_id") or ""
    rows = advance_rows(month, worker_id)
    if SimpleDocTemplate is None:
        flash("PDF export requires reportlab.", "danger")
        return redirect(url_for("advance", month=month))

    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=landscape(A4), rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    story = [Paragraph("REEDOY TEXTILE DYEING PRINTING & FINISHING", styles["Title"]), Paragraph(f"Advance Salary Report - {month}", styles["Heading2"]), Spacer(1, 10)]
    data = [["ID", "Worker ID", "English Name", "Department", "Date", "Amount (BDT)", "Note"]]
    total = 0.0
    for r in rows:
        amount = parse_num(r.get("amount"), 0)
        total += amount
        data.append([str(r.get("id") or ""), str(r.get("worker_id") or ""), str(r.get("worker_name") or ""), str(r.get("department") or ""), str(r.get("advance_date") or ""), f"{amount:,.2f}", str(r.get("note") or "")])
    data.append(["", "", "", "", "TOTAL", f"{total:,.2f}", ""])
    table = Table(data, repeatRows=1, colWidths=[45, 60, 180, 100, 80, 90, 180])
    if TableStyle is not None and colors is not None:
        table.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1764c0")),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.4, colors.grey),
            ("ALIGN", (5,1), (5,-1), "RIGHT"),
            ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
            ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#eaf2ff")),
        ]))
    story.append(table)
    doc.build(story)
    bio.seek(0)
    safe_month = month.replace(" ", "_")
    return send_file(bio, as_attachment=True, download_name=f"Advance_Salary_{safe_month}.pdf", mimetype="application/pdf")


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

        result = {}

        for row in rows:
            result[
                row["key"]
            ] = row["value"]

        if "company_name" not in result:
            result["company_name"] = (
                DEFAULT_COMPANY_NAME
            )

        return result

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
            INSERT INTO activity_log
            (username,action,log_time)
            VALUES (?,?,?)
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
            e
        )


# ============================================================
# CURRENT USER
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
# ADMIN REQUIRED
# ============================================================

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
                "danger"
            )

            return redirect(
                url_for("dashboard")
            )

        return function(
            *args,
            **kwargs
        )

    return wrapper


# ============================================================
# MONTH
# ============================================================

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


# ============================================================
# NUMBER PARSER
# ============================================================

def parse_num(
    value,
    default=0.0
):

    try:

        if value is None:
            return default

        text = str(value).strip()

        if text == "":
            return default

        return float(text)

    except Exception:

        return default


# ============================================================
# MONTH FROM DATE
# ============================================================

def month_name_from_date(value):

    try:

        text = str(value)[:10]

        year, month, day = (
            text.split("-")
        )

        return (
            f"{MONTHS[int(month)-1]} "
            f"{int(year)}"
        )

    except Exception:

        return ""


# ============================================================
# LEGACY ADVANCE TABLE
# ============================================================

def legacy_advance_source():

    if table_exists(
        "worker_advances"
    ):

        cols = columns(
            "worker_advances"
        )

        if {
            "worker_id",
            "amount",
        }.issubset(cols):

            return "worker_advances"

    if table_exists(
        "advance_salary"
    ):

        return "advance_salary"

    if table_exists(
        "advances"
    ):

        return "advances"

    return None


# ============================================================
# ADVANCE ROWS
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
    # Modern worker_advances
    # --------------------------------------------------------

    if table_name == "worker_advances":

        where = []
        params = []

        if month:

            where.append(
                "month_year=?"
            )

            params.append(month)

        if worker_id:

            where.append(
                "worker_id=?"
            )

            params.append(worker_id)

        query = f"""
            SELECT
                wa.id,
                wa.worker_id,
                wa.month_year,
                wa.advance_date,
                wa.amount,
                {("wa.note" if "note" in table_columns else "''")} AS note,
                w.name AS worker_name,
                w.bangla_name,
                w.department
            FROM worker_advances wa
            LEFT JOIN workers w
                ON w.id = wa.worker_id
        """

        if where:

            query += (
                " WHERE "
                + " AND ".join(where)
            )

        query += " ORDER BY wa.id DESC"

        return fetch_all(
            query,
            params
        )

    # --------------------------------------------------------
    # Legacy advance_salary
    # --------------------------------------------------------

    if table_name == "advance_salary":

        if "date" in table_columns:

            date_column = "date"

        elif "advance_date" in table_columns:

            date_column = "advance_date"

        else:

            date_column = None

        query = f"""
            SELECT
                a.id,
                a.worker_id,
                {
                    "a." + date_column
                    if date_column
                    else "NULL"
                } AS advance_date,
                a.amount,
                {("a.note" if "note" in table_columns else "''")} AS note,
                w.name AS worker_name,
                w.bangla_name,
                w.department
            FROM advance_salary a
            LEFT JOIN workers w
                ON w.id = a.worker_id
        """

        rows = fetch_all(query)

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
                ) == month
            ]

        if worker_id:

            rows = [
                row
                for row in rows
                if str(
                    row.get("worker_id")
                ) == str(worker_id)
            ]

        return rows

    # --------------------------------------------------------
    # Legacy advances
    # --------------------------------------------------------

    date_column = (
        "date"
        if "date" in table_columns
        else (
            "advance_date"
            if "advance_date" in table_columns
            else None
        )
    )
    note_expr = "a.note" if "note" in table_columns else "''"
    date_expr = f"a.{date_column}" if date_column else "NULL"

    query = f"""
        SELECT
            a.id,
            a.worker_id,
            {date_expr} AS advance_date,
            a.amount,
            {note_expr} AS note,
            w.name AS worker_name,
            w.bangla_name,
            w.department
        FROM {table_name} a
        LEFT JOIN workers w
            ON w.id = a.worker_id
    """

    rows = fetch_all(query)

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
            ) == month
        ]

    if worker_id:

        rows = [
            row
            for row in rows
            if str(
                row.get("worker_id")
            ) == str(worker_id)
        ]

    return rows


# ============================================================
# SAVE ADVANCE
# ============================================================

def save_advance_record(
    worker_id,
    month,
    advance_date,
    amount,
    note
):

    table_name = (
        legacy_advance_source()
    )

    # --------------------------------------------------------
    # Existing worker_advances
    # --------------------------------------------------------

    if table_name == "worker_advances":

        table_columns = columns(
            "worker_advances"
        )

        names = []
        values = []

        data = [
            (
                "worker_id",
                worker_id
            ),
            (
                "month_year",
                month
            ),
            (
                "advance_date",
                advance_date
            ),
            (
                "amount",
                amount
            ),
            (
                "note",
                note
            ),
        ]

        for name, value in data:

            if name in table_columns:

                names.append(name)
                values.append(value)

        execute(
            """
            INSERT INTO worker_advances
            (""" + ",".join(names) + """)
            VALUES
            (""" + ",".join(
                ["?"] * len(values)
            ) + """)
            """,
            values,
            commit=True,
        )

        return

    # --------------------------------------------------------
    # Existing advance_salary
    # --------------------------------------------------------

    if table_name == "advance_salary":

        table_columns = columns(
            "advance_salary"
        )

        date_column = (
            "date"
            if "date" in table_columns
            else (
                "advance_date"
                if "advance_date"
                in table_columns
                else None
            )
        )

        if not date_column:
            raise RuntimeError(
                "advance_salary has no date column."
            )

        names = [
            "worker_id",
            date_column,
            "amount",
        ]

        values = [
            worker_id,
            advance_date,
            amount,
        ]

        if "note" in table_columns:

            names.append("note")
            values.append(note)

        execute(
            """
            INSERT INTO advance_salary
            (""" + ",".join(names) + """)
            VALUES
            (""" + ",".join(
                ["?"] * len(values)
            ) + """)
            """,
            values,
            commit=True,
        )

        return

    # --------------------------------------------------------
    # Existing advances
    # --------------------------------------------------------

    if table_name == "advances":

        table_columns = columns(
            "advances"
        )

        if (
            "worker_id" in table_columns
            and "date" in table_columns
            and "amount" in table_columns
        ):

            names = [
                "worker_id",
                "date",
                "amount",
            ]

            values = [
                worker_id,
                advance_date,
                amount,
            ]

            if "note" in table_columns:

                names.append("note")
                values.append(note)

            execute(
                """
                INSERT INTO advances
                (""" + ",".join(names) + """)
                VALUES
                (""" + ",".join(
                    ["?"] * len(values)
                ) + """)
                """,
                values,
                commit=True,
            )

            return

    # --------------------------------------------------------
    # No advance table exists.
    # Create only new table.
    # --------------------------------------------------------

    if is_postgres():

        execute(
            """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount DOUBLE PRECISION NOT NULL,
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
        VALUES
        (?,?,?,?,?)
        """,
        (
            worker_id,
            month,
            advance_date,
            amount,
            note,
        ),
        commit=True,
    )


# ============================================================
# EDIT ADVANCE
# ============================================================

@app.route(
    "/advance/edit/<int:aid>",
    methods=["GET", "POST"]
)
@app.route(
    "/advances/edit/<int:aid>",
    methods=["GET", "POST"]
)
@login_required
def edit_advance(aid):

    rows = advance_rows()

    record = None

    for row in rows:
        if str(row.get("id")) == str(aid):
            record = row
            break

    if not record:
        flash(
            "Advance record not found.",
            "danger"
        )
        return redirect(
            url_for("advance")
        )

    if request.method == "GET":

        return safe_render_template(
            "advance_edit.html",
            record=record,
            workers=fetch_all(
                """
                SELECT
                    id,
                    name,
                    bangla_name,
                    department
                FROM workers
                ORDER BY id
                """
            ),
            language=session.get("language", "en"),
            settings=get_settings(),
        )

    form = request.form

    try:

        worker_id = int(
            form.get("worker_id")
        )

    except Exception:

        flash(
            "Invalid worker.",
            "danger"
        )

        return redirect(
            url_for(
                "edit_advance",
                aid=aid
            )
        )

    month = (
        form.get("month_year")
        or month_name_year(
            form.get("month"),
            form.get("year")
        )
        or record.get("month_year")
    )

    if not require_payroll_unlocked(month):
        return redirect(url_for("advance", month=month))

    advance_date = (
        form.get("advance_date")
        or record.get("advance_date")
        or datetime.date.today().isoformat()
    )

    amount = parse_num(
        form.get("amount"),
        0
    )

    note = (
        form.get("note", "")
        .strip()
    )

    if amount <= 0:

        flash(
            "Amount must be greater than zero.",
            "danger"
        )

        return redirect(
            url_for(
                "edit_advance",
                aid=aid
            )
        )

    try:

        update_advance_record(
            aid,
            worker_id,
            month,
            advance_date,
            amount,
            note,
        )

        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Advance updated: ID {aid}, Worker {worker_id}, {month}"
        )
        flash(
            "Advance updated successfully.",
            "success"
        )

    except Exception as e:

        app.logger.exception(
            "Advance update error"
        )

        flash(
            f"Could not update advance: {e}",
            "danger"
        )

        return redirect(
            url_for(
                "edit_advance",
                aid=aid
            )
        )

    return redirect(
        url_for(
            "advance",
            month=month
        )
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

    if not table_name:
        return

    execute(
        f"""
        DELETE FROM {table_name}
        WHERE id=?
        """,
        (advance_id,),
        commit=True,
    )


# ============================================================
# UPDATE ADVANCE
# ============================================================

def update_advance_record(
    advance_id,
    worker_id,
    month,
    advance_date,
    amount,
    note
):
    table_name = legacy_advance_source()

    if not table_name:
        raise RuntimeError(
            "No advance table exists."
        )

    table_columns = columns(
        table_name
    )

    if "worker_id" not in table_columns:
        raise RuntimeError(
            f"{table_name} has no worker_id column."
        )

    if "amount" not in table_columns:
        raise RuntimeError(
            f"{table_name} has no amount column."
        )

    # --------------------------------------------------------
    # Modern worker_advances
    # --------------------------------------------------------

    if table_name == "worker_advances":

        fields = []
        values = []

        if "worker_id" in table_columns:
            fields.append("worker_id=?")
            values.append(worker_id)

        if "month_year" in table_columns:
            fields.append("month_year=?")
            values.append(month)

        if "advance_date" in table_columns:
            fields.append("advance_date=?")
            values.append(advance_date)

        if "amount" in table_columns:
            fields.append("amount=?")
            values.append(amount)

        if "note" in table_columns:
            fields.append("note=?")
            values.append(note)

        values.append(advance_id)

        execute(
            """
            UPDATE worker_advances
            SET
            """ + ",".join(fields) + """
            WHERE id=?
            """,
            values,
            commit=True,
        )

        return

    # --------------------------------------------------------
    # Legacy advance_salary
    # --------------------------------------------------------

    if table_name == "advance_salary":

        date_column = (
            "date"
            if "date" in table_columns
            else (
                "advance_date"
                if "advance_date" in table_columns
                else None
            )
        )

        if not date_column:
            raise RuntimeError(
                "advance_salary has no date column."
            )

        fields = [
            "worker_id=?",
            f"{date_column}=?",
            "amount=?",
        ]

        values = [
            worker_id,
            advance_date,
            amount,
        ]

        if "note" in table_columns:
            fields.append("note=?")
            values.append(note)

        values.append(advance_id)

        execute(
            """
            UPDATE advance_salary
            SET
            """ + ",".join(fields) + """
            WHERE id=?
            """,
            values,
            commit=True,
        )

        return

    # --------------------------------------------------------
    # Legacy advances
    # --------------------------------------------------------

    if table_name == "advances":

        date_column = (
            "date"
            if "date" in table_columns
            else (
                "advance_date"
                if "advance_date" in table_columns
                else None
            )
        )

        if not date_column:
            raise RuntimeError(
                "advances has no date column."
            )

        fields = [
            "worker_id=?",
            f"{date_column}=?",
            "amount=?",
        ]

        values = [
            worker_id,
            advance_date,
            amount,
        ]

        if "note" in table_columns:
            fields.append("note=?")
            values.append(note)

        values.append(advance_id)

        execute(
            """
            UPDATE advances
            SET
            """ + ",".join(fields) + """
            WHERE id=?
            """,
            values,
            commit=True,
        )

        return


# ============================================================
# DAILY ATTENDANCE MAP
# ============================================================

def daily_map(worker_id, month):
    month_text = str(month or "").strip()

    year = None
    month_number = None

    # Accept YYYY-MM
    parts = month_text.split("-")
    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
        try:
            year = int(parts[0])
            month_number = int(parts[1])
        except Exception:
            year = None
            month_number = None

    # Accept "October 2026"
    if year is None:
        for fmt in ("%B %Y", "%b %Y"):
            try:
                parsed = datetime.datetime.strptime(month_text, fmt)
                year = parsed.year
                month_number = parsed.month
                break
            except Exception:
                pass

    # Always use database format YYYY-MM for the query
    db_month = month_text
    if year is not None and month_number is not None:
        db_month = f"{year:04d}-{month_number:02d}"

    rows = fetch_all(
        "SELECT day, status FROM daily_attendance "
        "WHERE worker_id=? AND month_year=?",
        (worker_id, db_month)
    )

    result = {}

    # Preserve all stored non-Friday attendance
    for row in rows:
        try:
            day_number = int(row["day"])
        except Exception:
            continue

        status = row["status"]

        # Friday is ALWAYS A
        if year is not None and month_number is not None:
            try:
                current_date = datetime.date(year, month_number, day_number)
                if current_date.weekday() == 4:
                    result[day_number] = "A"
                    continue
            except Exception:
                pass

        result[day_number] = str(status or "")

    # Add Friday=A even when there is no database attendance row
    if year is not None and month_number is not None:
        for day_number in range(1, 32):
            try:
                current_date = datetime.date(year, month_number, day_number)
            except ValueError:
                break

            if current_date.weekday() == 4:
                result[day_number] = "A"

    return result

def calculate_salary_bulk(
    workers,
    month
):

    worker_ids = [
        worker["id"]
        for worker in workers
    ]

    if not worker_ids:
        return {}

    # --------------------------------------------------------
    # Attendance
    # --------------------------------------------------------

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

    daily_by_worker = {}

    for row in daily_rows:

        try:

            worker_id = row.get(
                "worker_id"
            )

            day = int(
                row.get("day")
            )

            status = str(
                row.get("status")
                or "P"
            )

            daily_by_worker.setdefault(
                worker_id,
                {}
            )[day] = status

        except Exception:
            pass

    # --------------------------------------------------------
    # Advances
    # --------------------------------------------------------

    advances_by_worker = {}

    table_name = (
        legacy_advance_source()
    )

    if table_name:

        table_columns = columns(
            table_name
        )

        try:

            if (
                table_name
                == "worker_advances"
                and {
                    "worker_id",
                    "amount",
                    "month_year",
                }.issubset(
                    table_columns
                )
            ):

                advance_data = fetch_all(
                    """
                    SELECT
                        worker_id,
                        amount
                    FROM worker_advances
                    WHERE month_year=?
                    """,
                    (month,),
                )

                for row in advance_data:

                    worker_id = row.get(
                        "worker_id"
                    )

                    advances_by_worker[
                        worker_id
                    ] = (
                        advances_by_worker.get(
                            worker_id,
                            0
                        )
                        + parse_num(
                            row.get("amount")
                        )
                    )

            elif (
                table_name
                in (
                    "advance_salary",
                    "advances",
                )
                and "worker_id"
                in table_columns
                and "amount"
                in table_columns
            ):

                if "date" in table_columns:

                    date_column = "date"

                elif (
                    "advance_date"
                    in table_columns
                ):

                    date_column = (
                        "advance_date"
                    )

                else:

                    date_column = None

                if date_column:

                    advance_data = fetch_all(
                        f"""
                        SELECT
                            worker_id,
                            {date_column}
                                AS advance_date,
                            amount
                        FROM {table_name}
                        """
                    )

                    for row in advance_data:

                        if (
                            month_name_from_date(
                                row.get(
                                    "advance_date"
                                )
                            )
                            == month
                        ):

                            worker_id = row.get(
                                "worker_id"
                            )

                            advances_by_worker[
                                worker_id
                            ] = (
                                advances_by_worker.get(
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

    # --------------------------------------------------------
    # Calculate
    # --------------------------------------------------------

    result = {}

    try:

        month_name, year_text = (
            month.split()
        )

        year = int(year_text)

        month_number = (
            MONTHS.index(
                month_name
            ) + 1
        )

        days_in_month = 30

    except Exception:

        year = datetime.date.today().year
        month_number = (
            datetime.date.today().month
        )

        days_in_month = 30

    for worker in workers:

        worker_id = worker["id"]

        attendance = (
            attendance_by_worker.get(
                worker_id,
                {}
            )
        )

        absent = min(
            30,
            max(0, int(attendance.get("absent_days") or 0))
        )

        # Salary attendance is on a fixed 30-day basis.
        # Friday is a weekly holiday, so it is not an absence.
        # Payable/present days are therefore 30 minus actual absences,
        # capped at 30 regardless of whether the calendar month has
        # 28, 29, 30 or 31 days.
        present = max(
            0,
            30 - absent
        )

        overtime = parse_num(
            attendance.get(
                "ot_hours"
            ),
            0
        )

        basic_salary = parse_num(
            worker.get(
                "basic_salary"
            ),
            0
        )

        ot_rate = parse_num(
            worker.get(
                "ot_rate"
            ),
            0
        )

        nasta_rate = parse_num(
            worker.get(
                "refreshment_bill"
            ),
            0
        )

        # ----------------------------------------------------
        # Absent deduction
        # ----------------------------------------------------

        absent_cut = (
            basic_salary
            / days_in_month
            * absent
            if days_in_month
            else 0
        )

        earned_basic = max(
            0,
            basic_salary - absent_cut
        )

        # ----------------------------------------------------
        # OT
        # ----------------------------------------------------

        ot_amount = (
            overtime * ot_rate
        )

        # ----------------------------------------------------
        # Nasta
        # ----------------------------------------------------

        daily = daily_by_worker.get(
            worker_id,
            {}
        )

        if daily:

            billable_days = 0

            for day in range(
                1,
                days_in_month + 1
            ):

                status = daily.get(
                    day,
                    ""
                )

                try:

                    current_date = (
                        datetime.date(
                            year,
                            month_number,
                            day,
                        )
                    )

                    is_friday = (
                        current_date.weekday()
                        == 4
                    )

                except Exception:

                    is_friday = False

                if (
                    status == "P"
                    and not is_friday
                ):

                    billable_days += 1

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
                ).weekday() != 4
            )

            if days_in_month:

                nasta = round(
                    present
                    * (
                        non_friday_days
                        / days_in_month
                    )
                    * nasta_rate,
                    2,
                )

            else:

                nasta = 0

        # ----------------------------------------------------
        # Advance
        # ----------------------------------------------------

        advance = (
            advances_by_worker.get(
                worker_id,
                0
            )
        )

        # ----------------------------------------------------
        # Gross / Net
        # ----------------------------------------------------

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
            "ot": overtime,
            "absent_cut": absent_cut,
            "earned_basic": earned_basic,
            "ot_amt": ot_amount,
            "nasta": nasta,
            "gross": gross,
            "advance": advance,
            "net": net,
        }

    return result


# ============================================================
# SINGLE SALARY
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

    absent = min(
        30,
        max(0, int(attendance.get("absent_days") or 0))
    )

    # Fixed 30-day salary attendance basis. Weekly holidays are not
    # absences, so a full month with no absence is 30 payable days.
    present = max(
        0,
        30 - absent
    )

    overtime = parse_num(
        attendance.get(
            "ot_hours"
        ),
        0
    )

    try:

        month_name, year_text = (
            month.split()
        )

        year = int(year_text)

        month_number = (
            MONTHS.index(
                month_name
            ) + 1
        )

        days_in_month = 30

    except Exception:

        year = datetime.date.today().year
        month_number = (
            datetime.date.today().month
        )
        days_in_month = 30

    basic_salary = parse_num(
        worker.get(
            "basic_salary"
        ),
        0
    )

    ot_rate = parse_num(
        worker.get(
            "ot_rate"
        ),
        0
    )

    nasta_rate = parse_num(
        worker.get(
            "refreshment_bill"
        ),
        0
    )

    absent_cut = (
        basic_salary
        / days_in_month
        * absent
        if days_in_month
        else 0
    )

    earned_basic = max(
        0,
        basic_salary - absent_cut
    )

    ot_amount = (
        overtime * ot_rate
    )

    daily = daily_map(
        worker["id"],
        month
    )

    if daily:

        billable_days = 0

        for day in range(
            1,
            days_in_month + 1
        ):

            status = daily.get(
                day,
                ""
            )

            current_date = datetime.date(
                year,
                month_number,
                day
            )

            if (
                status == "P"
                and current_date.weekday()
                != 4
            ):

                billable_days += 1

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
            ).weekday() != 4
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

    advances = sum(
        parse_num(
            row.get("amount")
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
        "ot": overtime,
        "absent_cut": absent_cut,
        "earned_basic": earned_basic,
        "ot_amt": ot_amount,
        "nasta": nasta,
        "gross": gross,
        "advance": advances,
        "net": gross - advances,
    }



def salary_for_reporting(worker, month):
    """Calculate salary using saved daily P/A when available.
    Friday is weekly holiday and is excluded from present/absent counts.
    Falls back to the monthly attendance row when no daily records exist.
    """
    monthly = fetch_one(
        "SELECT * FROM attendance WHERE worker_id=? AND month_year=? ORDER BY id DESC LIMIT 1",
        (worker["id"], month),
    ) or {}

    daily = daily_map(worker["id"], month)
    if daily:
        try:
            month_name, year_text = month.split()
            year = int(year_text)
            month_number = MONTHS.index(month_name) + 1
            actual_days = calendar.monthrange(year, month_number)[1]
        except Exception:
            actual_days = 30
            year = datetime.date.today().year
            month_number = datetime.date.today().month

        present = 0
        absent = 0
        for day in range(1, actual_days + 1):
            current_date = datetime.date(year, month_number, day)
            if current_date.weekday() == 4:
                continue
            status = str(daily.get(day, "") or "").upper()
            if status == "P":
                present += 1
            elif status == "A":
                absent += 1

        # For salary/reporting, attendance is paid on a fixed 30-day
        # basis. Actual Friday holidays are excluded from absence, so
        # payable present days are 30 minus actual absences.
        absent = min(30, max(0, absent))
        present = max(0, 30 - absent)

        report_attendance = dict(monthly)
        report_attendance["present_days"] = present
        report_attendance["absent_days"] = absent
        return calculate_salary(worker, month, report_attendance)

    return calculate_salary(worker, month, monthly)

# ============================================================
# PAYROLL LOCK / MONTHLY CLOSING
# ============================================================

def is_payroll_locked(month):
    if not month or not table_exists("payroll_locks"):
        return False
    try:
        return bool(fetch_one("SELECT month_year FROM payroll_locks WHERE month_year=?", (month,)))
    except Exception:
        return False


def payroll_lock_info(month):
    if not month or not table_exists("payroll_locks"):
        return None
    try:
        return fetch_one("SELECT * FROM payroll_locks WHERE month_year=?", (month,))
    except Exception:
        return None


def require_payroll_unlocked(month):
    if is_payroll_locked(month):
        flash(f"Payroll for {month} is locked. Unlock it from Payroll Closing before making changes.", "danger")
        return False
    return True


# ============================================================
# TEMPLATE GLOBALS
# ============================================================

@app.context_processor
def inject_globals():

    settings = get_settings()

    language = session.get(
        "language",
        "en"
    )

    def translate(text):

        if language == "bn":

            return LANG.get(
                text,
                text
            )

        return text

    def department_translate(value):

        value = str(
            value or ""
        )

        if language == "bn":

            return DEPT_BN.get(
                value,
                value
            )

        return value

    return {
        "current_user": current_user(),
        "settings": settings,
        "language": language,
        "tr": translate,
        "dept": department_translate,
        "display_dept": department_translate,
        "months": MONTHS,
        "years": list(
            range(
                2024,
                2032
            )
        ),
        "departments": DEPARTMENTS,
    }


# ============================================================
# HEALTH CHECK
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

        return jsonify({
            "status": "ok",
            "database": (
                "postgresql"
                if is_postgres()
                else "sqlite"
            ),
            "workers": workers_count,
            "attendance": attendance_count,
            "daily_attendance": daily_count,
        })

    except Exception as e:

        return jsonify({
            "status": "error",
            "error": repr(e),
        }), 500


# ============================================================
# DATABASE CHECK
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

    for table_name in tables:

        try:

            exists = table_exists(
                table_name
            )

            information[
                table_name
            ] = {
                "exists": exists,
                "columns": sorted(
                    columns(table_name)
                ) if exists else [],
            }

        except Exception as e:

            information[
                table_name
            ] = {
                "error": repr(e)
            }

    return render_template(
        "db_check.html",
        info=information
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
                    == hash_password(password)
                    or stored_password
                    == password
                )

        if valid:

            session["user_id"] = (
                user["id"]
            )

            session["language"] = (
                session.get(
                    "language",
                    "en"
                )
            )

            try:

                if "last_login" in columns(
                    "users"
                ):

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

            next_url = request.args.get(
                "next"
            )

            if next_url:
                return redirect(next_url)

            return redirect(
                url_for("dashboard")
            )

        flash(
            "Invalid User ID or Password.",
            "danger"
        )

    # Use the normal template when available.
    # If login.html is missing (for example when this single app.py
    # file is copied without the templates folder), show a built-in
    # login page instead of returning HTTP 500.
    try:
        return render_template(
            "login.html"
        )
    except Exception as template_error:
        app.logger.warning(
            "login.html could not be loaded: %r. Using built-in login page.",
            template_error,
        )

        next_url = request.args.get("next", "")
        message_html = ""

        try:
            messages = list(get_flashed_messages())
        except Exception:
            messages = []

        if messages:
            message_html = "<div class=\"alert\">" + "<br>".join(
                str(m) for m in messages
            ) + "</div>"

        return f"""
<!doctype html>
<html lang=\"en\">
<head>
<meta charset=\"utf-8\">
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
<title>Reedoy Payroll - Login</title>
<style>
body {{ margin:0; font-family:Arial,sans-serif; background:#f4f7fb; min-height:100vh; display:flex; align-items:center; justify-content:center; }}
.box {{ width:360px; max-width:90%; background:#fff; padding:30px; border-radius:12px; box-shadow:0 8px 30px rgba(0,0,0,.12); }}
h1 {{ margin:0 0 8px; color:#174a7e; font-size:24px; text-align:center; }}
p.sub {{ text-align:center; color:#666; margin:0 0 24px; }}
label {{ display:block; margin:12px 0 6px; font-weight:bold; color:#333; }}
input {{ width:100%; box-sizing:border-box; padding:11px; border:1px solid #ccc; border-radius:6px; font-size:15px; }}
button {{ width:100%; margin-top:20px; padding:12px; border:0; border-radius:6px; background:#1769aa; color:#fff; font-size:16px; cursor:pointer; }}
button:hover {{ background:#125486; }}
.alert {{ background:#fde8e8; color:#b42318; border:1px solid #f5b5b5; padding:10px; border-radius:6px; margin-bottom:15px; }}
</style>
</head>
<body>
<div class=\"box\">
<h1>REEDOY PAYROLL</h1>
<p class=\"sub\">Login to continue</p>
{message_html}
<form method=\"post\" action=\"/login\">
<input type=\"hidden\" name=\"next\" value=\"{next_url}\">
<label for=\"username\">User ID</label>
<input id=\"username\" name=\"username\" type=\"text\" autocomplete=\"username\" required>
<label for=\"password\">Password</label>
<input id=\"password\" name=\"password\" type=\"password\" autocomplete=\"current-password\" required>
<button type=\"submit\">Login</button>
</form>
</div>
</body>
</html>
"""


# ============================================================
# LOGOUT
# ============================================================

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


# ============================================================
# LANGUAGE
# ============================================================

@app.route("/language/<lang>")
@app.route("/set_language/<lang>")
def language(lang):

    if lang == "bn":

        session["language"] = "bn"

    else:

        session["language"] = "en"

    return redirect(
        request.referrer
        or url_for("dashboard")
    )


# ============================================================
# INDEX
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


# ============================================================
# BUILT-IN FALLBACK TEMPLATES
# ============================================================
# This standalone version can run even when the external templates/
# folder is not present beside the .py file. The normal templates are
# still used automatically if they exist.

BUILTIN_OT_REPORT_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>OT Report | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1500px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.filters{display:flex;gap:12px;align-items:end;flex-wrap:wrap}.field label{display:block;font-weight:600;font-size:13px;margin-bottom:5px}.field select{padding:9px;border:1px solid #cbd5e1;border-radius:7px}.btn{display:inline-block;background:#1764c0;color:#fff;border:0;border-radius:8px;padding:9px 14px;text-decoration:none;cursor:pointer}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.card{background:#f8fafc;padding:14px;border-radius:10px}.label{color:#64748b;font-size:13px}.value{font-size:21px;font-weight:800;color:#1764c0;margin-top:5px}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1000px}th,td{padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap}th{background:#f8fafc}th:nth-child(1),td:nth-child(1),th:nth-child(2),td:nth-child(2),th:nth-child(3),td:nth-child(3){text-align:left}.msg{padding:9px;background:#ecfdf5;color:#166534;border-radius:7px;margin-bottom:12px}.actions{display:flex;gap:10px;flex-wrap:wrap}@media(max-width:900px){.cards{grid-template-columns:repeat(2,1fr)}}@media(max-width:500px){.cards{grid-template-columns:1fr}}@media print{@page{size:A4 landscape;margin:6mm}.nav,.filters,.actions{display:none!important}.wrap{max-width:none;width:100%;margin:0;padding:0}.box{box-shadow:none;border:0;padding:2px;margin-bottom:6px}.cards{grid-template-columns:repeat(4,1fr)}table{min-width:0;font-size:7px}th,td{padding:3px 2px}tr{page-break-inside:avoid}}</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('worker_search') }}">Worker Search</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('reports') }}">Reports</a><a href="{{ url_for('accounts') }}">Accounts</a><a href="{{ url_for('ot_report') }}">OT Report</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('payment_history') }}">Payment History</a><a href="{{ url_for('payroll_closing') }}">Payroll Closing</a><a href="{{ url_for('settings') }}">Settings</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap"><div class="box filters"><div><h1 style="margin:0">Overtime Report</h1><p style="color:#64748b">Monthly OT hours and OT amount by worker and department.</p></div><form method="get" action="{{ url_for('ot_report') }}"><div class="field"><label>Month</label><select name="month">{% for m in months %}<option value="{{m}}" {% if month.split()[0]==m %}selected{% endif %}>{{m}}</option>{% endfor %}</select></div><div class="field"><label>Year</label><select name="year">{% for y in years %}<option value="{{y}}" {% if month.split()[1]|int==y %}selected{% endif %}>{{y}}</option>{% endfor %}</select></div><button class="btn" type="submit">Generate</button></form></div>
<div class="cards"><div class="card"><div class="label">Workers with OT</div><div class="value">{{ totals.workers }}</div></div><div class="card"><div class="label">Total OT Hours</div><div class="value">{{ '%.2f'|format(totals.hours) }}</div></div><div class="card"><div class="label">Total OT Amount</div><div class="value">à§³ {{ '%.2f'|format(totals.amount) }}</div></div><div class="card"><div class="label">Payroll Month</div><div class="value">{{ month }}</div></div></div>
<div class="box"><h2>Worker-wise OT â€” {{ month }}</h2><div class="tablewrap"><table><tr><th>Worker ID</th><th>Worker Name</th><th>Department</th><th>Designation</th><th>Basic Salary</th><th>OT Rate</th><th>OT Hours</th><th>OT Amount</th></tr>{% for r in rows %}<tr><td>{{r.id}}</td><td>{{r.bangla_name if language == 'bn' and r.bangla_name else r.name}}</td><td>{{display_dept(r.department)}}</td><td>{{r.designation or ''}}</td><td>à§³ {{'%.2f'|format((r.basic_salary or 0)|float)}}</td><td>à§³ {{'%.2f'|format((r.ot_rate or 0)|float)}}</td><td>{{'%.2f'|format((r.ot or 0)|float)}}</td><td>à§³ {{'%.2f'|format((r.ot_amt or 0)|float)}}</td></tr>{% else %}<tr><td colspan="9">No workers found.</td></tr>{% endfor %}<tr><th colspan="6">Total</th><th>{{'%.2f'|format(totals.hours)}}</th><th>à§³ {{'%.2f'|format(totals.amount)}}</th></tr></table></div></div>
<div class="box"><h2>Department-wise OT Summary</h2><div class="tablewrap"><table><tr><th>Department</th><th>Workers with OT</th><th>OT Hours</th><th>OT Amount</th></tr>{% for d in departments_summary %}<tr><td>{{display_dept(d.department)}}</td><td>{{d.workers}}</td><td>{{'%.2f'|format(d.hours)}}</td><td>à§³ {{'%.2f'|format(d.amount)}}</td></tr>{% endfor %}<tr><th>Total</th><th>{{totals.workers}}</th><th>{{'%.2f'|format(totals.hours)}}</th><th>à§³ {{'%.2f'|format(totals.amount)}}</th></tr></table></div></div>
<div class="box actions"><a class="btn" href="{{url_for('ot_report_export',month=month.split()[0],year=month.split()[1])}}">Export CSV</a><button class="btn" onclick="window.print()">Print OT Report</button><a class="btn" href="{{url_for('dashboard')}}">Dashboard</a></div></main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""

BUILTIN_DASHBOARD_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.nav a:hover{background:#dbeafe}
.wrap{max-width:1200px;margin:28px auto;padding:0 18px}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}.card,.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.label{color:#64748b;font-size:14px}.value{font-size:28px;font-weight:800;color:#1764c0;margin-top:8px}h1{margin-top:0}.box{margin-top:20px;overflow:auto}table{width:100%;border-collapse:collapse}th,td{padding:10px;border-bottom:1px solid #edf1f6;text-align:left}th{background:#f8fafc}
@media(max-width:800px){.cards{grid-template-columns:repeat(2,1fr)}}@media(max-width:500px){.cards{grid-template-columns:1fr}}
</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance & Calendar</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('settings') }}">Settings</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('backup_maintenance') }}">Backup</a><a href="{{ url_for('data_sync') }}">â†” Sync</a><a href="{{ url_for('logout') }}">Logout</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('reports') }}">Reports</a><a href="{{ url_for('accounts') }}">Accounts</a><a href="{{ url_for('ot_report') }}">OT Report</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('payment_history') }}">Payment History</a><a href="{{ url_for('payroll_closing') }}">Payroll Closing</a></nav></header>
<main class="wrap"><h1>{{ tr('Dashboard') }}</h1><p>{{ month }}</p>
<section class="cards"><div class="card"><div class="label">Total Workers</div><div class="value">{{ total_workers|default(0) }}</div></div><div class="card"><div class="label">Present Today</div><div class="value">{{ today_present|default(0) }}</div></div><div class="card"><div class="label">Absent Today</div><div class="value">{{ today_absent|default(0) }}</div></div><div class="card"><div class="label">Gross Salary</div><div class="value">à§³ {{ '%.2f'|format(gross|default(0)|float) }}</div></div></section>
<section class="box"><h2>Attendance</h2><p>Attendance & Calendar module is ready.</p><a href="{{ url_for('attendance') }}">Open Attendance & Calendar â†’</a></section><section class="box"><h2>Payslip</h2><p>View the selected worker monthly payslip.</p><a href="{{ url_for('payslip') }}">Open Payslip â†’</a></section><section class="box"><h2>Advance Salary</h2><p>Add, edit and delete worker advance salary records.</p><a href="{{ url_for('advance') }}">Open Advance Salary â†’</a></section><section class="box"><h2>Backup & Maintenance</h2><p>Create and download a safe backup of the payroll database.</p><a href="{{ url_for('backup_maintenance') }}">Open Backup & Maintenance â†’</a></section><section class="box"><h2>Reports</h2><p>View monthly payroll totals, department summary and worker payroll details.</p><a href="{{ url_for('reports') }}">Open Reports â†’</a></section><section class="box"><h2>Payment History</h2><p>View saved payment transactions by month and worker.</p><a href="{{ url_for('payment_history') }}">Open Payment History â†’</a></section><section class="box"><h2>Accounts / Financial Summary</h2><p>View monthly payroll financial summary and department-wise financial totals.</p><a href="{{ url_for('accounts') }}">Open Accounts â†’</a></section><section class="box"><h2>OT Report</h2><p>View monthly overtime hours and OT amount by worker and department.</p><a href="{{ url_for('ot_report') }}">Open OT Report â†’</a></section><section class="box"><h2>Payment Management</h2><p>Track paid, unpaid and partially paid monthly salaries.</p><a href="{{ url_for('payments') }}">Open Payments â†’</a></section><section class="box"><h2>Payroll Closing</h2><p>Lock a completed payroll month so attendance, advances and payments cannot be changed accidentally.</p><a href="{{ url_for('payroll_closing') }}">Open Payroll Closing â†’</a></section>
</main></body></html>
"""

BUILTIN_PAYMENTS_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Payment Management | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1500px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}.card{background:#f8fafc;padding:14px;border-radius:10px}.label{color:#64748b;font-size:13px}.value{font-size:20px;font-weight:800;color:#1764c0;margin-top:5px}.filters{display:flex;gap:10px;align-items:end;flex-wrap:wrap}.field label{display:block;font-weight:600;font-size:13px;margin-bottom:5px}.field select,.field input{padding:8px;border:1px solid #cbd5e1;border-radius:7px}.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:8px 12px;text-decoration:none;cursor:pointer}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1200px}th,td{padding:7px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap}th{background:#f8fafc}th:nth-child(2),td:nth-child(2),th:nth-child(3),td:nth-child(3){text-align:left}input.small,select.small{padding:5px;font-size:12px;width:105px}.save{padding:6px 9px;font-size:12px}.paid{font-weight:700}.partial{font-weight:700}.unpaid{font-weight:700}.danger{background:#dc2626}@media(max-width:1000px){.cards{grid-template-columns:repeat(3,1fr)}}@media(max-width:600px){.cards{grid-template-columns:1fr 1fr}}@media print{@page{size:A4 landscape;margin:6mm}.nav,.filters,.actions,.editcol{display:none!important}.wrap{max-width:none;width:100%;margin:0;padding:0}.box{box-shadow:none;border:0;padding:2px}.cards{grid-template-columns:repeat(6,1fr)}table{min-width:0;font-size:7px}th,td{padding:3px 2px}}
</style></head><body><header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('reports') }}">Reports</a><a href="{{ url_for('accounts') }}">Accounts</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('payment_history') }}">Payment History</a><a href="{{ url_for('settings') }}">Settings</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('backup_maintenance') }}">Backup</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap"><div class="box filters"><form method="get" action="{{ url_for('payments') }}"><div class="field"><label>Month</label><select name="month">{% for m in months %}<option value="{{m}}" {% if month.split()[0]==m %}selected{% endif %}>{{m}}</option>{% endfor %}</select></div><div class="field"><label>Year</label><select name="year">{% for y in years %}<option value="{{y}}" {% if month.split()[1]|int==y %}selected{% endif %}>{{y}}</option>{% endfor %}</select></div><button class="btn">Load</button></form></div>
<div class="cards"><div class="card"><div class="label">Workers</div><div class="value">{{totals.workers}}</div></div><div class="card"><div class="label">Net Payable</div><div class="value">à§³ {{'%.2f'|format(totals.net)}}</div></div><div class="card"><div class="label">Paid</div><div class="value">à§³ {{'%.2f'|format(totals.paid)}}</div></div><div class="card"><div class="label">Due</div><div class="value">à§³ {{'%.2f'|format(totals.due)}}</div></div><div class="card"><div class="label">Paid Workers</div><div class="value">{{totals.paid_count}}</div></div><div class="card"><div class="label">Unpaid/Partial</div><div class="value">{{totals.unpaid_count}}</div></div></div>
<div class="box"><h1>Payment Management â€” {{month}}</h1><div class="tablewrap"><table><tr><th>ID</th><th>Worker</th><th>Department</th><th>Net Payable</th><th>Paid Amount</th><th>Due</th><th>Status</th><th class="editcol">Payment Date</th><th class="editcol">Method</th><th class="editcol">Note</th><th class="editcol">Actions</th></tr>{% for r in rows %}<tr><td>{{r.id}}</td><td>{{r.bangla_name or r.name}}</td><td>{{r.department or ''}}</td><td>{{'%.2f'|format(r.net)}}</td><td>{{'%.2f'|format(r.paid_amount)}}</td><td>{{'%.2f'|format(r.due)}}</td><td class="{{r.status|lower}}">{{r.status}}</td><form class="payment-edit-form" method="post" action="{{url_for('save_payment')}}"><input type="hidden" name="worker_id" value="{{r.id}}"><input type="hidden" name="month" value="{{month.split()[0]}}"><input type="hidden" name="year" value="{{month.split()[1]}}"><td class="editcol"><input class="small payment-field" type="date" name="payment_date" value="{{r.payment_date or ''}}" disabled></td><td class="editcol"><select class="small payment-field" name="payment_method" disabled><option {% if r.payment_method=='Cash' %}selected{% endif %}>Cash</option><option {% if r.payment_method=='Bank' %}selected{% endif %}>Bank</option><option {% if r.payment_method=='Mobile Banking' %}selected{% endif %}>Mobile Banking</option><option {% if r.payment_method=='Cheque' %}selected{% endif %}>Cheque</option></select></td><td class="editcol"><input class="small payment-field" name="note" value="{{r.note or ''}}" disabled></td><td class="editcol"><input class="small payment-field" type="number" step="0.01" min="0" max="{{r.net}}" name="paid_amount" value="{{'%.2f'|format(r.paid_amount)}}" disabled><button class="btn edit-btn" type="button" onclick="togglePaymentEdit(this)">Edit</button><button class="btn save save-btn" type="submit" style="display:none">Save</button></form> <form method="post" action="{{url_for('delete_payment')}}" style="display:inline" onsubmit="return confirm('Delete this payment record for {{r.bangla_name or r.name}} â€” {{month}}?');"><input type="hidden" name="worker_id" value="{{r.id}}"><input type="hidden" name="month" value="{{month.split()[0]}}"><input type="hidden" name="year" value="{{month.split()[1]}}"><button class="btn danger delete-btn" type="submit">Delete</button></form></td></tr>{% endfor %}</table></div></div><div class="box actions"><a class="btn" href="{{url_for('payments_export',month=month.split()[0],year=month.split()[1])}}">Export CSV</a> <button class="btn" onclick="window.print()">Print Payment Report</button></div></main>
<script>
function togglePaymentEdit(btn){
  var row=btn.closest('tr');
  if(!row) return;
  var fields=row.querySelectorAll('.payment-field');
  var save=row.querySelector('.save-btn');
  var editing=btn.dataset.editing==='1';
  fields.forEach(function(f){ f.disabled=editing; });
  btn.dataset.editing=editing?'0':'1';
  btn.textContent=editing?'Edit':'Cancel';
  if(save) save.style.display=editing?'none':'inline-block';
}

document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""

BUILTIN_PAYMENT_HISTORY_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Payment History | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1450px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.filters{display:flex;gap:10px;align-items:end;flex-wrap:wrap}.field label{display:block;font-weight:600;font-size:13px;margin-bottom:5px}.field select,.field input{padding:9px;border:1px solid #cbd5e1;border-radius:7px}.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:9px 14px;text-decoration:none;cursor:pointer}.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.card{background:#f8fafc;padding:14px;border-radius:10px}.label{color:#64748b;font-size:13px}.value{font-size:21px;font-weight:800;color:#1764c0;margin-top:5px}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1000px}th,td{padding:8px;border-bottom:1px solid #e2e8f0;text-align:left;white-space:nowrap}th{background:#f8fafc}td.amount,th.amount{text-align:right}.actions{display:flex;gap:8px;flex-wrap:wrap}@media(max-width:800px){.cards{grid-template-columns:repeat(2,1fr)}}@media print{.nav,.filters,.actions{display:none!important}.box{box-shadow:none;border:0}}
</style></head><body><header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('reports') }}">Reports</a><a href="{{ url_for('accounts') }}">Accounts</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('payment_history') }}">Payment History</a><a href="{{ url_for('settings') }}">Settings</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('backup_maintenance') }}">Backup</a><a href="{{ url_for('logout') }}">Logout</a></nav></header><main class="wrap"><div class="box filters"><form method="get" action="{{ url_for('payment_history') }}"><div class="field"><label>Month</label><select name="month">{% for m in months %}<option value="{{m}}" {% if month.split()[0]==m %}selected{% endif %}>{{m}}</option>{% endfor %}</select></div><div class="field"><label>Year</label><select name="year">{% for y in years %}<option value="{{y}}" {% if month.split()[1]|int==y %}selected{% endif %}>{{y}}</option>{% endfor %}</select></div><div class="field"><label>Worker ID (optional)</label><input name="worker_id" value="{{worker_id or ''}}" placeholder="e.g. 224"></div><button class="btn">Load History</button></form></div><div class="cards"><div class="card"><div class="label">Transactions</div><div class="value">{{totals.count}}</div></div><div class="card"><div class="label">Total Paid</div><div class="value">à§³ {{'%.2f'|format(totals.amount)}}</div></div><div class="card"><div class="label">Workers Paid</div><div class="value">{{totals.workers}}</div></div><div class="card"><div class="label">Selected Month</div><div class="value">{{month}}</div></div></div><div class="box"><h1>Payment History â€” {{month}}</h1><p>Every saved payment transaction is recorded here. Editing a payment creates a new history entry.</p><div class="tablewrap"><table><tr><th>ID</th><th>Worker ID</th><th>Worker Name</th><th>Department</th><th class="amount">Amount</th><th>Payment Date</th><th>Method</th><th>Note</th><th>Created At</th></tr>{% for r in rows %}<tr><td>{{r.id}}</td><td>{{r.worker_id}}</td><td>{{r.bangla_name or r.name}}</td><td>{{r.department or ''}}</td><td class="amount">à§³ {{'%.2f'|format((r.amount or 0)|float)}}</td><td>{{r.payment_date or ''}}</td><td>{{r.payment_method or ''}}</td><td>{{r.note or ''}}</td><td>{{r.created_at or ''}}</td></tr>{% else %}<tr><td colspan="9">No payment transactions found.</td></tr>{% endfor %}</table></div></div><div class="box actions"><a class="btn" href="{{url_for('payment_history_export',month=month.split()[0],year=month.split()[1],worker_id=worker_id or '')}}">Export CSV</a><button class="btn" onclick="window.print()">Print History</button><a class="btn" href="{{url_for('payments',month=month.split()[0],year=month.split()[1])}}">Back to Payments</a></div></main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""

BUILTIN_PAYROLL_CLOSING_TEMPLATE = r"""
<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Payroll Closing | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1100px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:14px}label{font-weight:600;display:block;margin-bottom:6px}select,input,textarea{padding:10px;border:1px solid #cbd5e1;border-radius:7px;width:100%;box-sizing:border-box}button,.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 16px;cursor:pointer;text-decoration:none;display:inline-block}.danger{background:#dc2626}.success{background:#16a34a}.msg{padding:10px;border-radius:7px;margin-bottom:10px;background:#fef3c7}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%}th,td{padding:9px;border-bottom:1px solid #e2e8f0;text-align:left}th{background:#f8fafc}.locked{color:#b91c1c;font-weight:700}.open{color:#15803d;font-weight:700}@media(max-width:700px){.grid{grid-template-columns:1fr}}@media print{.nav,.actions,.filters{display:none!important}.box{box-shadow:none}}</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('payment_history') }}">Payment History</a><a href="{{ url_for('payroll_closing') }}">Payroll Closing</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap"><div class="box filters"><h1>Payroll Closing / Monthly Lock</h1>{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg">{{ message }}</div>{% endfor %}{% endwith %}<form method="get"><div class="grid"><div><label>Month</label><select name="month">{% for m in months %}<option value="{{m}}" {% if month.split()[0]==m %}selected{% endif %}>{{m}}</option>{% endfor %}</select></div><div><label>Year</label><select name="year">{% for y in years %}<option value="{{y}}" {% if month.split()[1]|int==y %}selected{% endif %}>{{y}}</option>{% endfor %}</select></div></div><br><button type="submit">Load Month</button></form></div>
<div class="box"><h2>{{ month }}</h2>{% if selected_lock %}<p class="locked">ðŸ”’ Payroll is LOCKED</p><p>Locked at: {{ selected_lock.locked_at or '' }}<br>Locked by: {{ selected_lock.locked_by or '' }}<br>Note: {{ selected_lock.note or '' }}</p><form method="post" action="{{ url_for('payroll_unlock') }}"><input type="hidden" name="month" value="{{month.split()[0]}}"><input type="hidden" name="year" value="{{month.split()[1]}}"><button class="btn danger" type="submit">Unlock Payroll</button></form>{% else %}<p class="open">ðŸŸ¢ Payroll is OPEN</p><form method="post" action="{{ url_for('payroll_close') }}"><input type="hidden" name="month" value="{{month.split()[0]}}"><input type="hidden" name="year" value="{{month.split()[1]}}"><label>Lock Note</label><textarea name="note" rows="3" placeholder="Optional note"></textarea><br><br><button class="btn success" type="submit">Lock This Month</button></form>{% endif %}</div>
<div class="box"><h2>Recent Payroll Months</h2><div class="tablewrap"><table><tr><th>Month</th><th>Status</th><th>Locked At</th><th>Locked By</th></tr>{% for r in recent %}<tr><td>{{r.month}}</td><td>{% if r.locked %}<span class="locked">LOCKED</span>{% else %}<span class="open">OPEN</span>{% endif %}</td><td>{{r.locked_at}}</td><td>{{r.locked_by}}</td></tr>{% endfor %}</table></div></div></main></body></html>
"""

BUILTIN_ACCOUNTS_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Accounts | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1400px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:18px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.grid{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}.card{background:#f8fafc;padding:14px;border-radius:10px}.label{color:#64748b;font-size:13px}.value{font-size:21px;font-weight:800;color:#1764c0;margin-top:5px}.filters{display:flex;gap:10px;align-items:end;flex-wrap:wrap}.field label{display:block;font-weight:600;font-size:13px;margin-bottom:5px}.field select{padding:9px;border:1px solid #cbd5e1;border-radius:7px}.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:9px 14px;text-decoration:none;cursor:pointer}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:800px}th,td{padding:8px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap}th{background:#f8fafc}th:first-child,td:first-child{text-align:left}@media(max-width:900px){.grid{grid-template-columns:repeat(2,1fr)}}@media(max-width:500px){.grid{grid-template-columns:1fr}}@media print{@page{size:A4 landscape;margin:6mm}.nav,.filters,.actions{display:none!important}.wrap{max-width:none;width:100%;margin:0;padding:0}.box{box-shadow:none;border:0}.grid{grid-template-columns:repeat(5,1fr)}table{min-width:0;font-size:7px}th,td{padding:3px}}</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('reports') }}">Reports</a><a href="{{ url_for('accounts') }}">Accounts</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('settings') }}">Settings</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('backup_maintenance') }}">Backup</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap"><div class="box filters"><div><h1 style="margin:0">Accounts / Financial Summary</h1><p style="color:#64748b">Monthly payroll financial summary based on existing payroll records.</p></div><form method="get" action="{{ url_for('accounts') }}"><div class="field"><label>Month</label><select name="month">{% for m in months %}<option value="{{ m }}" {% if month.split()[0]==m %}selected{% endif %}>{{ m }}</option>{% endfor %}</select><label style="margin-top:6px">Year</label><select name="year">{% for y in years %}<option value="{{ y }}" {% if month.split()[1]|int==y %}selected{% endif %}>{{ y }}</option>{% endfor %}</select></div><button class="btn" type="submit">Generate</button></form></div>
<div class="grid"><div class="card"><div class="label">Workers</div><div class="value">{{ totals.workers }}</div></div><div class="card"><div class="label">Basic Salary</div><div class="value">à§³ {{ '%.2f'|format(totals.basic) }}</div></div><div class="card"><div class="label">Gross Salary</div><div class="value">à§³ {{ '%.2f'|format(totals.gross) }}</div></div><div class="card"><div class="label">Advance</div><div class="value">à§³ {{ '%.2f'|format(totals.advance) }}</div></div><div class="card"><div class="label">Net Payable</div><div class="value">à§³ {{ '%.2f'|format(totals.net) }}</div></div></div>
<div class="box"><h2>Financial Breakdown â€” {{ month }}</h2><div class="tablewrap"><table><tr><th>Item</th><th>Amount (BDT)</th></tr><tr><td>Basic Salary</td><td>{{ '%.2f'|format(totals.basic) }}</td></tr><tr><td>Absent Deduction</td><td>- {{ '%.2f'|format(totals.absent_cut) }}</td></tr><tr><td>OT Amount</td><td>+ {{ '%.2f'|format(totals.ot_amt) }}</td></tr><tr><td>Nasta</td><td>+ {{ '%.2f'|format(totals.nasta) }}</td></tr><tr><td><strong>Gross Salary</strong></td><td><strong>{{ '%.2f'|format(totals.gross) }}</strong></td></tr><tr><td>Advance Deduction</td><td>- {{ '%.2f'|format(totals.advance) }}</td></tr><tr><td><strong>Net Payable</strong></td><td><strong>{{ '%.2f'|format(totals.net) }}</strong></td></tr></table></div></div>
<div class="box"><h2>Department Financial Summary</h2><div class="tablewrap"><table><tr><th>Department</th><th>Workers</th><th>Basic</th><th>Gross</th><th>Advance</th><th>Net Payable</th></tr>{% for d in departments_summary %}<tr><td>{{ d.department }}</td><td>{{ d.workers }}</td><td>{{ '%.2f'|format(d.basic) }}</td><td>{{ '%.2f'|format(d.gross) }}</td><td>{{ '%.2f'|format(d.advance) }}</td><td>{{ '%.2f'|format(d.net) }}</td></tr>{% endfor %}<tr><th>Total</th><th>{{ totals.workers }}</th><th>{{ '%.2f'|format(totals.basic) }}</th><th>{{ '%.2f'|format(totals.gross) }}</th><th>{{ '%.2f'|format(totals.advance) }}</th><th>{{ '%.2f'|format(totals.net) }}</th></tr></table></div></div>
<div class="box actions"><a class="btn" href="{{ url_for('accounts_export', month=month.split()[0], year=month.split()[1]) }}">Export CSV</a> <button class="btn" onclick="window.print()">Print Accounts</button></div></main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""

BUILTIN_ATTENDANCE_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Attendance & Calendar | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1400px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}label{font-weight:600;font-size:14px;display:block;margin-bottom:6px}select,input{padding:9px;border:1px solid #cbd5e1;border-radius:7px;font-size:14px}button,.btn{background:#1764c0;color:white;border:0;border-radius:8px;padding:10px 16px;cursor:pointer;text-decoration:none;display:inline-block}.grid{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:12px}.summary{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}.stat{background:#f8fafc;padding:14px;border-radius:9px}.stat b{display:block;font-size:22px;margin-top:4px}.calendar{overflow:auto}table{border-collapse:collapse;width:100%;min-width:850px}th,td{border:1px solid #e2e8f0;padding:8px;text-align:center}th{background:#f8fafc}.day{font-weight:700}.weekend{background:#fff7ed}.msg{padding:10px;border-radius:7px;margin-bottom:10px}.danger{background:#fee2e2;color:#991b1b}.success{background:#dcfce7;color:#166534}.workername{font-size:18px;font-weight:700}.muted{color:#64748b;font-size:13px}@media(max-width:800px){.grid,.summary{grid-template-columns:1fr 1fr}}
</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('logout') }}">Logout</a><a href="{{ url_for('department') }}">Department Salary</a></nav></header>
<main class="wrap">
<div class="box"><h1>Attendance & Calendar</h1>
{% with messages = get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}
<form method="get" action="{{ url_for('attendance') }}"><div class="grid"><div><label>Worker</label><select name="worker_id"><option value="">-- Select Worker --</option>{% for w in workers %}<option value="{{ w.id }}" {% if worker and w.id|string == worker.id|string %}selected{% endif %}>{{ w.id }} - {{ w.bangla_name if language=='bn' and w.bangla_name else w.name }}{% if w.department %} ({{ w.department }}){% endif %}</option>{% endfor %}</select></div><div><label>Month</label><select name="month">{% for m in months %}<option value="{{ m }}" {% if month.split()[0]==m %}selected{% endif %}>{{ m }}</option>{% endfor %}</select></div><div><label>Year</label><select name="year">{% for y in years %}<option value="{{ y }}" {% if month.split()[1]|int==y %}selected{% endif %}>{{ y }}</option>{% endfor %}</select></div><div style="align-self:end"><button type="submit">Load Attendance</button></div></div></form></div>
{% if worker %}<div class="box"><div class="workername">{{ worker.bangla_name if language=='bn' and worker.bangla_name else worker.name }}</div><div class="muted">ID: {{ worker.id }} &nbsp; | &nbsp; Department: {{ worker.department or '' }} &nbsp; | &nbsp; Designation: {{ worker.designation or '' }} &nbsp; | &nbsp; Month: {{ month }}</div></div>
<form method="post" action="{{ url_for('save_attendance') }}"><input type="hidden" name="worker_id" value="{{ worker.id }}"><input type="hidden" name="month" value="{{ month.split()[0] }}"><input type="hidden" name="year" value="{{ month.split()[1] }}">
<div class="box"><h2>Monthly Summary</h2><div class="summary"><div class="stat"><label>Present Days</label><input type="number" min="0" name="present_days" value="{{ summary.get('present',0) }}"></div><div class="stat"><label>Absent Days</label><input type="number" min="0" name="absent_days" value="{{ summary.get('absent',0) }}"></div><div class="stat"><label>Overtime Hours</label><input type="number" min="0" step="0.01" name="ot_hours" value="{{ summary.get('ot',0) }}"></div><div class="stat"><label>Net Salary</label><b>à§³ {{ '%.2f'|format(summary.get('net',0)|float) }}</b></div></div></div>
<div class="box calendar"><h2>Daily Attendance â€” {{ month }}</h2><div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px"><button type="button" class="btn" onclick="markAllAttendance('P')">âœ“ Mark All Present</button><button type="button" class="btn" style="background:#b91c1c" onclick="markAllAttendance('A')">âœ• Mark All Absent</button><button type="button" class="btn" style="background:#64748b" onclick="markAllAttendance('')">â†º Clear All / Not Set</button></div><table><thead><tr><th>Day</th><th>Date</th><th>Status</th></tr></thead><tbody>{% for d in days %}<tr class="{% if d.date.weekday()==4 %}weekend{% endif %}"><td class="day">{{ d.day }}</td><td>{{ d.date.strftime('%d-%b-%Y') }}</td><td><select name="day_{{ d.day }}"><option value="" {% if not d.status %}selected{% endif %}>-- Not Set --</option><option value="P" {% if d.status=='P' %}selected{% endif %}>P - Present</option><option value="A" {% if d.status=='A' %}selected{% endif %}>A - Absent</option></select></td></tr>{% endfor %}</tbody></table></div>
<div class="box"><button type="submit">Save Attendance</button> <a class="btn" style="background:#64748b" href="{{ url_for('attendance') }}">Clear</a></div></form>{% endif %}
</main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
function markAllAttendance(status){
  var selects=document.querySelectorAll('select[name^="day_"]');
  selects.forEach(function(sel){
    var row=sel.closest('tr');
    var dayCell=row ? row.querySelector('.day') : null;
    var dayNumber=dayCell ? parseInt(dayCell.textContent.trim(),10) : 0;
    // Friday is the weekly holiday. It must always display as Absent.
    if(row && row.classList.contains('weekend')){
      sel.value='A';
    } else {
      sel.value=status;
    }
  });
}
</script>
</body></html>
"""


BUILTIN_ADVANCE_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Advance Salary | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1250px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.grid{display:grid;grid-template-columns:repeat(5,minmax(130px,1fr));gap:12px}.field label{font-weight:600;font-size:13px;display:block;margin-bottom:6px}.field input,.field select{width:100%;box-sizing:border-box;padding:9px;border:1px solid #cbd5e1;border-radius:7px}.btn,button{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 15px;text-decoration:none;cursor:pointer;display:inline-block}.btn.edit{background:#0f766e}.btn.delete{background:#b91c1c}.msg{padding:10px;border-radius:7px;margin-bottom:10px}.danger{background:#fee2e2;color:#991b1b}.success{background:#dcfce7;color:#166534}.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:800px}th,td{padding:10px;border-bottom:1px solid #e2e8f0;text-align:left}th{background:#f8fafc}.amount{font-weight:700;text-align:right}.actions{white-space:nowrap}@media(max-width:900px){.grid{grid-template-columns:1fr 1fr}}
</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('logout') }}">Logout</a><a href="{{ url_for('department') }}">Department Salary</a></nav></header>
<main class="wrap"><div class="box"><h1>Advance Salary</h1>
{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}
<form method="get"><div class="grid"><div class="field"><label>Worker</label><select name="worker_id"><option value="">All Workers</option>{% for w in workers %}<option value="{{ w.id }}" {% if request.args.get('worker_id','')|string == w.id|string %}selected{% endif %}>{{ w.id }} - {{ w.bangla_name if language=='bn' and w.bangla_name else w.name }}</option>{% endfor %}</select></div><div class="field"><label>Month</label><select name="month">{% for m in months %}<option value="{{ m }}" {% if month.split()[0]==m %}selected{% endif %}>{{ m }}</option>{% endfor %}</select></div><div class="field"><label>Year</label><select name="year">{% for y in years %}<option value="{{ y }}" {% if month.split()[1]|int==y %}selected{% endif %}>{{ y }}</option>{% endfor %}</select></div><div class="field" style="align-self:end"><button type="submit">View Advances</button></div></div></form></div>
<div class="box"><h2>Add Advance Salary</h2><form method="post" action="{{ url_for('save_advance') }}"><div class="grid"><div class="field"><label>Worker</label><select name="worker_id" required><option value="">Select Worker</option>{% for w in workers %}<option value="{{ w.id }}">{{ w.id }} - {{ w.bangla_name if language=='bn' and w.bangla_name else w.name }}</option>{% endfor %}</select></div><div class="field"><label>Date</label><input type="date" name="advance_date" value="{{ today_date }}" required></div><div class="field"><label>Amount (BDT)</label><input type="number" name="amount" min="0.01" step="0.01" required></div><div class="field"><label>Note</label><input type="text" name="note"></div><div class="field" style="align-self:end"><input type="hidden" name="month_year" value="{{ month }}"><button type="submit">Save Advance</button></div></div></form></div>
<div class="box"><div style="display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap"><h2 style="margin:0">Advance Records â€” {{ month }}</h2><div style="font-size:18px;font-weight:700;color:#1764c0">Total Advance: BDT {{ '%.2f'|format(total_advance|float) }}</div></div></div><div class="box" style="display:flex;gap:8px;flex-wrap:wrap"><button type="button" onclick="window.print()">Print</button><a class="btn" href="{{ url_for('advance_export', month=month.split(' ')[0], year=month.split(' ')[1], worker_id=request.args.get('worker_id','')) }}">Export Excel</a><a class="btn edit" href="{{ url_for('advance_pdf', month=month.split(' ')[0], year=month.split(' ')[1], worker_id=request.args.get('worker_id','')) }}">Export PDF</a></div><div class="box tablewrap"><table><thead><tr><th>ID</th><th>Worker</th><th>Department</th><th>Date</th><th>Amount</th><th>Note</th><th>Actions</th></tr></thead><tbody>{% if rows %}{% for r in rows %}<tr><td>{{ r.id }}</td><td>{{ r.worker_id }} - {{ r.bangla_name if language=='bn' and r.bangla_name else r.worker_name }}</td><td>{{ display_dept(r.department) }}</td><td>{{ r.advance_date or '' }}</td><td class="amount">BDT {{ '%.2f'|format(r.amount|float) }}</td><td>{{ r.note or '' }}</td><td class="actions"><a class="btn edit" href="{{ url_for('edit_advance', aid=r.id) }}">Edit</a> <a class="btn delete" href="{{ url_for('delete_advance', aid=r.id) }}" onclick="return confirm('Delete this advance record?')">Delete</a></td></tr>{% endfor %}{% else %}<tr><td colspan="7">No advance records found.</td></tr>{% endif %}</tbody></table></div></main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""

BUILTIN_ADVANCE_EDIT_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Edit Advance | REEDOY PAYROLL</title><style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.wrap{max-width:700px;margin:30px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:24px;box-shadow:0 4px 14px rgba(0,0,0,.04)}label{display:block;font-weight:600;margin:12px 0 6px}input,select{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:7px}button,.btn{margin-top:16px;background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 15px;text-decoration:none;cursor:pointer;display:inline-block}.cancel{background:#64748b;margin-left:8px}.nav{margin-bottom:18px}.nav a{color:#1764c0;text-decoration:none;margin-right:12px}</style></head><body><main class="wrap"><div class="box"><div class="nav"><a href="{{ url_for('advance') }}">â† Advance Salary</a><a href="{{ url_for('dashboard') }}">Dashboard</a></div><h1>Edit Advance Salary</h1><form method="post"><label>Worker</label><select name="worker_id" required>{% for w in workers %}<option value="{{ w.id }}" {% if w.id|string == record.worker_id|string %}selected{% endif %}>{{ w.id }} - {{ w.bangla_name if language=='bn' and w.bangla_name else w.name }}</option>{% endfor %}</select><label>Date</label><input type="date" name="advance_date" value="{{ record.advance_date or '' }}" required><label>Amount (BDT)</label><input type="number" name="amount" min="0.01" step="0.01" value="{{ record.amount or 0 }}" required><label>Note</label><input type="text" name="note" value="{{ record.note or '' }}"><input type="hidden" name="month_year" value="{{ record.month_year or '' }}"><button type="submit">Update Advance</button><a class="btn cancel" href="{{ url_for('advance', month=record.month_year) }}">Cancel</a></form></div></main></body></html>
"""

BUILTIN_LOGIN_TEMPLATE = r"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>REEDOY PAYROLL Login</title><style>body{font-family:Arial,sans-serif;background:#f3f6fb;display:flex;align-items:center;justify-content:center;min-height:100vh}.box{background:#fff;padding:30px;border-radius:14px;box-shadow:0 8px 30px rgba(0,0,0,.08);width:340px}input{width:100%;padding:11px;margin:7px 0 14px;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:7px}button{width:100%;padding:11px;background:#1764c0;color:#fff;border:0;border-radius:7px;cursor:pointer}.msg{padding:8px;background:#fee2e2;color:#991b1b;margin-bottom:12px;border-radius:6px}</style></head><body><div class="box"><h2>REEDOY PAYROLL</h2>{% with messages=get_flashed_messages(with_categories=true) %}{% for c,m in messages %}<div class="msg">{{ m }}</div>{% endfor %}{% endwith %}<form method="post"><label>User ID</label><input name="username" required><label>Password</label><input type="password" name="password" required><button type="submit">Login</button></form></div></body></html>"""

BUILTIN_PAYSLIP_TEMPLATE = r"""<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Reedoy Payroll - Payslip</title><style>body{font-family:Arial,sans-serif;background:#f4f7fb;margin:0;color:#1f2937}.wrap{max-width:900px;margin:30px auto;padding:0 16px}.box{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:24px;margin-bottom:18px;box-shadow:0 4px 16px rgba(0,0,0,.05)}h1{margin-top:0;color:#174a7e}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}.item{padding:12px;background:#f8fafc;border-radius:8px}.label{font-size:13px;color:#64748b}.value{font-size:18px;font-weight:700;margin-top:4px}.net{font-size:26px;color:#1764c0}.nav a{margin-right:14px;color:#1764c0;text-decoration:none}select,button{padding:10px;border:1px solid #cbd5e1;border-radius:7px}button{background:#1764c0;color:#fff;border:0;cursor:pointer}.actions{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0 18px}.btn{display:inline-block;padding:10px 14px;border-radius:7px;background:#1764c0;color:#fff;text-decoration:none;border:0;cursor:pointer;font-size:14px}.btn.print{background:#0f766e}@media print{body{background:#fff}.wrap{max-width:none;margin:0;padding:0}.box{box-shadow:none;border:0;border-radius:0}.nav,form,.actions{display:none}.item{background:#fff;border:1px solid #ddd}h1{font-size:22px}}@media(max-width:600px){.grid{grid-template-columns:1fr}}</style></head><body><main class="wrap"><div class="box"><div class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a></div><h1>Single Payslip</h1><form method="get"><div class="grid"><div><div class="label">Worker</div><select name="worker_id" required><option value="">Select Worker</option>{% for w in workers %}<option value="{{ w.id }}" {% if worker and w.id|string == worker.id|string %}selected{% endif %}>{{ w.id }} - {{ w.bangla_name if language == 'bn' and w.bangla_name else w.name }}</option>{% endfor %}</select></div><div><div class="label">Payroll Month</div><select name="month">{% for m in months %}<option value="{{ m }}" {% if month.startswith(m + ' ') %}selected{% endif %}>{{ m }}</option>{% endfor %}</select><select name="year">{% for y in years %}<option value="{{ y }}" {% if month.endswith(y|string) %}selected{% endif %}>{{ y }}</option>{% endfor %}</select></div></div><br><button type="submit">View Payslip</button></form></div>{% if worker %}<div class="box"><h2>REEDOY TEXTILE DYEING PRINTING & FINISHING</h2><p><b>Payroll Month:</b> {{ month }}</p><div class="actions"><a class="btn" href="{{ url_for('export_payslip_pdf', worker_id=worker.id, month=month.split()[0], year=month.split()[1]) }}">Export PDF</a><button class="btn print" type="button" onclick="window.print()">Print Payslip</button></div><div class="grid"><div class="item"><div class="label">Worker ID</div><div class="value">{{ worker.id }}</div></div><div class="item"><div class="label">English Name</div><div class="value">{{ worker.name }}</div></div><div class="item"><div class="label">Bangla Name</div><div class="value">{{ worker.bangla_name or 'â€”' }}</div></div><div class="item"><div class="label">Department</div><div class="value">{{ display_dept(worker.department) }}</div></div><div class="item"><div class="label">Basic Salary</div><div class="value">BDT {{ '%.2f'|format(worker.basic_salary|float) }}</div></div><div class="item"><div class="label">Salary Basis</div><div class="value">30 Days</div></div><div class="item"><div class="label">Present Days</div><div class="value">{{ summary.present|default(0) }}</div></div><div class="item"><div class="label">Absent Days</div><div class="value">{{ summary.absent|default(0) }}</div></div><div class="item"><div class="label">Absent Deduction</div><div class="value">BDT {{ '%.2f'|format(summary.absent_cut|default(0)|float) }}</div></div><div class="item"><div class="label">Overtime</div><div class="value">{{ summary.ot|default(0) }} hours / BDT {{ '%.2f'|format(summary.ot_amt|default(0)|float) }}</div></div><div class="item"><div class="label">Refreshment</div><div class="value">BDT {{ '%.2f'|format(summary.nasta|default(0)|float) }}</div></div><div class="item"><div class="label">Advance Salary</div><div class="value">BDT {{ '%.2f'|format(summary.advance|default(0)|float) }}</div></div><div class="item"><div class="label">Gross Salary</div><div class="value">BDT {{ '%.2f'|format(summary.gross|default(0)|float) }}</div></div><div class="item"><div class="label">Net Payable</div><div class="value net">BDT {{ '%.2f'|format(summary.net|default(0)|float) }}</div></div></div></div>{% endif %}</main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>"""


BUILTIN_WORKER_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Workers Management | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1400px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.toolbar{display:flex;gap:10px;flex-wrap:wrap}.toolbar input{flex:1;min-width:220px}input,select{padding:10px;border:1px solid #cbd5e1;border-radius:7px;width:100%;box-sizing:border-box}button,.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 16px;cursor:pointer;text-decoration:none;display:inline-block}.btn.edit{background:#0f766e}.btn.delete{background:#dc2626}.msg{padding:10px;border-radius:7px;margin-bottom:10px}.danger{background:#fee2e2;color:#991b1b}.success{background:#dcfce7;color:#166534}table{border-collapse:collapse;width:100%;min-width:950px}th,td{border:1px solid #e2e8f0;padding:9px;text-align:left}th{background:#f8fafc}.tablewrap{overflow:auto}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.field label{display:block;font-weight:600;margin-bottom:6px}.actions{display:flex;gap:8px;flex-wrap:wrap}@media(max-width:700px){.grid{grid-template-columns:1fr}}@media print{.nav,.toolbar,.actions{display:none}.box{box-shadow:none}}
</style></head><body><header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('logout') }}">Logout</a><a href="{{ url_for('department') }}">Department Salary</a></nav></header><main class="wrap"><div class="box"><h1>Workers Management</h1>{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}<div class="toolbar"><form method="get" style="display:flex;gap:10px;flex:1;flex-wrap:wrap"><input name="q" value="{{ q or '' }}" placeholder="Search ID, English name, Bangla name, department or designation"><select name="status" style="max-width:180px"><option value="">All Status</option><option value="Active" {{ 'selected' if status=='Active' else '' }}>Active</option><option value="Inactive" {{ 'selected' if status=='Inactive' else '' }}>Inactive</option></select><button type="submit">Search / Filter</button></form><a class="btn" href="{{ url_for('add_worker') }}">+ Add Worker</a></div></div><div class="box tablewrap"><table><thead><tr><th>ID</th><th>English Name</th><th>Bangla Name</th><th>Department</th><th>Designation</th><th>Basic Salary</th><th>OT Rate</th><th>Status</th><th>Actions</th></tr></thead><tbody>{% for w in workers %}<tr><td>{{ w.id }}</td><td>{{ w.name or '' }}</td><td>{{ w.bangla_name or '' }}</td><td>{{ w.department or '' }}</td><td>{{ w.designation or '' }}</td><td>{{ '%.2f'|format((w.basic_salary or 0)|float) }}</td><td>{{ '%.2f'|format((w.ot_rate or 0)|float) }}</td><td><span class="btn {{ 'edit' if (w.status or 'Active')=='Active' else 'delete' }}" style="cursor:default">{{ w.status or 'Active' }}</span></td><td class="actions"><a class="btn edit" href="{{ url_for('edit_worker',worker_id=w.id) }}">Edit</a><form method="post" action="{{ url_for('toggle_worker_status',worker_id=w.id) }}" style="display:inline"><input type="hidden" name="q" value="{{ q or '' }}"><input type="hidden" name="status" value="{{ status or '' }}"><button class="btn {{ 'delete' if (w.status or 'Active')=='Active' else 'edit' }}" type="submit" onclick="return confirm('Change status for {{ w.name }}?');">{{ 'Set Inactive' if (w.status or 'Active')=='Active' else 'Set Active' }}</button></form><a class="btn delete" href="{{ url_for('delete_worker',worker_id=w.id) }}">Delete</a></td></tr>{% else %}<tr><td colspan="9">No workers found.</td></tr>{% endfor %}</tbody></table></div></main><script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""

BUILTIN_WORKER_FORM_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{ form_title }} | REEDOY PAYROLL</title><style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none}.wrap{max-width:900px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:24px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.field label{display:block;font-weight:600;margin-bottom:6px}.field{margin-bottom:4px}input,select{padding:10px;border:1px solid #cbd5e1;border-radius:7px;width:100%;box-sizing:border-box}.actions{margin-top:20px;display:flex;gap:10px}.btn,button{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 16px;text-decoration:none;cursor:pointer}.cancel{background:#64748b}@media(max-width:700px){.grid{grid-template-columns:1fr}}</style></head><body><header class="top"><div class="brand">REEDOY PAYROLL</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('department') }}">Department Salary</a></nav></header><main class="wrap"><div class="box"><h1>{{ form_title }}</h1>{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div>{{ message }}</div>{% endfor %}{% endwith %}<form method="post"><div class="grid"><div class="field"><label>English Name *</label><input name="name" required value="{{ worker.name if worker and worker.name is defined else (worker.get('name','') if worker else '') }}"></div><div class="field"><label>Bangla Name</label><input name="bangla_name" value="{{ worker.bangla_name if worker and worker.bangla_name is defined else (worker.get('bangla_name','') if worker else '') }}"></div><div class="field"><label>Department</label><select name="department">{% for d in departments if d != 'All Departments' %}<option value="{{ d }}" {% if worker and ((worker.department is defined and worker.department==d) or (worker.get('department','')==d if worker.get is defined else false)) %}selected{% endif %}>{{ d }}</option>{% endfor %}</select></div><div class="field"><label>Designation</label><input name="designation" value="{{ worker.designation if worker and worker.designation is defined else (worker.get('designation','') if worker else '') }}"></div><div class="field"><label>Basic Salary (BDT)</label><input type="number" step="0.01" min="0" name="basic_salary" value="{{ worker.basic_salary if worker and worker.basic_salary is defined else (worker.get('basic_salary',0) if worker else 0) }}" id="basic_salary"></div><div class="field"><label>OT Rate</label><input type="number" step="0.01" min="0" name="ot_rate" value="{{ worker.ot_rate if worker and worker.ot_rate is defined else (worker.get('ot_rate',0) if worker else 0) }}" id="ot_rate"></div><div class="field"><label>Refreshment Bill</label><input type="number" step="0.01" min="0" name="refreshment_bill" value="{{ worker.refreshment_bill if worker and worker.refreshment_bill is defined else (worker.get('refreshment_bill',0) if worker else 0) }}"></div><div class="field"><label>Status</label><select name="status"><option value="Active" {% if not worker or (worker.status is defined and worker.status in ('', None)) or (worker.get('status','Active') if worker.get is defined else worker.status) == 'Active' %}selected{% endif %}>Active</option><option value="Inactive" {% if worker and ((worker.status is defined and worker.status=='Inactive') or (worker.get('status','')=='Inactive' if worker.get is defined else false)) %}selected{% endif %}>Inactive</option></select></div></div><div class="actions"><button type="submit">{{ submit_text }}</button><a class="btn cancel" href="{{ url_for('workers') }}">Cancel</a></div>
<script>
(function () {
  const basic = document.getElementById("basic_salary");
  const ot = document.getElementById("ot_rate");
  if (!basic || !ot) return;

  function updateOtRate() {
    const b = parseFloat(basic.value);
    if (!Number.isFinite(b) || b < 0) {
      ot.value = "";
      return;
    }
    ot.value = Math.round(b / 30 / 12);
  }

  basic.addEventListener("input", updateOtRate);
  basic.addEventListener("change", updateOtRate);
})();
</script>

</form></div></main></body></html>
"""


BUILTIN_DEPARTMENT_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Department Salary Sheet | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}
.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}
.wrap{max-width:1500px;margin:24px auto;padding:0 18px}
.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
.filters{display:grid;grid-template-columns:repeat(4,minmax(150px,1fr));gap:12px;align-items:end}
label{display:block;font-weight:600;font-size:13px;margin-bottom:6px}
select{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:7px;background:#fff}
button,.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 15px;text-decoration:none;cursor:pointer;display:inline-block}
.btn.green{background:#0f766e}.btn.gray{background:#64748b}
.cards{display:grid;grid-template-columns:repeat(5,1fr);gap:12px}
.card{background:#f8fafc;border-radius:10px;padding:15px}.label{color:#64748b;font-size:13px}.value{font-size:21px;font-weight:800;color:#1764c0;margin-top:5px}
.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1450px}th,td{border:1px solid #e2e8f0;padding:8px;text-align:right;white-space:nowrap}th{background:#f8fafc;text-align:center}th.name,td.name,th.dept,td.dept,th.desig,td.desig{text-align:left}.total{font-weight:800;background:#eef6ff}.actions{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
@media(max-width:900px){.filters{grid-template-columns:1fr 1fr}.cards{grid-template-columns:repeat(2,1fr)}}
@media print{
  @page{size:A4 landscape;margin:6mm}
  html,body{background:#fff;width:100%;margin:0;padding:0}
  .nav,.filters,.actions{display:none!important}
  .top{display:none!important}
  .wrap{max-width:none;width:100%;margin:0;padding:0}
  .box{box-shadow:none;border:0;padding:2px;margin:0 0 4px}
  .box h1,.box h2{margin:2px 0 4px}
  .cards{grid-template-columns:repeat(5,1fr);gap:4px}
  .card{padding:4px}
  .label{font-size:8px}
  .value{font-size:11px;margin-top:1px}
  .tablewrap{overflow:visible;width:100%}
  table{width:100%;min-width:0;table-layout:fixed;font-size:6.5px}
  th,td{padding:3px 2px;white-space:nowrap;overflow:hidden}
  th.name,td.name{width:15%}
  th.dept,td.dept{width:9%}
  th.desig,td.desig{width:9%}
  th:first-child,td:first-child{width:4%}
  th:nth-child(5),td:nth-child(5){width:7%}
  th:nth-child(6),td:nth-child(6){width:5%}
  th:nth-child(7),td:nth-child(7){width:5%}
  th:nth-child(8),td:nth-child(8){width:6%}
  th:nth-child(n+9),td:nth-child(n+9){width:6.5%}
  tr{page-break-inside:avoid}
}
</style>
</head>
<body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap">
<div class="box">
<h1>Department Salary Sheet</h1>
<form method="get">
<div class="filters">
<div><label>Payroll Month</label><select name="month">{% for m in months %}<option value="{{ m }}" {% if month.split()[0]==m %}selected{% endif %}>{{ m }}</option>{% endfor %}</select></div>
<div><label>Year</label><select name="year">{% for y in years %}<option value="{{ y }}" {% if month.split()[1]|int==y %}selected{% endif %}>{{ y }}</option>{% endfor %}</select></div>
<div><label>Department</label><select name="department"><option value="">All Departments</option>{% for d in departments %}<option value="{{ d }}" {% if department==d %}selected{% endif %}>{{ d }}</option>{% endfor %}</select></div>
<div><button type="submit">Generate Sheet</button></div>
</div>
</form>
</div>
<div class="box"><h2>{{ month }}{% if department %} â€” {{ department }}{% else %} â€” All Departments{% endif %}</h2>
<div style="margin:8px 0 14px;padding:10px 12px;background:#eef6ff;border:1px solid #d7e7fb;border-radius:8px;font-weight:700;color:#334155">Salary Basis: 30 Days &nbsp; | &nbsp; Weekly Holiday: Friday &nbsp; | &nbsp; Friday is excluded from absence deduction</div>
<div class="cards">
<div class="card"><div class="label">Workers</div><div class="value">{{ totals.workers }}</div></div>
<div class="card"><div class="label">Basic Salary</div><div class="value">à§³ {{ '%.2f'|format(totals.basic|float) }}</div></div>
<div class="card"><div class="label">Gross Salary</div><div class="value">à§³ {{ '%.2f'|format(totals.gross|float) }}</div></div>
<div class="card"><div class="label">Advance</div><div class="value">à§³ {{ '%.2f'|format(totals.advance|float) }}</div></div>
<div class="card"><div class="label">Net Payable</div><div class="value">à§³ {{ '%.2f'|format(totals.net|float) }}</div></div>
</div></div>
<div class="box tablewrap">
<table>
<thead><tr><th>ID</th><th class="name">Worker Name</th><th class="dept">Department</th><th class="desig">Designation</th><th>Basic</th><th>Present</th><th>Absent</th><th>OT Hours</th><th>Absent Ded.</th><th>OT Amount</th><th>Nasta</th><th>Gross</th><th>Advance</th><th>Net Payable</th></tr></thead>
<tbody>
{% for r in rows %}<tr><td>{{ r.id }}</td><td class="name">{{ r.bangla_name if language=='bn' and r.bangla_name else r.name }}</td><td class="dept">{{ display_dept(r.department) }}</td><td class="desig">{{ r.designation or '' }}</td><td>{{ '%.2f'|format(r.basic_salary|default(0)|float) }}</td><td>{{ r.present|default(0) }}</td><td>{{ r.absent|default(0) }}</td><td>{{ r.ot|default(0) }}</td><td>{{ '%.2f'|format(r.absent_cut|default(0)|float) }}</td><td>{{ '%.2f'|format(r.ot_amt|default(0)|float) }}</td><td>{{ '%.2f'|format(r.nasta|default(0)|float) }}</td><td>{{ '%.2f'|format(r.gross|default(0)|float) }}</td><td>{{ '%.2f'|format(r.advance|default(0)|float) }}</td><td>{{ '%.2f'|format(r.net|default(0)|float) }}</td></tr>{% else %}<tr><td colspan="14">No workers found.</td></tr>{% endfor %}
{% if rows %}<tr class="total"><td colspan="4" class="name">TOTAL</td><td>{{ '%.2f'|format(totals.basic|float) }}</td><td>{{ totals.present }}</td><td>{{ totals.absent }}</td><td>{{ totals.ot }}</td><td>{{ '%.2f'|format(totals.absent_cut|float) }}</td><td>{{ '%.2f'|format(totals.ot_amt|float) }}</td><td>{{ '%.2f'|format(totals.nasta|float) }}</td><td>{{ '%.2f'|format(totals.gross|float) }}</td><td>{{ '%.2f'|format(totals.advance|float) }}</td><td>{{ '%.2f'|format(totals.net|float) }}</td></tr>{% endif %}
</tbody></table></div>
<div class="box actions"><a class="btn green" href="{{ url_for('export_department', month=month.split()[0], year=month.split()[1], department=department) }}">Export Excel</a><button class="btn" type="button" onclick="window.print()">Print Salary Sheet</button><a class="btn gray" href="{{ url_for('dashboard') }}">Dashboard</a></div>
</main>
<script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""


BUILTIN_SETTINGS_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Company Settings | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}
.wrap{max-width:1000px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:22px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.field label{display:block;font-weight:700;margin-bottom:6px}.field input,.field textarea{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px}
.field.full{grid-column:1/-1}.field textarea{min-height:90px;resize:vertical}.btn,button{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 16px;text-decoration:none;cursor:pointer;display:inline-block}.btn.secondary{background:#475569}
.msg{padding:10px 12px;border-radius:8px;margin-bottom:12px}.success{background:#dcfce7;color:#166534}.danger{background:#fee2e2;color:#991b1b}
.preview{display:flex;align-items:center;gap:15px;padding:14px;background:#f8fafc;border-radius:10px}.preview img{width:64px;height:64px;object-fit:contain;border-radius:8px;background:#fff;border:1px solid #e2e8f0}
@media(max-width:700px){.grid{grid-template-columns:1fr}.field.full{grid-column:auto}}
</style></head><body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap"><div class="box"><h1>Company Settings</h1>
{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}
<form method="post"><div class="grid">
<div class="field"><label>Company Name</label><input name="company_name" value="{{ settings.get('company_name','') }}" required></div>
<div class="field"><label>Phone</label><input name="company_phone" value="{{ settings.get('company_phone','') }}"></div>
<div class="field full"><label>Company Address</label><textarea name="company_address">{{ settings.get('company_address','') }}</textarea></div>
<div class="field"><label>Email</label><input type="email" name="company_email" value="{{ settings.get('company_email','') }}"></div>
<div class="field"><label>Logo URL</label><input name="company_logo" value="{{ settings.get('company_logo','') }}" placeholder="https://..."></div>
</div><br><button type="submit">Save Settings</button> <a class="btn secondary" href="{{ url_for('dashboard') }}">Back</a></form>
</div>
<div class="box"><h2>Company Preview</h2><div class="preview">{% if settings.get('company_logo') %}<img src="{{ settings.get('company_logo') }}" alt="Logo">{% endif %}<div><strong>{{ settings.get('company_name','REEDOY PAYROLL') }}</strong><br>{{ settings.get('company_address','') }}<br>{{ settings.get('company_phone','') }}{% if settings.get('company_email') %} | {{ settings.get('company_email') }}{% endif %}</div></div></div>
</main></body></html>
"""

BUILTIN_USERS_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>User Management | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}.wrap{max-width:1200px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}.btn{background:#1764c0;color:#fff;border-radius:8px;padding:9px 14px;text-decoration:none;display:inline-block}.btn.delete{background:#b91c1c}table{border-collapse:collapse;width:100%}th,td{padding:10px;border-bottom:1px solid #e2e8f0;text-align:left}th{background:#f8fafc}.badge{padding:4px 8px;border-radius:12px;background:#e2e8f0;font-size:12px}.active{background:#dcfce7;color:#166534}.inactive{background:#fee2e2;color:#991b1b}.msg{padding:10px;border-radius:8px;margin-bottom:12px}.success{background:#dcfce7;color:#166534}.danger{background:#fee2e2;color:#991b1b}
</style></head><body><header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('settings') }}">Company Settings</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap"><div class="box"><h1>User Management</h1>{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}<a class="btn" href="{{ url_for('add_user') }}">+ Add User</a></div>
<div class="box"><table><thead><tr><th>ID</th><th>Username</th><th>Full Name</th><th>Role</th><th>Status</th><th>Created</th><th>Last Login</th><th>Action</th></tr></thead><tbody>
{% for u in users %}<tr><td>{{ u.id }}</td><td>{{ u.username }}</td><td>{{ u.full_name or '' }}</td><td>{{ u.role or '' }}</td><td><span class="badge {{ 'active' if u.active else 'inactive' }}">{{ 'Active' if u.active else 'Inactive' }}</span></td><td>{{ u.created_at or '' }}</td><td>{{ u.last_login or '' }}</td><td>{% if current_user and current_user.get('id') != u.id %}<a class="btn" href="{{ url_for('edit_user',user_id=u.id) }}">Edit</a> <form method="post" action="{{ url_for('toggle_user',user_id=u.id) }}" style="display:inline"><button class="btn" type="submit">{{ 'Deactivate' if u.active else 'Activate' }}</button></form> <a class="btn delete" href="{{ url_for('delete_user',user_id=u.id) }}" onclick="return confirm('Delete this user? This cannot be undone.')">Delete</a>{% else %}Current User{% endif %}</td></tr>{% else %}<tr><td colspan="8">No users found.</td></tr>{% endfor %}
</tbody></table></div></main></body></html>
"""

BUILTIN_USER_FORM_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{{ "Edit User" if edit_mode else "Add User" }} | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%}.brand{font-weight:800;font-size:20px;color:#1764c0}.wrap{max-width:700px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:22px}.field{margin-bottom:14px}.field label{display:block;font-weight:700;margin-bottom:6px}.field input,.field select{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:8px}.btn,button{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 16px;text-decoration:none;cursor:pointer}.btn.secondary{background:#475569}.msg{padding:10px;border-radius:8px;margin-bottom:12px}.danger{background:#fee2e2;color:#991b1b}</style></head>
<body><header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div></header><main class="wrap"><div class="box"><h1>Add User</h1>
{% with messages=get_flashed_messages(with_categories=true) %}{% for category,message in messages %}<div class="msg {{ category }}">{{ message }}</div>{% endfor %}{% endwith %}
<form method="post"><div class="field"><label>Username</label><input name="username" value="{{ user.get('username','') if user else '' }}" required></div><div class="field"><label>Full Name</label><input name="full_name" value="{{ user.get('full_name','') if user else '' }}"></div><div class="field"><label>Password{% if edit_mode %} (leave blank to keep current password){% endif %}</label><input type="password" name="password" {% if not edit_mode %}required{% endif %}></div><div class="field"><label>Role</label><select name="role">{% for role in ['Administrator','Manager','Operator','Viewer'] %}<option value="{{ role }}" {% if user and user.get('role')==role %}selected{% endif %}>{{ role }}</option>{% endfor %}</select></div><button type="submit">{{ "Update User" if edit_mode else "Create User" }}</button> <a class="btn secondary" href="{{ url_for('users') }}">Cancel</a></form>
</div></main></body></html>
"""

BUILTIN_ACTIVITY_TEMPLATE = r"""
<!doctype html><html lang="{{ 'bn' if language == 'bn' else 'en' }}"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Activity Log | REEDOY PAYROLL</title>
<style>body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none}.wrap{max-width:1400px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;overflow:auto}.filters{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:12px}.field label{display:block;font-weight:700;margin-bottom:5px}.field input{width:100%;box-sizing:border-box;padding:9px;border:1px solid #cbd5e1;border-radius:8px}.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 14px;text-decoration:none;display:inline-block;cursor:pointer}.btn.secondary{background:#475569}.btn.green{background:#15803d}table{border-collapse:collapse;width:100%;min-width:800px}th,td{padding:9px;border-bottom:1px solid #e2e8f0;text-align:left}th{background:#f8fafc}.muted{color:#64748b;font-size:13px}@media(max-width:800px){.filters{grid-template-columns:1fr 1fr}}@media(max-width:500px){.filters{grid-template-columns:1fr}}</style></head>
<body><header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div><nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('settings') }}">Company Settings</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap">
<div class="box"><h1>Activity Log</h1><form method="get"><div class="filters"><div class="field"><label>Username</label><input name="username" value="{{ username or '' }}" placeholder="All users"></div><div class="field"><label>Action contains</label><input name="action" value="{{ action or '' }}" placeholder="e.g. Payment"></div><div class="field"><label>From Date</label><input type="date" name="from_date" value="{{ from_date or '' }}"></div><div class="field"><label>To Date</label><input type="date" name="to_date" value="{{ to_date or '' }}"></div></div><div style="margin-top:12px"><button class="btn" type="submit">Filter</button> <a class="btn secondary" href="{{ url_for('activity') }}">Reset</a> <a class="btn green" href="{{ url_for('activity_export', username=username, action=action, from_date=from_date, to_date=to_date) }}">Export CSV</a></div></form><p class="muted">Showing up to 500 matching records.</p></div>
<div class="box"><table><thead><tr><th>ID</th><th>Username</th><th>Action</th><th>Time</th></tr></thead><tbody>{% for row in logs %}<tr><td>{{ row.get('id','') }}</td><td>{{ row.get('username','') }}</td><td>{{ row.get('action','') }}</td><td>{{ row.get('log_time',row.get('created_at','')) }}</td></tr>{% else %}<tr><td colspan="4">No activity found.</td></tr>{% endfor %}</tbody></table></div></main></body></html>
"""



BUILTIN_DATA_SYNC_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Online â†” Offline Data Sync | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}
.wrap{max-width:1100px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
h1,h2{margin-top:0}.btn{display:inline-block;background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 15px;text-decoration:none;cursor:pointer;font-size:14px}
.btn.danger{background:#b42318}.btn.secondary{background:#475569}.file{padding:12px;border:1px solid #cbd5e1;border-radius:8px;width:100%;box-sizing:border-box;margin:8px 0 12px}
.note{background:#eff6ff;border-left:4px solid #1764c0;padding:12px;border-radius:6px}.warn{background:#fff7ed;border-left:4px solid #ea580c;padding:12px;border-radius:6px}
.msg{padding:11px;border-radius:8px;margin-bottom:12px}.success{background:#dcfce7;color:#166534}.danger-msg{background:#fee2e2;color:#991b1b}
table{width:100%;border-collapse:collapse}th,td{padding:9px;border-bottom:1px solid #e2e8f0;text-align:left}th{background:#f8fafc}
</style>
</head>
<body>
<header class="top">
<div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav">
<a href="{{ url_for('dashboard') }}">Dashboard</a>
<a href="{{ url_for('workers') }}">Workers</a>
<a href="{{ url_for('attendance') }}">Attendance</a>
<a href="{{ url_for('payslip') }}">Payslip</a>
<a href="{{ url_for('advance') }}">Advance Salary</a>
<a href="{{ url_for('department') }}">Department Salary</a>
<a href="{{ url_for('reports') }}">Reports</a>
<a href="{{ url_for('backup_maintenance') }}">Backup</a>
<a href="{{ url_for('data_sync') }}">â†” Sync</a>
<a href="{{ url_for('logout') }}">Logout</a>
</nav>
</header>
<main class="wrap">
{% with messages=get_flashed_messages(with_categories=true) %}
{% for category,message in messages %}
<div class="msg {{ 'success' if category=='success' else 'danger-msg' }}">{{message}}</div>
{% endfor %}
{% endwith %}

<div class="box">
<h1>Online â†” Offline Data Sync</h1>
<p>à¦à¦‡ tab à¦¦à¦¿à§Ÿà§‡ Online à¦à¦¬à¦‚ Offline software-à¦à¦° à¦¸à¦®à§à¦ªà§‚à¦°à§à¦£ payroll data à¦à¦• backup file-à¦à¦° à¦®à¦¾à¦§à§à¦¯à¦®à§‡ à¦à¦•à§‡ à¦…à¦ªà¦°à§‡à¦° à¦®à¦§à§à¦¯à§‡ à¦¨à§‡à¦“à§Ÿà¦¾ à¦¯à¦¾à¦¬à§‡à¥¤</p>
<div class="note">
<strong>Backup file:</strong> Reedoy portable backup (.rdb)<br>
Worker ID, Worker Name, Bangla Name, Attendance, Daily Attendance, Advance, Salary/Payment, Payroll Lock, Settings, Users à¦“ Activity data backup-à¦à¦° à¦®à¦§à§à¦¯à§‡ à¦°à¦¾à¦–à¦¾ à¦¹à¦¬à§‡à¥¤
</div>
</div>

<div class="box">
<h2>1. Current Database â†’ Backup File</h2>
<p>à¦à¦‡ software-à¦à¦° à¦¬à¦°à§à¦¤à¦®à¦¾à¦¨ data à¦à¦•à¦Ÿà¦¿ portable backup file-à¦ à¦¨à¦¿à¦¨à¥¤ à¦¤à¦¾à¦°à¦ªà¦° à¦¸à§‡à¦‡ file à¦…à¦¨à§à¦¯ Online/Offline software-à¦ Restore à¦•à¦°à¦¾ à¦¯à¦¾à¦¬à§‡à¥¤</p>
<p><strong>Current database:</strong> {{ 'ONLINE / PostgreSQL' if is_online else 'OFFLINE / SQLite' }}</p>
<form method="post" action="{{ url_for('create_portable_backup') }}">
<button class="btn" type="submit">Create Portable Backup & Download</button>
</form>
</div>

<div class="box">
<h2>2. Backup File â†’ Current Database</h2>
<div class="warn">
<strong>à¦¸à¦¤à¦°à§à¦•à¦¤à¦¾:</strong> Restore à¦•à¦°à¦²à§‡ à¦à¦‡ software-à¦à¦° à¦¬à¦°à§à¦¤à¦®à¦¾à¦¨ data backup file-à¦à¦° data à¦¦à¦¿à§Ÿà§‡ replace à¦¹à¦¬à§‡à¥¤ Restore à¦•à¦°à¦¾à¦° à¦†à¦—à§‡ à¦¬à¦°à§à¦¤à¦®à¦¾à¦¨ data-à¦à¦° à¦à¦•à¦Ÿà¦¿ safety backup à¦¸à§à¦¬à§Ÿà¦‚à¦•à§à¦°à¦¿à§Ÿà¦­à¦¾à¦¬à§‡ à¦¤à§ˆà¦°à¦¿ à¦¹à¦¬à§‡à¥¤ Worker ID à¦“ à¦ªà§à¦°à§‹à¦¨à§‹ record ID à¦…à¦ªà¦°à¦¿à¦¬à¦°à§à¦¤à¦¿à¦¤ à¦°à¦¾à¦–à¦¾à¦° à¦šà§‡à¦·à§à¦Ÿà¦¾ à¦•à¦°à¦¾ à¦¹à¦¬à§‡à¥¤
</div>
<form method="post" action="{{ url_for('restore_portable_backup') }}" enctype="multipart/form-data" onsubmit="return confirm('Restore à¦•à¦°à¦²à§‡ à¦¬à¦°à§à¦¤à¦®à¦¾à¦¨ database-à¦à¦° data backup file-à¦à¦° data à¦¦à¦¿à§Ÿà§‡ replace à¦¹à¦¬à§‡à¥¤ Continue?');">
<input class="file" type="file" name="backup_file" accept=".rdb,.zip" required>
<button class="btn danger" type="submit">Restore Backup to This Software</button>
</form>
</div>

<div class="box">
<h2>à¦¬à§à¦¯à¦¬à¦¹à¦¾à¦°à§‡à¦° à¦¨à¦¿à§Ÿà¦®</h2>
<table>
<tr><th>à¦•à¦¾à¦œ</th><th>à¦¯à§‡à¦­à¦¾à¦¬à§‡ à¦•à¦°à¦¬à§‡à¦¨</th></tr>
<tr><td>Online â†’ Offline</td><td>Online-à¦ Create Portable Backup â†’ file download â†’ Offline-à¦ à¦à¦‡ tab â†’ Restore</td></tr>
<tr><td>Offline â†’ Online</td><td>Offline-à¦ Create Portable Backup â†’ file à¦¨à¦¿à§Ÿà§‡ Online-à¦ à¦à¦‡ tab â†’ Restore</td></tr>
<tr><td>Data à¦¨à¦¿à¦°à¦¾à¦ªà¦¤à§à¦¤à¦¾</td><td>Restore-à¦à¦° à¦†à¦—à§‡ à¦¬à¦°à§à¦¤à¦®à¦¾à¦¨ database-à¦à¦° safety backup à¦¤à§ˆà¦°à¦¿ à¦¹à¦¬à§‡</td></tr>
<tr><td>Database</td><td>Online PostgreSQL à¦à¦¬à¦‚ Offline SQLiteâ€”à¦¦à§à¦‡ à¦§à¦°à¦¨à§‡à¦° database support à¦•à¦°à¦¬à§‡</td></tr>
</table>
</div>

<div class="box">
<a class="btn secondary" href="{{ url_for('backup_maintenance') }}">â† Backup & Maintenance</a>
<a class="btn secondary" href="{{ url_for('dashboard') }}">Dashboard</a>
</div>
</main>
</body>
</html>
"""

BUILTIN_BACKUP_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Backup & Maintenance | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}
.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}
.wrap{max-width:1000px;margin:24px auto;padding:0 18px}
.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:22px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:11px 16px;text-decoration:none;cursor:pointer;display:inline-block}
.btn.secondary{background:#475569}
.msg{padding:11px 13px;border-radius:8px;margin-bottom:14px}
.success{background:#dcfce7;color:#166534}
.danger{background:#fee2e2;color:#991b1b}
table{width:100%;border-collapse:collapse}
th,td{padding:9px;border-bottom:1px solid #e2e8f0;text-align:left}
th{background:#f8fafc}
.muted{color:#64748b}
</style>
</head>
<body>
<header class="top">
<div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav">
<a href="{{ url_for('dashboard') }}">Dashboard</a>
<a href="{{ url_for('workers') }}">Workers</a>
<a href="{{ url_for('attendance') }}">Attendance</a>
<a href="{{ url_for('payslip') }}">Payslip</a>
<a href="{{ url_for('advance') }}">Advance Salary</a>
<a href="{{ url_for('department') }}">Department Salary</a>
<a href="{{ url_for('settings') }}">Settings</a>
<a href="{{ url_for('users') }}">Users</a>
<a href="{{ url_for('activity') }}">Activity Log</a>
<a href="{{ url_for('logout') }}">Logout</a>
</nav>
</header>
<main class="wrap">
<div class="box">
<h1>Backup & Maintenance</h1>
<p>Create a complete SQLite database backup without changing or deleting any payroll data.</p>
<p class="muted">Database: {{ db_path }}</p>
<form method="post" action="{{ url_for('create_database_backup') }}">
<button class="btn" type="submit">Create Database Backup</button>
</form>
</div>

<div class="box">
<h2>Existing Backups</h2>
<p class="muted">Before restoring, the current database is automatically backed up for safety.</p>
<table>
<thead><tr><th>Backup File</th><th>Size</th><th>Created</th><th>Action</th></tr></thead>
<tbody>
{% for row in backups %}
<tr>
<td>{{ row.name }}</td>
<td>{{ row.size }}</td>
<td>{{ row.created }}</td>
<td>
<a class="btn secondary" href="{{ url_for('download_database_backup', filename=row.name) }}">Download</a>
<form method="post" action="{{ url_for('restore_database_backup') }}" style="display:inline" onsubmit="return confirm('Restore this backup? The current database will first be backed up automatically.');">
<input type="hidden" name="filename" value="{{ row.name }}">
<button class="btn" type="submit">Restore</button>
</form>
</td>
</tr>
{% else %}
<tr><td colspan="4">No backup has been created yet.</td></tr>
{% endfor %}
</tbody>
</table>
</div>

<div class="box">
<a class="btn secondary" href="{{ url_for('dashboard') }}">â† Back to Dashboard</a>
</div>
</main>
</body>
</html>
"""


BUILTIN_REPORTS_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Reports | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}
.wrap{max-width:1500px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
.grid{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:12px}.field label{display:block;font-weight:700;margin-bottom:6px}.field select{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:8px}
button,.btn{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 15px;text-decoration:none;cursor:pointer;display:inline-block}.btn.secondary{background:#475569}.btn.green{background:#15803d}
.cards{display:grid;grid-template-columns:repeat(6,1fr);gap:12px}.card{background:#f8fafc;padding:14px;border-radius:10px}.label{color:#64748b;font-size:13px}.value{font-size:21px;font-weight:800;margin-top:5px;color:#1764c0}
.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1000px}th,td{padding:8px 7px;border-bottom:1px solid #e2e8f0;text-align:right;white-space:nowrap}th{background:#f8fafc}th:first-child,td:first-child{text-align:left}.muted{color:#64748b;font-size:13px}
@media(max-width:1000px){.cards{grid-template-columns:repeat(3,1fr)}.grid{grid-template-columns:1fr 1fr}}
@media(max-width:600px){.cards,.grid{grid-template-columns:1fr}}
@media print{
 @page{size:A4 landscape;margin:6mm}html,body{background:#fff;width:100%;margin:0;padding:0}.top,.filters,.actions{display:none!important}.wrap{max-width:none;width:100%;margin:0;padding:0}.box{box-shadow:none;border:0;padding:2px;margin-bottom:5px}.cards{grid-template-columns:repeat(6,1fr);gap:4px}.card{padding:5px}.label{font-size:7px}.value{font-size:10px}.tablewrap{overflow:visible}table{width:100%;min-width:0;table-layout:fixed;font-size:6.5px}th,td{padding:3px 2px;overflow:hidden}.name{width:17%;text-align:left!important}.dept{width:10%;text-align:left!important}tr{page-break-inside:avoid}
}
</style></head>
<body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance Salary</a><a href="{{ url_for('department') }}">Department Salary</a><a href="{{ url_for('settings') }}">Settings</a><a href="{{ url_for('users') }}">Users</a><a href="{{ url_for('activity') }}">Activity Log</a><a href="{{ url_for('backup_maintenance') }}">Backup</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap">
<div class="box filters"><h1>Payroll Reports</h1><p class="muted">Monthly payroll overview and department-wise summary.</p>
<form method="get" action="{{ url_for('reports') }}"><div class="grid"><div class="field"><label>Month</label><select name="month">{% for m in months %}<option value="{{ m }}" {% if month.split()[0]==m %}selected{% endif %}>{{ m }}</option>{% endfor %}</select></div><div class="field"><label>Year</label><select name="year">{% for y in years %}<option value="{{ y }}" {% if month.split()[1]|int==y %}selected{% endif %}>{{ y }}</option>{% endfor %}</select></div><div class="field" style="align-self:end"><button type="submit">Generate Report</button></div><div class="field actions" style="align-self:end"><a class="btn green" href="{{ url_for('export_reports',month=month.split()[0],year=month.split()[1]) }}">Export Excel</a> <button class="btn secondary" type="button" onclick="window.print()">Print Report</button></div></div></form></div>
<div class="box"><h2>{{ month }} â€” Payroll Summary</h2><div class="cards">
<div class="card"><div class="label">Workers</div><div class="value">{{ totals.workers }}</div></div>
<div class="card"><div class="label">Basic Salary</div><div class="value">à§³ {{ '%.2f'|format(totals.basic|float) }}</div></div>
<div class="card"><div class="label">Absent Deduction</div><div class="value">à§³ {{ '%.2f'|format(totals.absent_cut|float) }}</div></div>
<div class="card"><div class="label">OT Amount</div><div class="value">à§³ {{ '%.2f'|format(totals.ot_amt|float) }}</div></div>
<div class="card"><div class="label">Gross Salary</div><div class="value">à§³ {{ '%.2f'|format(totals.gross|float) }}</div></div>
<div class="card"><div class="label">Net Payable</div><div class="value">à§³ {{ '%.2f'|format(totals.net|float) }}</div></div>
</div></div>
<div class="box tablewrap"><h2>Department Summary</h2><table><thead><tr><th>Department</th><th>Workers</th><th>Basic Salary</th><th>Present</th><th>Absent</th><th>OT Hours</th><th>Absent Ded.</th><th>OT Amount</th><th>Nasta</th><th>Gross</th><th>Advance</th><th>Net Payable</th></tr></thead><tbody>{% for row in departments_summary %}<tr><td class="dept">{{ row.department }}</td><td>{{ row.workers }}</td><td>{{ '%.2f'|format(row.basic|float) }}</td><td>{{ '%.2f'|format(row.present|float) }}</td><td>{{ '%.2f'|format(row.absent|float) }}</td><td>{{ '%.2f'|format(row.ot|float) }}</td><td>{{ '%.2f'|format(row.absent_cut|float) }}</td><td>{{ '%.2f'|format(row.ot_amt|float) }}</td><td>{{ '%.2f'|format(row.nasta|float) }}</td><td>{{ '%.2f'|format(row.gross|float) }}</td><td>{{ '%.2f'|format(row.advance|float) }}</td><td>{{ '%.2f'|format(row.net|float) }}</td></tr>{% else %}<tr><td colspan="12">No department data found.</td></tr>{% endfor %}<tr><th>Total</th><th>{{ totals.workers }}</th><th>{{ '%.2f'|format(totals.basic|float) }}</th><th>{{ '%.2f'|format(totals.present|float) }}</th><th>{{ '%.2f'|format(totals.absent|float) }}</th><th>{{ '%.2f'|format(totals.ot|float) }}</th><th>{{ '%.2f'|format(totals.absent_cut|float) }}</th><th>{{ '%.2f'|format(totals.ot_amt|float) }}</th><th>{{ '%.2f'|format(totals.nasta|float) }}</th><th>{{ '%.2f'|format(totals.gross|float) }}</th><th>{{ '%.2f'|format(totals.advance|float) }}</th><th>{{ '%.2f'|format(totals.net|float) }}</th></tr></tbody></table></div>
<div class="box tablewrap"><h2>Worker Payroll Detail</h2><table><thead><tr><th>ID</th><th class="name">Worker Name</th><th class="dept">Department</th><th>Basic</th><th>Present</th><th>Absent</th><th>OT Hours</th><th>Absent Ded.</th><th>OT Amount</th><th>Nasta</th><th>Gross</th><th>Advance</th><th>Net Payable</th></tr></thead><tbody>{% for row in rows %}<tr><td>{{ row.id }}</td><td class="name">{{ row.bangla_name if language=='bn' and row.bangla_name else row.name }}</td><td class="dept">{{ row.department or '' }}</td><td>{{ '%.2f'|format(row.basic_salary|float) }}</td><td>{{ '%.2f'|format(row.present|float) }}</td><td>{{ '%.2f'|format(row.absent|float) }}</td><td>{{ '%.2f'|format(row.ot|float) }}</td><td>{{ '%.2f'|format(row.absent_cut|float) }}</td><td>{{ '%.2f'|format(row.ot_amt|float) }}</td><td>{{ '%.2f'|format(row.nasta|float) }}</td><td>{{ '%.2f'|format(row.gross|float) }}</td><td>{{ '%.2f'|format(row.advance|float) }}</td><td>{{ '%.2f'|format(row.net|float) }}</td></tr>{% endfor %}</tbody></table></div>
<div class="box actions"><a class="btn secondary" href="{{ url_for('dashboard') }}">â† Back to Dashboard</a></div>
</main><script>
document.addEventListener('DOMContentLoaded', function(){
  function addSelectSearch(sel){
    if(!sel || sel.dataset.workerSearchReady) return;
    sel.dataset.workerSearchReady='1';
    var box=document.createElement('input');
    box.type='search'; box.className='worker-search-box';
    box.placeholder='ðŸ” Search Worker: ID / Name / Bangla Name';
    box.style.cssText='width:100%;box-sizing:border-box;padding:10px 12px;margin:0 0 7px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff';
    sel.parentNode.insertBefore(box,sel);
    var original=Array.from(sel.options).map(function(o){return {text:o.text,value:o.value,html:o.outerHTML};});
    box.addEventListener('input',function(){
      var q=box.value.trim().toLowerCase(), current=sel.value;
      sel.innerHTML='';
      original.forEach(function(o){if(!q || o.text.toLowerCase().indexOf(q)!==-1) sel.insertAdjacentHTML('beforeend',o.html);});
      if(Array.from(sel.options).some(function(o){return o.value===current;})) sel.value=current;
    });
  }
  document.querySelectorAll('select[name="worker_id"]').forEach(addSelectSearch);
  var path=window.location.pathname;
  var tablePages=['/payments','/payment-history','/ot-report','/department','/reports','/accounts'];
  if(tablePages.indexOf(path)!==-1 && !document.querySelector('.worker-search-table-box')){
    var target=null;
    document.querySelectorAll('table').forEach(function(t){if(!target && /worker|à¦•à¦°à§à¦®à§€/i.test(t.innerText.slice(0,500))) target=t;});
    if(target){
      var wrap=document.createElement('div'); wrap.className='worker-search-table-box';
      wrap.style.cssText='margin:0 0 10px;padding:10px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px';
      wrap.innerHTML='<label style="display:block;font-weight:600;margin-bottom:6px">ðŸ” Search Worker</label><input type="search" placeholder="Worker ID / English Name / Bangla Name" style="width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px;background:#fff">';
      target.parentNode.insertBefore(wrap,target);
      var inp=wrap.querySelector('input');
      inp.addEventListener('input',function(){var q=inp.value.trim().toLowerCase();Array.from(target.querySelectorAll('tr')).forEach(function(row,i){if(i>0) row.style.display=(!q||row.innerText.toLowerCase().indexOf(q)!==-1)?'':'none';});});
    }
  }
});
</script>
</body></html>
"""



BUILTIN_WORKER_SEARCH_TEMPLATE = r"""
<!doctype html>
<html lang="{{ 'bn' if language == 'bn' else 'en' }}">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Worker Search | REEDOY PAYROLL</title>
<style>
body{margin:0;background:#f3f6fb;color:#1e293b;font-family:Arial,"Noto Sans Bengali",sans-serif}
.top{background:#fff;border-bottom:1px solid #e5eaf2;padding:15px 4%;display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap}
.brand{font-weight:800;font-size:20px;color:#1764c0}.nav a{display:inline-block;margin:3px;padding:9px 12px;border-radius:8px;background:#eef2f7;color:#334155;text-decoration:none;font-size:14px}
.wrap{max-width:1450px;margin:24px auto;padding:0 18px}.box{background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:20px;margin-bottom:18px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
.grid{display:grid;grid-template-columns:2fr 1fr 1fr auto;gap:10px;align-items:end}.field label{display:block;font-weight:700;margin-bottom:6px}.field input,.field select{width:100%;box-sizing:border-box;padding:10px;border:1px solid #cbd5e1;border-radius:8px}
.btn,button{background:#1764c0;color:#fff;border:0;border-radius:8px;padding:10px 14px;text-decoration:none;display:inline-block;cursor:pointer}.btn.secondary{background:#475569}.btn.small{padding:7px 10px;font-size:13px}.muted{color:#64748b;font-size:13px}
.tablewrap{overflow:auto}table{border-collapse:collapse;width:100%;min-width:1150px}th,td{padding:9px 8px;border-bottom:1px solid #e2e8f0;text-align:left;white-space:nowrap}th{background:#f8fafc}.actions a{margin-right:5px}
.badge{display:inline-block;padding:4px 8px;border-radius:12px;background:#e2e8f0;font-size:12px}.active{background:#dcfce7;color:#166534}.inactive{background:#fee2e2;color:#991b1b}
@media(max-width:850px){.grid{grid-template-columns:1fr 1fr}.grid .wide{grid-column:1/-1}}
</style></head>
<body>
<header class="top"><div class="brand">{{ settings.get('company_name','REEDOY PAYROLL') if settings else 'REEDOY PAYROLL' }}</div>
<nav class="nav"><a href="{{ url_for('dashboard') }}">Dashboard</a><a href="{{ url_for('workers') }}">Workers</a><a href="{{ url_for('attendance') }}">Attendance</a><a href="{{ url_for('payslip') }}">Payslip</a><a href="{{ url_for('advance') }}">Advance</a><a href="{{ url_for('payments') }}">Payments</a><a href="{{ url_for('logout') }}">Logout</a></nav></header>
<main class="wrap">
<div class="box">
<h1>Worker Search</h1>
<p class="muted">Search by Worker ID, English name, Bangla name, phone, department or designation.</p>
<form method="get" action="{{ url_for('worker_search') }}">
<div class="grid">
<div class="field wide"><label>Search</label><input name="q" value="{{ q or '' }}" placeholder="Worker ID / Name / Bangla Name / Phone / Designation"></div>
<div class="field"><label>Department</label><select name="department"><option value="">All Departments</option>{% for d in departments %}<option value="{{ d }}" {% if department==d %}selected{% endif %}>{{ d }}</option>{% endfor %}</select></div>
<div class="field"><label>Status</label><select name="status"><option value="">All Status</option><option value="Active" {% if status=='Active' %}selected{% endif %}>Active</option><option value="Inactive" {% if status=='Inactive' %}selected{% endif %}>Inactive</option></select></div>
<div><button type="submit">Search</button> <a class="btn secondary" href="{{ url_for('worker_search') }}">Reset</a></div>
</div></form>
</div>
<div class="box">
<strong>{{ results|length }}</strong> worker(s) found.
<div class="tablewrap" style="margin-top:12px"><table><thead><tr><th>ID</th><th>English Name</th><th>Bangla Name</th><th>Department</th><th>Designation</th><th>Phone</th><th>Basic Salary</th><th>Status</th><th>Quick Actions</th></tr></thead>
<tbody>{% for w in results %}<tr><td>{{ w.id }}</td><td>{{ w.name or '' }}</td><td>{{ w.bangla_name or '' }}</td><td>{{ w.department or '' }}</td><td>{{ w.designation or '' }}</td><td>{{ w.phone or '' }}</td><td>à§³ {{ '%.2f'|format((w.basic_salary or 0)|float) }}</td><td><span class="badge {{ 'active' if (w.status or 'Active')=='Active' else 'inactive' }}">{{ w.status or 'Active' }}</span></td><td class="actions"><a class="btn small" href="{{ url_for('edit_worker', worker_id=w.id) }}">Edit</a><form method="post" action="{{ url_for('toggle_worker_status', worker_id=w.id) }}" style="display:inline"><input type="hidden" name="q" value="{{ q or '' }}"><input type="hidden" name="department" value="{{ department or '' }}"><input type="hidden" name="status" value="{{ status or '' }}"><button class="btn small {{ 'secondary' if (w.status or 'Active')=='Inactive' else '' }}" type="submit" onclick="return confirm('Change status for {{ w.name }}?');">{{ 'Set Active' if (w.status or 'Active')=='Inactive' else 'Set Inactive' }}</button></form><a class="btn small secondary" href="{{ url_for('attendance', worker_id=w.id) }}">Attendance</a><a class="btn small" href="{{ url_for('payslip', worker_id=w.id) }}">Payslip</a><a class="btn small" href="{{ url_for('advance', worker_id=w.id) }}">Advance</a><a class="btn small secondary" href="{{ url_for('payments', worker_id=w.id) }}">Payment</a></td></tr>{% else %}<tr><td colspan="9">No matching worker found.</td></tr>{% endfor %}</tbody></table></div>
</div>
</main></body></html>
"""

def safe_render_template(template_name, **context):
    """Use normal Jinja template when present; otherwise use built-in fallback."""
    try:
        return render_template(template_name, **context)
    except Exception as exc:
        if isinstance(exc, TemplateNotFound) or template_name in {"dashboard.html", "attendance.html", "login.html", "payslip.html", "advance.html", "advance_edit.html", "worker.html", "worker_form.html", "department.html", "settings.html", "users.html", "user_form.html", "activity.html", "backup.html", "data_sync.html", "reports.html", "accounts.html", "ot_report.html", "worker_search.html"}:
            fallback = {
                "dashboard.html": BUILTIN_DASHBOARD_TEMPLATE,
                "attendance.html": BUILTIN_ATTENDANCE_TEMPLATE,
                "login.html": BUILTIN_LOGIN_TEMPLATE,
                "payslip.html": BUILTIN_PAYSLIP_TEMPLATE,
                    "advance.html": BUILTIN_ADVANCE_TEMPLATE,
                    "advance_edit.html": BUILTIN_ADVANCE_EDIT_TEMPLATE,
                "worker.html": BUILTIN_WORKER_TEMPLATE,
                "worker_form.html": BUILTIN_WORKER_FORM_TEMPLATE,
                "department.html": BUILTIN_DEPARTMENT_TEMPLATE,
                "settings.html": BUILTIN_SETTINGS_TEMPLATE,
                "users.html": BUILTIN_USERS_TEMPLATE,
                "user_form.html": BUILTIN_USER_FORM_TEMPLATE,
                "activity.html": BUILTIN_ACTIVITY_TEMPLATE,
                "backup.html": BUILTIN_BACKUP_TEMPLATE,
                 "data_sync.html": BUILTIN_DATA_SYNC_TEMPLATE,
                "reports.html": BUILTIN_REPORTS_TEMPLATE,
                "accounts.html": BUILTIN_ACCOUNTS_TEMPLATE,
        "payments.html": BUILTIN_PAYMENTS_TEMPLATE,
        "payment_history.html": BUILTIN_PAYMENT_HISTORY_TEMPLATE,
        "payroll_closing.html": BUILTIN_PAYROLL_CLOSING_TEMPLATE,
        "ot_report.html": BUILTIN_OT_REPORT_TEMPLATE,
                "worker_search.html": BUILTIN_WORKER_SEARCH_TEMPLATE,
            }.get(template_name)
            if fallback:
                app.logger.warning("%s could not be loaded: %r. Using built-in fallback.", template_name, exc)
                return render_template_string(fallback, **context)
        raise


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
        month
    )

    gross = sum(
        salary["gross"]
        for salary in salaries.values()
    )

    advance = sum(
        salary["advance"]
        for salary in salaries.values()
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
            today.day,
        ),
    )

    today_map = {
        row["status"]: row["c"]
        for row in today_rows
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

    return safe_render_template(
        "dashboard.html",
        workers=workers,
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
        total_workers=len(workers),
    )


# ============================================================
# WORKER LIST
# ============================================================

@app.route("/workers")
@login_required
def workers():

    query_text = request.args.get(
        "q",
        ""
    ).strip()

    status_filter = request.args.get(
        "status",
        ""
    ).strip()
    if status_filter not in ("", "Active", "Inactive"):
        status_filter = ""

    params = []
    conditions = []

    if query_text:
        like_value = f"%{query_text.lower()}%"
        conditions.append("""
            (CAST(id AS TEXT) LIKE ?
             OR LOWER(COALESCE(name, '')) LIKE ?
             OR LOWER(COALESCE(bangla_name, '')) LIKE ?
             OR LOWER(COALESCE(department, '')) LIKE ?
             OR LOWER(COALESCE(designation, '')) LIKE ?)
        """)
        params.extend([like_value] * 5)

    if status_filter:
        conditions.append("COALESCE(NULLIF(TRIM(status), ''), 'Active')=?")
        params.append(status_filter)

    sql = "SELECT * FROM workers"
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY id DESC"

    worker_rows = fetch_all(
        sql,
        params
    )

    return safe_render_template(
        "worker.html",
        workers=worker_rows,
        q=query_text,
        status=status_filter,
    )


# ============================================================
# ADD WORKER
# ============================================================

@app.route("/workers/toggle-status/<int:worker_id>", methods=["POST"])
@login_required
def toggle_worker_status(worker_id):
    worker = fetch_one("SELECT id,name,status FROM workers WHERE id=?", (worker_id,))
    if not worker:
        flash("Worker not found.", "danger")
        return redirect(url_for("workers"))
    old_status = (worker.get("status") or "Active").strip()
    new_status = "Inactive" if old_status == "Active" else "Active"
    execute("UPDATE workers SET status=? WHERE id=?", (new_status, worker_id), commit=True)
    log_activity(current_user().get("username") if current_user() else "system", f"Worker status changed: {worker['name']} ({old_status} -> {new_status})")
    flash(f"{worker['name']} is now {new_status}.", "success")
    # Preserve the search/filter state when status is changed from Worker Search.
    return redirect(url_for(
        "worker_search",
        q=request.form.get("q", ""),
        department=request.form.get("department", ""),
        status=request.form.get("status", ""),
    ) if request.referrer and "/worker-search" in request.referrer else url_for(
        "workers",
        q=request.form.get("q", ""),
        status=request.form.get("status", ""),
    ))


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

        bangla_name = form.get(
            "bangla_name",
            ""
        ).strip()

        status = form.get("status", "Active").strip() or "Active"
        if status not in ("Active", "Inactive"):
            status = "Active"

        department = form.get(
            "department",
            ""
        ).strip()

        designation = form.get(
            "designation",
            ""
        ).strip()

        basic_salary = parse_num(
            form.get("basic_salary"),
            0
        )

        # Auto OT Rate for new worker
        # Formula = Basic Salary / 30 / 12
        # Rounded to nearest whole BDT
        ot_rate = int((basic_salary / 30 / 12) + 0.5) if basic_salary > 0 else 0

        refreshment_bill = parse_num(
            form.get(
                "refreshment_bill"
            ),
            0
        )

        if not name:

            flash(
                "Worker name is required.",
                "danger"
            )

            return safe_render_template(
                "worker_form.html",
                worker=form,
                departments=(
                    get_worker_departments()
                ),
                form_title="Add Worker",
                submit_text="Save",
            )

        if (
            department
            not in DEPARTMENTS
            or department
            == "All Departments"
        ):

            department = "General"

        try:

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
                    bangla_name,
                    status
                )
                VALUES
                (?,?,?,?,?,?,?,?)
                """,
                (
                    name,
                    basic_salary,
                    ot_rate,
                    department,
                    designation,
                    refreshment_bill,
                    bangla_name,
                    status,
                ),
                commit=True,
            )

            log_activity(
                current_user().get("username") if current_user() else "system",
                f"Worker added: {name}"
            )

            flash(
                "Worker saved successfully.",
                "success"
            )

            return redirect(
                url_for("workers")
            )

        except Exception as e:

            app.logger.exception(
                "Worker insert error"
            )

            flash(
                f"Could not save worker: {e}",
                "danger"
            )

            return safe_render_template(
                "worker_form.html",
                worker=form,
                departments=(
                    get_worker_departments()
                ),
                form_title="Add Worker",
                submit_text="Save",
            )

    return safe_render_template(
        "worker_form.html",
        worker=None,
        departments=(
            get_worker_departments()
        ),
        form_title="Add Worker",
        submit_text="Save",
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

        name = form.get(
            "name",
            ""
        ).strip()

        bangla_name = form.get(
            "bangla_name",
            ""
        ).strip()

        status = form.get("status", "Active").strip() or "Active"
        if status not in ("Active", "Inactive"):
            status = "Active"

        department = form.get(
            "department",
            ""
        ).strip()

        designation = form.get(
            "designation",
            ""
        ).strip()

        basic_salary = parse_num(
            form.get("basic_salary"),
            0
        )

        # OT Rate is recalculated automatically when Basic Salary changes.
        # If Basic Salary is unchanged, keep the manually entered OT Rate.
        submitted_ot_rate = parse_num(
            form.get("ot_rate"),
            0
        )

        old_basic_salary = parse_num(
            worker.get("basic_salary"),
            0
        )

        if abs(basic_salary - old_basic_salary) > 0.000001:
            ot_rate = (
                int((basic_salary / 30 / 12) + 0.5)
                if basic_salary > 0
                else 0
            )
        else:
            ot_rate = submitted_ot_rate

        refreshment_bill = parse_num(
            form.get(
                "refreshment_bill"
            ),
            0
        )

        if not name:

            flash(
                "Worker name is required.",
                "danger"
            )

            return safe_render_template(
                "worker_form.html",
                worker=form,
                departments=(
                    get_worker_departments()
                ),
                form_title="Edit Worker",
                submit_text="Update",
            )

        if (
            department
            not in DEPARTMENTS
            or department
            == "All Departments"
        ):

            department = "General"

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
                    bangla_name=?,
                    status=?
                WHERE id=?
                """,
                (
                    name,
                    basic_salary,
                    ot_rate,
                    department,
                    designation,
                    refreshment_bill,
                    bangla_name,
                    status,
                    worker_id,
                ),
                commit=True,
            )

            log_activity(
                current_user().get("username") if current_user() else "system",
                f"Worker updated: ID {worker_id} - {name}"
            )

            flash(
                "Worker updated successfully.",
                "success"
            )

            return redirect(
                url_for("workers")
            )

        except Exception as e:

            app.logger.exception(
                "Worker update error"
            )

            flash(
                f"Could not update worker: {e}",
                "danger"
            )

            return safe_render_template(
                "worker_form.html",
                worker=form,
                departments=(
                    get_worker_departments()
                ),
                form_title="Edit Worker",
                submit_text="Update",
            )

    return safe_render_template(
        "worker_form.html",
        worker=worker,
        departments=(
            get_worker_departments()
        ),
        form_title="Edit Worker",
        submit_text="Update",
    )


# ============================================================
# DELETE WORKER
# ============================================================

@app.route(
    "/workers/delete/<int:worker_id>",
    methods=["GET", "POST"]
)
@login_required
def delete_worker(worker_id):

    worker = fetch_one(
        """
        SELECT id, name, bangla_name, department
        FROM workers
        WHERE id=?
        """,
        (worker_id,),
    )

    if not worker:
        flash("Worker not found.", "danger")
        return redirect(url_for("workers"))

    if request.method == "GET":
        return f"""
        <!doctype html><html><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"><title>Confirm Worker Delete</title>
        <style>body{{font-family:Arial,sans-serif;background:#f3f6fb;margin:0;padding:40px 15px;color:#1e293b}}.box{{max-width:620px;margin:40px auto;background:#fff;border:1px solid #e5eaf2;border-radius:14px;padding:28px;box-shadow:0 6px 24px rgba(15,23,42,.08)}}h2{{margin-top:0}}.warning{{background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;padding:14px;border-radius:8px;margin:18px 0}}.btn{{display:inline-block;border:0;border-radius:8px;padding:10px 16px;text-decoration:none;cursor:pointer;font-weight:600;margin-right:8px}}.danger{{background:#dc2626;color:#fff}}.cancel{{background:#e2e8f0;color:#334155}}</style>
        </head><body><div class=\"box\"><h2>Confirm Worker Delete</h2>
        <p><strong>Worker ID:</strong> {worker.get("id")}</p><p><strong>English Name:</strong> {worker.get("name") or ""}</p><p><strong>Bangla Name:</strong> {worker.get("bangla_name") or ""}</p><p><strong>Department:</strong> {worker.get("department") or ""}</p>
        <div class=\"warning\">This action can affect payroll-related records. The system will NOT delete this worker if attendance, daily attendance, or advance records already exist.</div>
        <form method=\"post\" style=\"display:inline\"><button class=\"btn danger\" type=\"submit\">Confirm Delete</button></form><a class=\"btn cancel\" href=\"{url_for("workers")}\">Cancel</a>
        </div>
</body></html>
        """

    dependent_tables = ["attendance", "daily_attendance", "worker_advances", "advance_salary", "advances"]
    try:
        for table_name in dependent_tables:
            if not table_exists(table_name):
                continue
            table_columns = columns(table_name)
            if "worker_id" not in table_columns:
                continue
            dependent = fetch_one(f"SELECT 1 FROM {table_name} WHERE worker_id=? LIMIT 1", (worker_id,))
            if dependent:
                flash(f"Worker cannot be deleted because payroll/attendance records already exist in {table_name}. Please keep the worker record to preserve history.", "danger")
                return redirect(url_for("workers"))

        execute("DELETE FROM workers WHERE id=?", (worker_id,), commit=True)
        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Worker deleted: ID {worker_id}"
        )
        flash("Worker deleted successfully.", "success")
    except Exception as e:
        app.logger.exception("Worker delete error")
        flash(f"Could not delete worker: {e}", "danger")

    return redirect(url_for("workers"))


# ============================================================
# ATTENDANCE
# ============================================================

@app.route("/attendance")
@login_required
def attendance():

    requested_month = request.args.get(
        "month"
    )

    requested_year = request.args.get(
        "year"
    )

    if (
        requested_month
        and requested_year
    ):

        month = month_name_year(
            requested_month,
            requested_year
        )

    else:

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
            (worker_id,),
        )

        if worker:

            daily = daily_map(
                worker["id"],
                month
            )

            month_name, year_text = (
                month.split()
            )

            year = int(
                year_text
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
                    "status": daily.get(
                        day,
                        ""
                    ),
                    "date": datetime.date(
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

            # Live daily summary: P/A records in the calendar are the
            # source for Present/Absent counts. Not Set days are ignored.
            daily_present = 0
            daily_absent = 0
            for day_number, status in daily.items():
                try:
                    current_date = datetime.date(year, month_number, int(day_number))
                    if current_date.weekday() == 4:
                        continue
                except Exception:
                    continue
                if status == "P":
                    daily_present += 1
                elif status == "A":
                    daily_absent += 1
            if daily:
                summary["present"] = daily_present
                summary["absent"] = daily_absent
                summary["present_days"] = daily_present
                summary["absent_days"] = daily_absent

    return safe_render_template(
        "attendance.html",
        workers=fetch_all(
            """
            SELECT
                id,
                name,
                bangla_name,
                department
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
    """Save monthly and daily attendance safely in one transaction."""

    form = request.form

    # --------------------------------------------------------
    # Worker validation
    # --------------------------------------------------------
    try:
        worker_id = int(form.get("worker_id") or 0)
    except (TypeError, ValueError):
        flash("Invalid worker.", "danger")
        return redirect(url_for("attendance"))

    worker_exists = fetch_one(
        "SELECT id FROM workers WHERE id=?",
        (worker_id,),
    )

    if not worker_exists:
        flash("Worker not found.", "danger")
        return redirect(url_for("attendance"))

    # --------------------------------------------------------
    # Month validation
    # --------------------------------------------------------
    try:
        if form.get("year"):
            month = month_name_year(
                form.get("month"),
                form.get("year")
            )
        else:
            month = form.get("month_year") or month_name_year()

        month_name, year_text = month.split()
        attendance_year = int(year_text)
        attendance_month = MONTHS.index(month_name) + 1
        days_in_month = calendar.monthrange(
            attendance_year,
            attendance_month
        )[1]
    except Exception:
        flash("Invalid attendance month.", "danger")
        return redirect(url_for("attendance", worker_id=worker_id))

    if not require_payroll_unlocked(month):
        return redirect(url_for("attendance", worker_id=worker_id, month=attendance_month, year=attendance_year))

    # --------------------------------------------------------
    # Monthly value validation
    # --------------------------------------------------------
    try:
        present = int(form.get("present_days") or 0)
        absent = int(form.get("absent_days") or 0)
    except (TypeError, ValueError):
        flash("Present and absent days must be whole numbers.", "danger")
        return redirect(url_for(
            "attendance",
            worker_id=worker_id,
            month=attendance_month,
            year=attendance_year,
        ))

    try:
        overtime = parse_num(form.get("ot_hours"), 0)
    except Exception:
        overtime = 0

    if present < 0 or absent < 0 or overtime < 0:
        flash("Attendance values cannot be negative.", "danger")
        return redirect(url_for(
            "attendance",
            worker_id=worker_id,
            month=attendance_month,
            year=attendance_year,
        ))

    if present + absent > days_in_month:
        flash(
            f"Present + absent days cannot exceed {days_in_month} days.",
            "danger"
        )
        return redirect(url_for(
            "attendance",
            worker_id=worker_id,
            month=attendance_month,
            year=attendance_year,
        ))

    # --------------------------------------------------------
    # Collect daily statuses first. Nothing is written until
    # every supplied day has passed validation.
    # --------------------------------------------------------
    daily_statuses = {}

    raw_json = form.get("statuses_json")
    if raw_json:
        try:
            import json
            parsed = json.loads(raw_json)
            if isinstance(parsed, dict):
                daily_statuses.update(parsed)
        except Exception:
            flash("Invalid daily attendance data.", "danger")
            return redirect(url_for(
                "attendance",
                worker_id=worker_id,
                month=attendance_month,
                year=attendance_year,
            ))

    for key, value in form.items():
        if key.startswith("day_"):
            daily_statuses[key[4:]] = value

    normalized_daily = {}
    unset_daily = set()

    for day, status in daily_statuses.items():
        try:
            day_number = int(day)
        except (TypeError, ValueError):
            flash("Invalid attendance day.", "danger")
            return redirect(url_for(
                "attendance",
                worker_id=worker_id,
                month=attendance_month,
                year=attendance_year,
            ))

        if day_number < 1 or day_number > days_in_month:
            flash(
                f"Attendance day must be between 1 and {days_in_month}.",
                "danger"
            )
            return redirect(url_for(
                "attendance",
                worker_id=worker_id,
                month=attendance_month,
                year=attendance_year,
            ))

        status_text = str(status or "").strip().upper()

        # Friday is the weekly holiday. Keep it as Absent in the
        # attendance calendar, while salary calculation separately
        # ignores Friday for absent deduction.
        try:
            current_date = datetime.date(
                attendance_year,
                attendance_month,
                day_number,
            )
            if current_date.weekday() == 4:
                status_text = "A"
        except Exception:
            pass

        if not status_text or status_text in {"-", "NOT SET", "NONE"}:
            # Explicitly clear an existing daily attendance record.
            # A blank/Not Set selection means the user wants that day's P/A record removed, if one exists.
            unset_daily.add(day_number)
            continue
        if status_text.startswith("A"):
            normalized_daily[day_number] = "A"
        elif status_text.startswith("P"):
            normalized_daily[day_number] = "P"
        else:
            flash("Invalid daily attendance status.", "danger")
            return redirect(url_for(
                "attendance",
                worker_id=worker_id,
                month=attendance_month,
                year=attendance_year,
            ))

    # --------------------------------------------------------
    # Daily -> Monthly synchronization
    # --------------------------------------------------------
    # If the request contains daily attendance fields, the monthly
    # Present/Absent totals will be recalculated from the COMPLETE
    # saved daily calendar after the daily changes are applied.
    # This is important when a partial form/API request changes only
    # one or a few days: existing P/A days must not be lost.
    # --------------------------------------------------------
    sync_monthly_from_daily = bool(daily_statuses)

    # --------------------------------------------------------
    # One transaction for the complete save.
    # If any database operation fails, ALL attendance changes
    # from this request are rolled back together.
    # --------------------------------------------------------
    conn = None
    cur = None

    try:
        conn = db_connect()
        cur = conn.cursor()

        sql_prefix = placeholders("SELECT id FROM daily_attendance WHERE worker_id=? AND month_year=? AND day=? ORDER BY id DESC LIMIT 1")

        # Explicitly selected Not Set days: remove any existing P/A record.
        for day_number in unset_daily:
            cur.execute(
                placeholders("DELETE FROM daily_attendance WHERE worker_id=? AND month_year=? AND day=?"),
                (worker_id, month, day_number),
            )

        for day_number, status in normalized_daily.items():
            cur.execute(
                sql_prefix,
                (worker_id, month, day_number),
            )
            existing_row = cur.fetchone()

            if existing_row:
                existing_id = existing_row[0]
                cur.execute(
                    placeholders("UPDATE daily_attendance SET status=? WHERE id=?"),
                    (status, existing_id),
                )
            else:
                cur.execute(
                    placeholders("""
                        INSERT INTO daily_attendance
                        (worker_id, month_year, day, status)
                        VALUES (?,?,?,?)
                    """),
                    (worker_id, month, day_number, status),
                )

        # Recalculate from all saved daily records, not only from the
        # fields submitted in this request. This preserves existing
        # attendance when a partial update is submitted.
        if sync_monthly_from_daily:
            cur.execute(
                placeholders("""
                    SELECT day, status
                    FROM daily_attendance
                    WHERE worker_id=? AND month_year=?
                """),
                (worker_id, month),
            )
            present = 0
            absent = 0
            for row in cur.fetchall():
                try:
                    day_number = int(row[0])
                    status_value = str(row[1] or "").upper()
                    # Friday is the weekly holiday and never counts as
                    # Present or Absent for monthly salary attendance.
                    current_date = datetime.date(year, month_number, day_number)
                    if current_date.weekday() == 4:
                        continue
                    if status_value == "P":
                        present += 1
                    elif status_value == "A":
                        absent += 1
                except Exception:
                    continue

        cur.execute(
            placeholders("""
                SELECT id
                FROM attendance
                WHERE worker_id=?
                AND month_year=?
                ORDER BY id DESC
                LIMIT 1
            """),
            (worker_id, month),
        )
        existing = cur.fetchone()

        if existing:
            cur.execute(
                placeholders("""
                    UPDATE attendance
                    SET present_days=?, absent_days=?, ot_hours=?
                    WHERE id=?
                """),
                (present, absent, overtime, existing[0]),
            )
        else:
            attendance_columns = columns("attendance")
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
                overtime,
            ]

            if "advance_deduction" in attendance_columns:
                names.append("advance_deduction")
                values.append(0)

            cur.execute(
                placeholders(
                    "INSERT INTO attendance (" + ",".join(names) + ") VALUES (" +
                    ",".join(["?"] * len(values)) + ")"
                ),
                values,
            )

        conn.commit()

    except Exception as e:
        try:
            if conn:
                conn.rollback()
        except Exception:
            pass

        app.logger.exception("Attendance save error")
        flash(f"Could not save attendance: {e}", "danger")
        return redirect(url_for(
            "attendance",
            worker_id=worker_id,
            month=attendance_month,
            year=attendance_year,
        ))

    finally:
        try:
            if cur:
                cur.close()
        except Exception:
            pass
        db_release(conn)

    log_activity(
        current_user().get("username") if current_user() else "system",
        f"Attendance saved: Worker {worker_id}, {month}"
    )

    flash("Attendance saved successfully.", "success")

    return redirect(url_for(
        "attendance",
        worker_id=worker_id,
        month=month.split()[0],
        year=month.split()[1],
    ))


# ============================================================
# PAYSLIP
# ============================================================

@app.route("/payslip")
@login_required
def payslip():

    # Respect a month/year selected in the payslip form or URL.
    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

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
            (worker_id,),
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

    return safe_render_template(
        "payslip.html",
        workers=fetch_all(
            """
            SELECT
                id,
                name,
                bangla_name,
                department,
                basic_salary
            FROM workers
            ORDER BY id
            """
        ),
        worker=worker,
        summary=summary,
        month=month,
    )



# ============================================================
# ACCOUNTS / FINANCIAL SUMMARY
# ============================================================

@app.route("/accounts")
@login_required
def accounts():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    workers_list = fetch_all("SELECT * FROM workers ORDER BY id")
    salary_map = calculate_salary_bulk(workers_list, month)
    totals = {"workers":0,"basic":0,"present":0,"absent":0,"ot":0,"absent_cut":0,"ot_amt":0,"nasta":0,"gross":0,"advance":0,"net":0}
    department_map = {}
    for worker in workers_list:
        row = dict(worker); row.update(salary_map.get(worker["id"], {}))
        dept = worker.get("department") or "Unassigned"
        if dept not in department_map:
            department_map[dept] = {"department":dept,"workers":0,"basic":0,"gross":0,"advance":0,"net":0}
        d = department_map[dept]; d["workers"] += 1; totals["workers"] += 1
        for key, src in [("basic","basic_salary"),("gross","gross"),("advance","advance"),("net","net")]:
            val = float(row.get(src) or 0); d[key] += val; totals[key] += val
        for key, src in [("present","present"),("absent","absent"),("ot","ot"),("absent_cut","absent_cut"),("ot_amt","ot_amt"),("nasta","nasta")]:
            totals[key] += float(row.get(src) or 0)
    return safe_render_template("accounts.html", totals=totals, departments_summary=sorted(department_map.values(), key=lambda x:str(x["department"]).lower()), month=month, months=MONTHS, years=list(range(datetime.date.today().year-2, datetime.date.today().year+3)), settings=get_settings())

@app.route("/accounts/export")
@login_required
def accounts_export():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    workers_list = fetch_all("SELECT * FROM workers ORDER BY id")
    salary_map = calculate_salary_bulk(workers_list, month)
    output = io.StringIO(); writer = csv.writer(output)
    writer.writerow(["Department","Workers","Basic Salary","Gross Salary","Advance","Net Payable"])
    dep = {}
    for w in workers_list:
        row=dict(w); row.update(salary_map.get(w["id"], {})); name=w.get("department") or "Unassigned"
        d=dep.setdefault(name,[0,0,0,0,0]); d[0]+=1; d[1]+=float(row.get("basic_salary") or 0); d[2]+=float(row.get("gross") or 0); d[3]+=float(row.get("advance") or 0); d[4]+=float(row.get("net") or 0)
    for name,d in sorted(dep.items()): writer.writerow([name,d[0],f"{d[1]:.2f}",f"{d[2]:.2f}",f"{d[3]:.2f}",f"{d[4]:.2f}"])
    writer.writerow([]); writer.writerow(["Month",month]); output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")), mimetype="text/csv", as_attachment=True, download_name=f"accounts_{month.replace(' ','_')}.csv")

# ============================================================
# PAYMENT MANAGEMENT
# ============================================================

@app.route("/payments")
@login_required
def payments():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    workers_list = fetch_all("SELECT * FROM workers ORDER BY id")
    salary_map = calculate_salary_bulk(workers_list, month)
    payment_rows = fetch_all(
        "SELECT * FROM payroll_payments WHERE month_year=?",
        (month,),
    ) if table_exists("payroll_payments") else []
    payment_map = {row["worker_id"]: row for row in payment_rows}
    rows = []
    totals = {"workers":0,"net":0.0,"paid":0.0,"due":0.0,"paid_count":0,"unpaid_count":0}
    for worker in workers_list:
        salary = salary_map.get(worker["id"], {})
        net = float(salary.get("net") or 0)
        p = payment_map.get(worker["id"], {})
        paid = float(p.get("paid_amount") or 0)
        if paid < 0: paid = 0
        if paid > net: paid = net
        due = max(net - paid, 0)
        status = "Paid" if net > 0 and paid >= net else ("Partial" if paid > 0 else "Unpaid")
        row = dict(worker)
        row.update({"net":net,"paid_amount":paid,"due":due,"status":status,
                    "payment_date":p.get("payment_date") if p else "",
                    "payment_method":p.get("payment_method") if p else "Cash",
                    "note":p.get("note") if p else ""})
        rows.append(row)
        totals["workers"] += 1; totals["net"] += net; totals["paid"] += paid; totals["due"] += due
        if status == "Paid": totals["paid_count"] += 1
        else: totals["unpaid_count"] += 1
    return safe_render_template("payments.html", rows=rows, totals=totals, month=month,
        months=MONTHS, years=list(range(datetime.date.today().year-2, datetime.date.today().year+3)),
        settings=get_settings())


@app.route("/payments/save", methods=["POST"])
@login_required
def save_payment():
    worker_id = request.form.get("worker_id", "").strip()
    month = month_name_year(request.form.get("month"), request.form.get("year"))
    if not require_payroll_unlocked(month):
        return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))
    try:
        worker_id_int = int(worker_id)
    except Exception:
        flash("Invalid worker ID.", "danger"); return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))
    worker = fetch_one("SELECT * FROM workers WHERE id=?", (worker_id_int,))
    if not worker:
        flash("Worker not found.", "danger"); return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))
    salary = calculate_salary(worker, month)
    net = float(salary.get("net") or 0)
    try:
        paid = float(request.form.get("paid_amount") or 0)
    except Exception:
        paid = -1
    if paid < 0 or paid > net:
        flash(f"Paid amount must be between 0 and {net:.2f}.", "danger")
        return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))
    payment_date = request.form.get("payment_date", "").strip() or datetime.date.today().isoformat()
    method = request.form.get("payment_method", "Cash").strip() or "Cash"
    note = request.form.get("note", "").strip()
    status = "Paid" if net > 0 and paid >= net else ("Partial" if paid > 0 else "Unpaid")
    existing = fetch_one("SELECT id FROM payroll_payments WHERE worker_id=? AND month_year=?", (worker_id_int, month))
    if existing:
        execute("""UPDATE payroll_payments SET net_payable=?,paid_amount=?,payment_date=?,payment_method=?,status=?,note=?,updated_at=? WHERE id=?""",
                (net,paid,payment_date,method,status,note,nowstr(),existing["id"]),commit=True)
    else:
        execute("""INSERT INTO payroll_payments (worker_id,month_year,net_payable,paid_amount,payment_date,payment_method,status,note,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (worker_id_int,month,net,paid,payment_date,method,status,note,nowstr(),nowstr()),commit=True)

    # Keep payment history in sync with the monthly payment record.
    # IMPORTANT: editing/saving the same monthly payment must not create a
    # duplicate transaction. One worker + month has one current payment
    # transaction in this history table.
    existing_tx = fetch_one(
        "SELECT id FROM payment_transactions WHERE worker_id=? AND month_year=? ORDER BY id DESC LIMIT 1",
        (worker_id_int, month),
    )

    if existing_tx:
        execute(
            """UPDATE payment_transactions
               SET amount=?, payment_date=?, payment_method=?, note=?
               WHERE id=?""",
            (paid, payment_date, method, note, existing_tx["id"]),
            commit=True,
        )
    else:
        execute(
            """INSERT INTO payment_transactions
               (worker_id,month_year,amount,payment_date,payment_method,note,created_at)
               VALUES (?,?,?,?,?,?,?)""",
            (worker_id_int, month, paid, payment_date, method, note, nowstr()),
            commit=True,
        )

    log_activity(
        current_user().get("username") if current_user() else "system",
        f"Payment saved: Worker {worker_id_int}, {month}, BDT {paid:.2f}"
    )
    flash("Payment record saved successfully.", "success")
    return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))


@app.route("/payments/delete", methods=["POST"])
@login_required
def delete_payment():
    worker_id = request.form.get("worker_id", "").strip()
    month = month_name_year(request.form.get("month"), request.form.get("year"))
    if not require_payroll_unlocked(month):
        return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))
    try:
        worker_id_int = int(worker_id)
    except Exception:
        flash("Invalid worker ID.", "danger")
        return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))
    worker = fetch_one("SELECT id FROM workers WHERE id=?", (worker_id_int,))
    if not worker:
        flash("Worker not found.", "danger")
        return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))

    # Remove the current monthly payment and its matching history transaction.
    # Attendance, salary and advance records are never touched.
    if table_exists("payroll_payments"):
        execute("DELETE FROM payroll_payments WHERE worker_id=? AND month_year=?", (worker_id_int, month), commit=True)
    if table_exists("payment_transactions"):
        execute("DELETE FROM payment_transactions WHERE worker_id=? AND month_year=?", (worker_id_int, month), commit=True)

    log_activity(
        current_user().get("username") if current_user() else "system",
        f"Payment deleted: Worker {worker_id_int}, {month}"
    )
    flash("Payment record deleted successfully.", "success")
    return redirect(url_for("payments", month=month.split()[0], year=month.split()[1]))


@app.route("/payments/export")
@login_required
def payments_export():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    workers_list = fetch_all("SELECT * FROM workers ORDER BY id")
    salary_map = calculate_salary_bulk(workers_list, month)
    payment_rows = fetch_all("SELECT * FROM payroll_payments WHERE month_year=?", (month,))
    pmap = {r["worker_id"]: r for r in payment_rows}
    output=io.StringIO(); writer=csv.writer(output)
    writer.writerow(["Worker ID","Worker Name","Department","Net Payable","Paid Amount","Due","Status","Payment Date","Payment Method","Note"])
    for w in workers_list:
        net=float(salary_map.get(w["id"],{}).get("net") or 0); p=pmap.get(w["id"],{}); paid=min(max(float(p.get("paid_amount") or 0),0),net); due=max(net-paid,0)
        status="Paid" if net>0 and paid>=net else ("Partial" if paid>0 else "Unpaid")
        writer.writerow([w.get("id"),w.get("name") or "",w.get("department") or "",f"{net:.2f}",f"{paid:.2f}",f"{due:.2f}",status,p.get("payment_date") or "",p.get("payment_method") or "Cash",p.get("note") or ""])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")),mimetype="text/csv",as_attachment=True,download_name=f"payments_{month.replace(' ','_')}.csv")



@app.route("/payment-history")
@login_required
def payment_history():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    worker_id = request.args.get("worker_id", "").strip()
    params = [month]
    where = ["pt.month_year=?"]
    if worker_id:
        try:
            int(worker_id)
            where.append("pt.worker_id=?")
            params.append(int(worker_id))
        except Exception:
            worker_id = ""
    query = """SELECT pt.*, w.name, w.bangla_name, w.department FROM payment_transactions pt LEFT JOIN workers w ON w.id=pt.worker_id WHERE """ + " AND ".join(where) + " ORDER BY pt.id DESC"
    rows = fetch_all(query, params)
    totals = {"count":len(rows),"amount":sum(float(r.get("amount") or 0) for r in rows),"workers":len({r.get("worker_id") for r in rows})}
    return safe_render_template("payment_history.html", rows=rows, totals=totals, month=month, worker_id=worker_id, months=MONTHS, years=list(range(datetime.date.today().year-2, datetime.date.today().year+3)), settings=get_settings())


@app.route("/payment-history/export")
@login_required
def payment_history_export():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    worker_id = request.args.get("worker_id", "").strip()
    params=[month]; where=["pt.month_year=?"]
    if worker_id:
        try:
            params.append(int(worker_id)); where.append("pt.worker_id=?")
        except Exception:
            worker_id=""
    rows=fetch_all("SELECT pt.*, w.name, w.bangla_name, w.department FROM payment_transactions pt LEFT JOIN workers w ON w.id=pt.worker_id WHERE " + " AND ".join(where) + " ORDER BY pt.id DESC", params)
    output=io.StringIO(); writer=csv.writer(output)
    writer.writerow(["Transaction ID","Worker ID","Worker Name","Department","Amount","Payment Date","Payment Method","Note","Created At"])
    for r in rows:
        writer.writerow([r.get("id"),r.get("worker_id"),r.get("name") or "",r.get("department") or "",f"{float(r.get('amount') or 0):.2f}",r.get("payment_date") or "",r.get("payment_method") or "",r.get("note") or "",r.get("created_at") or ""])
    output.seek(0)
    return send_file(io.BytesIO(output.getvalue().encode("utf-8-sig")),mimetype="text/csv",as_attachment=True,download_name=f"payment_history_{month.replace(' ','_')}.csv")

# ============================================================
# PAYROLL CLOSING / MONTHLY LOCK
# ============================================================

@app.route("/payroll-closing")
@admin_required
def payroll_closing():
    month = month_name_year(request.args.get("month"), request.args.get("year"))
    selected_lock = payroll_lock_info(month)
    recent = []
    today = datetime.date.today()
    first = today.replace(day=1)
    for offset in range(-6, 7):
        total_month = first.month - 1 + offset
        year = first.year + total_month // 12
        month_num = total_month % 12 + 1
        m = f"{MONTHS[month_num-1]} {year}"
        lock = payroll_lock_info(m)
        recent.append({"month":m,"locked":bool(lock),"locked_at":lock.get("locked_at") if lock else "","locked_by":lock.get("locked_by") if lock else ""})
    return safe_render_template("payroll_closing.html", month=month, selected_lock=selected_lock, recent=recent, months=MONTHS, years=list(range(datetime.date.today().year-2,datetime.date.today().year+3)), settings=get_settings())

@app.route("/payroll-closing/lock", methods=["POST"])
@admin_required
def payroll_close():
    month = month_name_year(request.form.get("month"), request.form.get("year"))
    note = request.form.get("note", "").strip()
    user = current_user() or {}
    username = user.get("username") or user.get("full_name") or "Administrator"
    if not is_payroll_locked(month):
        if is_postgres():
            execute("INSERT INTO payroll_locks (month_year,locked_at,locked_by,note) VALUES (?,?,?,?) ON CONFLICT(month_year) DO NOTHING", (month,nowstr(),username,note), commit=True)
        else:
            execute("INSERT OR IGNORE INTO payroll_locks (month_year,locked_at,locked_by,note) VALUES (?,?,?,?)", (month,nowstr(),username,note), commit=True)
        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Payroll locked: {month}"
        )
        flash(f"Payroll for {month} has been locked.", "success")
    else:
        flash(f"Payroll for {month} is already locked.", "danger")
    return redirect(url_for("payroll_closing", month=month.split()[0], year=month.split()[1]))

@app.route("/payroll-closing/unlock", methods=["POST"])
@admin_required
def payroll_unlock():
    month = month_name_year(request.form.get("month"), request.form.get("year"))
    execute("DELETE FROM payroll_locks WHERE month_year=?", (month,), commit=True)
    log_activity(
        current_user().get("username") if current_user() else "system",
        f"Payroll unlocked: {month}"
    )
    flash(f"Payroll for {month} has been unlocked.", "success")
    return redirect(url_for("payroll_closing", month=month.split()[0], year=month.split()[1]))


# ============================================================
# OT REPORT
# ============================================================

@app.route("/ot-report")
@login_required
def ot_report():

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

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
    totals = {"workers": 0, "hours": 0.0, "amount": 0.0}
    department_map = {}

    for worker in workers_list:
        salary = salary_map.get(worker["id"], {})
        ot_hours = float(salary.get("ot") or 0)
        ot_amount = float(salary.get("ot_amt") or 0)

        # Keep all workers in the detailed table, but count only workers with OT.
        row = dict(worker)
        row.update(salary)
        rows.append(row)

        if ot_hours or ot_amount:
            totals["workers"] += 1

        totals["hours"] += ot_hours
        totals["amount"] += ot_amount

        dept = worker.get("department") or "Unassigned"
        if dept not in department_map:
            department_map[dept] = {"department": dept, "workers": 0, "hours": 0.0, "amount": 0.0}
        department_map[dept]["hours"] += ot_hours
        department_map[dept]["amount"] += ot_amount
        if ot_hours or ot_amount:
            department_map[dept]["workers"] += 1

    departments_summary = sorted(
        department_map.values(),
        key=lambda x: str(x["department"]).lower(),
    )

    return safe_render_template(
        "ot_report.html",
        rows=rows,
        totals=totals,
        departments_summary=departments_summary,
        month=month,
        months=MONTHS,
        years=list(range(datetime.date.today().year - 2, datetime.date.today().year + 3)),
        settings=get_settings(),
    )


@app.route("/ot-report/export")
@login_required
def ot_report_export():

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    workers_list = fetch_all(
        "SELECT * FROM workers ORDER BY id"
    )
    salary_map = calculate_salary_bulk(workers_list, month)

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Worker ID",
        "Worker Name",
        "Department",
        "Designation",
        "Basic Salary",
        "OT Rate",
        "OT Hours",
        "OT Amount",
        "Payroll Month",
    ])

    for worker in workers_list:
        salary = salary_map.get(worker["id"], {})
        writer.writerow([
            worker.get("id", ""),
            worker.get("name", ""),
            worker.get("department", ""),
            worker.get("designation", ""),
            float(worker.get("basic_salary") or 0),
            float(worker.get("ot_rate") or 0),
            float(salary.get("ot") or 0),
            float(salary.get("ot_amt") or 0),
            month,
        ])

    output.seek(0)
    filename = "OT_Report_" + month.replace(" ", "_") + ".csv"
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name=filename,
    )


# REPORTS
# ============================================================

@app.route("/reports")
@login_required
def reports():

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    workers_list = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    salary_map = {
        worker["id"]: salary_for_reporting(worker, month)
        for worker in workers_list
    }

    rows = []
    totals = {
        "workers": 0,
        "basic": 0,
        "present": 0,
        "absent": 0,
        "ot": 0,
        "absent_cut": 0,
        "ot_amt": 0,
        "nasta": 0,
        "gross": 0,
        "advance": 0,
        "net": 0,
    }

    department_map = {}

    for worker in workers_list:

        salary = salary_map.get(worker["id"], {})
        row = dict(worker)
        row.update(salary)
        rows.append(row)

        department_name = (
            worker.get("department")
            or "Unassigned"
        )

        if department_name not in department_map:
            department_map[department_name] = {
                "department": department_name,
                "workers": 0,
                "basic": 0,
                "present": 0,
                "absent": 0,
                "ot": 0,
                "absent_cut": 0,
                "ot_amt": 0,
                "nasta": 0,
                "gross": 0,
                "advance": 0,
                "net": 0,
            }

        target = department_map[department_name]
        target["workers"] += 1

        for key, source_key in [
            ("basic", "basic_salary"),
            ("present", "present"),
            ("absent", "absent"),
            ("ot", "ot"),
            ("absent_cut", "absent_cut"),
            ("ot_amt", "ot_amt"),
            ("nasta", "nasta"),
            ("gross", "gross"),
            ("advance", "advance"),
            ("net", "net"),
        ]:
            value = float(row.get(source_key) or 0)
            target[key] += value
            totals[key] += value

        totals["workers"] += 1

    departments_summary = sorted(
        department_map.values(),
        key=lambda item: str(item["department"]).lower()
    )

    return safe_render_template(
        "reports.html",
        rows=rows,
        totals=totals,
        departments_summary=departments_summary,
        month=month,
        months=MONTHS,
        years=list(range(
            datetime.date.today().year - 2,
            datetime.date.today().year + 3,
        )),
        settings=get_settings(),
    )


@app.route("/reports/export")
@login_required
def export_reports():

    if openpyxl is None:
        flash(
            "openpyxl is not installed. Excel export is unavailable.",
            "danger",
        )
        return redirect(url_for("reports"))

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    workers_list = fetch_all(
        """
        SELECT *
        FROM workers
        ORDER BY id
        """
    )

    salary_map = {
        worker["id"]: salary_for_reporting(worker, month)
        for worker in workers_list
    }

    wb = Workbook()
    ws = wb.active
    ws.title = "Payroll Report"

    company = get_settings().get(
        "company_name",
        DEFAULT_COMPANY_NAME,
    )

    ws.append([company])
    ws.append([f"Payroll Report - {month}"])
    ws.append([])
    ws.append([
        "Worker ID", "Worker Name", "Department",
        "Basic Salary", "Present", "Absent", "OT Hours",
        "Absent Deduction", "OT Amount", "Nasta",
        "Gross Salary", "Advance", "Net Payable",
    ])

    for worker in workers_list:
        salary = salary_map.get(worker["id"], {})
        ws.append([
            worker.get("id"),
            worker.get("name") or "",
            worker.get("department") or "",
            float(worker.get("basic_salary") or 0),
            float(salary.get("present") or 0),
            float(salary.get("absent") or 0),
            float(salary.get("ot") or 0),
            float(salary.get("absent_cut") or 0),
            float(salary.get("ot_amt") or 0),
            float(salary.get("nasta") or 0),
            float(salary.get("gross") or 0),
            float(salary.get("advance") or 0),
            float(salary.get("net") or 0),
        ])

    for cell in ws[4]:
        cell.font = cell.font.copy(bold=True)

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = ws.dimensions

    for column_cells in ws.columns:
        max_len = 0
        column_letter = column_cells[0].column_letter
        for cell in column_cells:
            max_len = max(
                max_len,
                len(str(cell.value or "")),
            )
        ws.column_dimensions[column_letter].width = min(
            max(max_len + 2, 10),
            28,
        )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    safe_month = month.replace(" ", "_")
    filename = f"Reedoy_Payroll_Report_{safe_month}.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ============================================================
# DEPARTMENT SALARY
# ============================================================

@app.route("/department")
@login_required
def department():

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    department_name = request.args.get(
        "department",
        ""
    ).strip()

    if department_name:
        workers_list = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (department_name,),
        )
    else:
        workers_list = fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    salary_map = {
        worker["id"]: salary_for_reporting(worker, month)
        for worker in workers_list
    }

    rows = []
    totals = {
        "workers": 0,
        "basic": 0,
        "present": 0,
        "absent": 0,
        "ot": 0,
        "absent_cut": 0,
        "ot_amt": 0,
        "nasta": 0,
        "gross": 0,
        "advance": 0,
        "net": 0,
    }

    for worker in workers_list:
        salary = salary_map.get(worker["id"], {})
        row = dict(worker)
        row.update(salary)
        rows.append(row)

        totals["workers"] += 1
        totals["basic"] += float(row.get("basic_salary") or 0)
        totals["present"] += float(row.get("present") or 0)
        totals["absent"] += float(row.get("absent") or 0)
        totals["ot"] += float(row.get("ot") or 0)
        totals["absent_cut"] += float(row.get("absent_cut") or 0)
        totals["ot_amt"] += float(row.get("ot_amt") or 0)
        totals["nasta"] += float(row.get("nasta") or 0)
        totals["gross"] += float(row.get("gross") or 0)
        totals["advance"] += float(row.get("advance") or 0)
        totals["net"] += float(row.get("net") or 0)

    department_list = fetch_all(
        """
        SELECT DISTINCT department
        FROM workers
        WHERE department IS NOT NULL
        AND department <> ''
        ORDER BY department
        """
    )

    return safe_render_template(
        "department.html",
        rows=rows,
        month=month,
        department=department_name,
        departments=[row["department"] for row in department_list],
        totals=totals,
        months=MONTHS,
        years=list(range(datetime.date.today().year - 2, datetime.date.today().year + 3)),
        settings=get_settings(),
    )

# ============================================================
# REPORT COMPATIBILITY
# ============================================================

@app.route("/report", methods=["GET"])
@login_required
def report():
    return department()


# ============================================================
# ADVANCE
# ============================================================

@app.route("/advance")
@app.route("/advances")
@login_required
def advance():

    # Respect the month/year selected by the user.
    # Previously this always used the current month, so loading an older
    # month from the Advance Salary filter did not actually change the data.
    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    worker_id = request.args.get(
        "worker_id"
    )

    total_advance = sum(parse_num(r.get("amount"), 0) for r in advance_rows(month, worker_id))

    return safe_render_template(
        "advance.html",
        workers=fetch_all(
            """
            SELECT
                id,
                name,
                bangla_name,
                department
            FROM workers
            ORDER BY id
            """
        ),
        rows=advance_rows(
            month,
            worker_id
        ),
        total_advance=total_advance,
        month=month,
        months=MONTHS,
        years=list(range(datetime.date.today().year - 2, datetime.date.today().year + 3)),
        today_date=datetime.date.today().isoformat(),
        settings=get_settings(),
    )


# Compatibility endpoint.

app.add_url_rule(
    "/advances",
    endpoint="advances",
    view_func=advance,
    methods=["GET"]
)


# ============================================================
# SAVE ADVANCE
# ============================================================

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

    try:

        worker_id = int(
            form.get("worker_id")
        )

    except Exception:

        flash(
            "Invalid worker.",
            "danger"
        )

        return redirect(
            url_for("advance")
        )

    month = (
        form.get("month_year")
        or month_name_year(
            form.get("month"),
            form.get("year")
        )
    )

    if not require_payroll_unlocked(month):
        return redirect(url_for("advance", month=month))

    advance_date = (
        form.get("advance_date")
        or datetime.date.today()
        .isoformat()
    )

    amount = parse_num(
        form.get("amount"),
        0
    )

    note = form.get(
        "note",
        ""
    ).strip()

    if amount <= 0:

        flash(
            "Amount must be greater than zero.",
            "danger"
        )

        return redirect(
            url_for(
                "advance",
                month=month
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

        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Advance added: Worker {worker_id}, {month}, BDT {amount:.2f}"
        )
        flash(
            "Advance saved successfully.",
            "success"
        )

    except Exception as e:

        app.logger.exception(
            "Advance save error"
        )

        flash(
            f"Could not save advance: {e}",
            "danger"
        )

    return redirect(
        url_for(
            "advance",
            month=month
        )
    )


# ============================================================
# EDIT ADVANCE
@app.route(
    "/advance/delete/<int:aid>",
    methods=["GET", "POST"]
)
@app.route(
    "/advances/delete/<int:aid>",
    methods=["GET", "POST"]
)
@login_required
def delete_advance(aid):

    try:
        record = next((r for r in advance_rows() if str(r.get("id")) == str(aid)), None)
        if record and not require_payroll_unlocked(record.get("month_year")):
            return redirect(url_for("advance", month=record.get("month_year")))

        delete_advance_record(
            aid
        )

        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Advance deleted: ID {aid}, Worker {record.get('worker_id')}"
        )
        flash(
            "Advance deleted successfully.",
            "success"
        )

    except Exception as e:

        app.logger.exception(
            "Advance delete error"
        )

        flash(
            f"Could not delete advance: {e}",
            "danger"
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
@admin_required
def settings():

    if request.method == "POST":

        for key in [
            "company_name",
            "company_address",
            "company_phone",
            "company_email",
            "company_logo",
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
                    VALUES (?,?)
                    ON CONFLICT(key)
                    DO UPDATE
                    SET value=EXCLUDED.value
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
                    INTO company_settings
                    (key,value)
                    VALUES (?,?)
                    """,
                    (
                        key,
                        value,
                    ),
                    commit=True,
                )

        log_activity(
            current_user().get("username") if current_user() else "system",
            "Company settings updated"
        )
        flash(
            "Settings saved successfully.",
            "success"
        )

    return safe_render_template(
        "settings.html",
        settings=get_settings()
    )



# ============================================================
# WORKER SEARCH
# ============================================================

@app.route("/worker-search")
@login_required
def worker_search():
    q = (request.args.get("q") or "").strip()
    department = (request.args.get("department") or "").strip()
    status = (request.args.get("status") or "").strip()

    sql = """
        SELECT id,name,bangla_name,basic_salary,department,designation,
               phone,status
        FROM workers
        WHERE 1=1
    """
    params = []

    if q:
        sql += """ AND (
            CAST(id AS TEXT) LIKE ?
            OR LOWER(COALESCE(name,'')) LIKE LOWER(?)
            OR LOWER(COALESCE(bangla_name,'')) LIKE LOWER(?)
            OR LOWER(COALESCE(phone,'')) LIKE LOWER(?)
            OR LOWER(COALESCE(designation,'')) LIKE LOWER(?)
            OR LOWER(COALESCE(department,'')) LIKE LOWER(?)
        )"""
        like = f"%{q}%"
        params.extend([like,like,like,like,like,like])

    if department:
        sql += " AND department=?"
        params.append(department)

    if status == "Active":
        sql += " AND COALESCE(NULLIF(TRIM(status), ''), 'Active')=?"
        params.append("Active")
    elif status == "Inactive":
        sql += " AND status=?"
        params.append("Inactive")

    sql += " ORDER BY id"
    results = fetch_all(sql, params)

    departments = fetch_all(
        "SELECT DISTINCT department FROM workers WHERE department IS NOT NULL AND TRIM(department)<>'' ORDER BY department"
    )
    departments = [r["department"] for r in departments]

    return safe_render_template(
        "worker_search.html",
        results=results,
        departments=departments,
        q=q,
        department=department,
        status=status,
        settings=get_settings(),
    )

# ============================================================
# USERS
# ============================================================

@app.route("/users")
@admin_required
def users():

    return safe_render_template(
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
        settings=get_settings(),
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

            return safe_render_template(
                "user_form.html",
                user=form,
                settings=get_settings(),
            )

        password_hash = (
            hash_password(password)
        )

        user_columns = columns(
            "users"
        )

        names = []
        values = []

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

        for name, value in data:

            if name in user_columns:

                names.append(name)
                values.append(value)

        try:

            execute(
                """
                INSERT INTO users
                (""" + ",".join(names) + """)
                VALUES
                (""" + ",".join(
                    ["?"] * len(values)
                ) + """)
                """,
                values,
                commit=True,
            )

            log_activity(
                current_user().get("username") if current_user() else "system",
                f"User created: {username}"
            )
            flash(
                "User created successfully.",
                "success"
            )

            return redirect(
                url_for("users")
            )

        except Exception as e:

            app.logger.exception(
                "User creation error"
            )

            flash(
                f"Could not create user: {e}",
                "danger"
            )

    return safe_render_template(
        "user_form.html",
        user=None,
        settings=get_settings(),
    )


# ============================================================
# EDIT USER
# ============================================================

@app.route(
    "/users/edit/<int:user_id>",
    methods=["GET", "POST"]
)
@admin_required
def edit_user(user_id):

    target = fetch_one("SELECT * FROM users WHERE id=?", (user_id,))
    if not target:
        flash("User not found.", "danger")
        return redirect(url_for("users"))

    if request.method == "POST":
        form = request.form
        username = form.get("username", "").strip()
        full_name = form.get("full_name", "").strip()
        role = form.get("role", "Operator").strip() or "Operator"
        password = form.get("password", "")

        if not username:
            flash("Username is required.", "danger")
            return safe_render_template("user_form.html", user=target, edit_mode=True, settings=get_settings())

        duplicate = fetch_one("SELECT id FROM users WHERE username=? AND id<>?", (username, user_id))
        if duplicate:
            flash("This username already exists.", "danger")
            return safe_render_template("user_form.html", user=target, edit_mode=True, settings=get_settings())

        try:
            if password:
                password_hash = hash_password(password)
                execute("UPDATE users SET username=?, full_name=?, role=?, password_hash=?, password=? WHERE id=?",
                        (username, full_name, role, password_hash, password_hash, user_id), commit=True)
            else:
                execute("UPDATE users SET username=?, full_name=?, role=? WHERE id=?",
                        (username, full_name, role, user_id), commit=True)

            actor = current_user().get("username") if current_user() else "system"
            log_activity(actor, f"User updated: {username}")
            flash("User updated successfully.", "success")
            return redirect(url_for("users"))
        except Exception as e:
            app.logger.exception("User update error")
            flash(f"Could not update user: {e}", "danger")

    return safe_render_template("user_form.html", user=target, edit_mode=True, settings=get_settings())


# ============================================================
# ACTIVATE / DEACTIVATE USER
# ============================================================

@app.route(
    "/users/toggle/<int:user_id>",
    methods=["POST"]
)
@admin_required
def toggle_user(user_id):

    actor = current_user()
    target = fetch_one("SELECT id, username, active FROM users WHERE id=?", (user_id,))
    if not target:
        flash("User not found.", "danger")
        return redirect(url_for("users"))
    if actor and actor.get("id") == user_id:
        flash("You cannot deactivate the logged-in user.", "danger")
        return redirect(url_for("users"))

    new_active = 0 if int(target.get("active") or 0) else 1
    execute("UPDATE users SET active=? WHERE id=?", (new_active, user_id), commit=True)
    actor_name = actor.get("username") if actor else "system"
    log_activity(actor_name, f"User {'activated' if new_active else 'deactivated'}: {target.get('username')}")
    flash(f"User {'activated' if new_active else 'deactivated'} successfully.", "success")
    return redirect(url_for("users"))


# ============================================================
# DELETE USER
# ============================================================

@app.route(
    "/users/delete/<int:user_id>"
)
@admin_required
def delete_user(user_id):

    user = current_user()

    if (
        user
        and user.get("id")
        == user_id
    ):

        flash(
            "You cannot delete the logged-in user.",
            "danger"
        )

    else:

        target_user = fetch_one("SELECT username FROM users WHERE id=?", (user_id,))
        execute(
            """
            DELETE FROM users
            WHERE id=?
            """,
            (user_id,),
            commit=True,
        )
        log_activity(
            user.get("username") if user else "system",
            f"User deleted: {target_user.get('username') if target_user else user_id}"
        )

        flash(
            "User deleted successfully.",
            "success"
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

    username = request.args.get("username", "").strip()
    action = request.args.get("action", "").strip()
    from_date = request.args.get("from_date", "").strip()
    to_date = request.args.get("to_date", "").strip()

    where = []
    params = []

    if username:
        where.append("username LIKE ?")
        params.append("%" + username + "%")
    if action:
        where.append("action LIKE ?")
        params.append("%" + action + "%")
    if from_date:
        where.append("log_time >= ?")
        params.append(from_date + " 00:00:00")
    if to_date:
        where.append("log_time <= ?")
        params.append(to_date + " 23:59:59")

    sql = "SELECT * FROM activity_log"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY id DESC LIMIT 500"

    return safe_render_template(
        "activity.html",
        logs=fetch_all(sql, params),
        username=username,
        action=action,
        from_date=from_date,
        to_date=to_date,
        settings=get_settings(),
    )


@app.route("/activity/export")
@admin_required
def activity_export():

    username = request.args.get("username", "").strip()
    action = request.args.get("action", "").strip()
    from_date = request.args.get("from_date", "").strip()
    to_date = request.args.get("to_date", "").strip()

    where = []
    params = []
    if username:
        where.append("username LIKE ?")
        params.append("%" + username + "%")
    if action:
        where.append("action LIKE ?")
        params.append("%" + action + "%")
    if from_date:
        where.append("log_time >= ?")
        params.append(from_date + " 00:00:00")
    if to_date:
        where.append("log_time <= ?")
        params.append(to_date + " 23:59:59")

    sql = "SELECT id, username, action, log_time FROM activity_log"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY id DESC"

    rows = fetch_all(sql, params)
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["ID", "Username", "Action", "Time"])
    for row in rows:
        writer.writerow([row.get("id", ""), row.get("username", ""), row.get("action", ""), row.get("log_time", "")])
    output.seek(0)

    log_activity(
        current_user().get("username") if current_user() else "system",
        "Activity Log exported"
    )
    return send_file(
        io.BytesIO(output.getvalue().encode("utf-8-sig")),
        mimetype="text/csv",
        as_attachment=True,
        download_name="activity_log.csv",
    )



# ============================================================
# BACKUP & MAINTENANCE
# ============================================================

BACKUP_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "backups",
)


def get_database_backups():
    os.makedirs(BACKUP_DIR, exist_ok=True)

    rows = []

    for name in os.listdir(BACKUP_DIR):
        if not (
            name.startswith("factory_payroll_backup_")
            and name.lower().endswith(".db")
        ):
            continue

        path = os.path.join(BACKUP_DIR, name)

        if not os.path.isfile(path):
            continue

        try:
            stat = os.stat(path)
            rows.append(
                {
                    "name": name,
                    "size": f"{stat.st_size:,} bytes",
                    "created": datetime.datetime.fromtimestamp(
                        stat.st_mtime
                    ).strftime("%Y-%m-%d %H:%M:%S"),
                }
            )
        except Exception:
            pass

    rows.sort(
        key=lambda row: row["name"],
        reverse=True,
    )

    return rows



# ============================================================
# PORTABLE ONLINE <-> OFFLINE DATA BACKUP
# ============================================================

PORTABLE_BACKUP_VERSION = 1


def _portable_table_names():
    """Return application tables from the active SQLite/PostgreSQL database."""
    if is_postgres():
        rows = fetch_all("""
            SELECT table_name
            FROM information_schema.tables
            WHERE table_schema='public'
              AND table_type='BASE TABLE'
            ORDER BY table_name
        """)
        return [
            str(r["table_name"])
            for r in rows
            if not str(r["table_name"]).startswith("pg_")
        ]

    rows = fetch_all("""
        SELECT name
        FROM sqlite_master
        WHERE type='table'
          AND name NOT LIKE 'sqlite_%'
        ORDER BY name
    """)
    return [str(r["name"]) for r in rows]


def _portable_columns(table_name):
    """Return columns in stable database order."""
    if is_postgres():
        rows = fetch_all("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema='public'
              AND table_name=?
            ORDER BY ordinal_position
        """, (table_name,))
        return [str(r["column_name"]) for r in rows]

    conn = db_connect()
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA table_info(" + table_name + ")")
        return [str(row[1]) for row in cur.fetchall()]
    finally:
        try:
            cur.close()
        except Exception:
            pass
        db_release(conn)


def _portable_json_value(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (datetime.datetime, datetime.date, datetime.time)):
        return {"__reedoy_type__": "datetime", "value": value.isoformat()}
    if isinstance(value, Decimal):
        return {"__reedoy_type__": "decimal", "value": str(value)}
    if isinstance(value, (bytes, bytearray, memoryview)):
        return {
            "__reedoy_type__": "bytes",
            "value": base64.b64encode(bytes(value)).decode("ascii"),
        }
    return {"__reedoy_type__": "string", "value": str(value)}


def _portable_restore_value(value):
    if not isinstance(value, dict) or "__reedoy_type__" not in value:
        return value

    kind = value.get("__reedoy_type__")
    raw = value.get("value")

    if kind == "bytes":
        try:
            return base64.b64decode(raw)
        except Exception:
            return raw

    # PostgreSQL accepts ISO date/datetime strings for normal date/timestamp
    # columns, and SQLite stores them as text. Keep the portable value textual.
    if kind in {"datetime", "decimal", "string"}:
        return raw

    return raw


def _portable_snapshot():
    """Build a database-independent snapshot of the active database."""
    tables = []

    for table_name in _portable_table_names():
        cols = _portable_columns(table_name)
        if not cols:
            continue

        rows = fetch_all(
            "SELECT * FROM " + table_name
        )

        data_rows = []
        for row in rows:
            data_rows.append([
                _portable_json_value(row.get(col))
                for col in cols
            ])

        tables.append({
            "name": table_name,
            "columns": cols,
            "rows": data_rows,
        })

    return {
        "format": "REEDOY_PORTABLE_BACKUP",
        "version": PORTABLE_BACKUP_VERSION,
        "created_at": nowstr(),
        "database_type": "postgresql" if is_postgres() else "sqlite",
        "company": get_settings().get("company_name", DEFAULT_COMPANY_NAME),
        "tables": tables,
    }


def _write_portable_backup(path):
    snapshot = _portable_snapshot()

    with zipfile.ZipFile(
        path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
    ) as zf:
        zf.writestr(
            "manifest.json",
            json.dumps(
                snapshot,
                ensure_ascii=False,
                separators=(",", ":"),
            ).encode("utf-8"),
        )


def _read_portable_backup(path):
    with zipfile.ZipFile(path, "r") as zf:
        if "manifest.json" not in zf.namelist():
            raise ValueError("This is not a valid Reedoy portable backup.")

        raw = zf.read("manifest.json")

    snapshot = json.loads(raw.decode("utf-8"))

    if snapshot.get("format") != "REEDOY_PORTABLE_BACKUP":
        raise ValueError("Invalid Reedoy backup format.")

    if int(snapshot.get("version", 0)) != PORTABLE_BACKUP_VERSION:
        raise ValueError("Unsupported Reedoy backup version.")

    if not isinstance(snapshot.get("tables"), list):
        raise ValueError("Backup table data is missing.")

    return snapshot


def _portable_restore_snapshot(snapshot):
    """
    Replace rows in tables that exist in both source and target.
    Existing table structures are never dropped or recreated.
    """
    source_tables = {
        str(t.get("name")): t
        for t in snapshot.get("tables", [])
        if t.get("name")
    }

    target_tables = _portable_table_names()

    # Delete children first. Most app tables have no hard FK constraints,
    # but this order also keeps the common payroll relationships safe.
    preferred = [
        "activity_log",
        "payment_transactions",
        "payroll_payments",
        "payroll_locks",
        "worker_advances",
        "advance_salary",
        "advances",
        "daily_attendance",
        "attendance",
        "workers",
        "users",
        "company_settings",
    ]

    delete_order = (
        [x for x in preferred if x in target_tables]
        + [x for x in target_tables if x not in preferred]
    )

    # Parent tables are restored first.
    insert_order = list(reversed(delete_order))

    conn = db_connect()
    cur = conn.cursor()

    restored = []
    skipped = []

    try:
        # SQLite foreign keys are not normally enabled by this application,
        # but explicitly disabling them for this transaction prevents legacy
        # child-table ordering from blocking a legitimate full restore.
        if not is_postgres():
            try:
                cur.execute("PRAGMA foreign_keys=OFF")
            except Exception:
                pass

        # Clear current rows.
        for table_name in delete_order:
            if table_name not in source_tables:
                continue
            cur.execute(placeholders("DELETE FROM " + table_name))

        # Insert source rows using only columns available in the target.
        for table_name in insert_order:
            source = source_tables.get(table_name)
            if not source:
                continue

            target_cols = _portable_columns(table_name)
            source_cols = [str(c) for c in source.get("columns", [])]
            common_cols = [c for c in source_cols if c in target_cols]

            if not common_cols:
                skipped.append(table_name)
                continue

            source_index = {
                col: idx
                for idx, col in enumerate(source_cols)
            }

            quoted_cols = ",".join(common_cols)
            marks = ",".join(["?"] * len(common_cols))
            sql = (
                "INSERT INTO " + table_name
                + " (" + quoted_cols + ") VALUES (" + marks + ")"
            )

            for row_values in source.get("rows", []):
                values = [
                    _portable_restore_value(
                        row_values[source_index[col]]
                    )
                    for col in common_cols
                ]
                cur.execute(placeholders(sql), values)

            restored.append({
                "table": table_name,
                "rows": len(source.get("rows", [])),
            })

        # Restore PostgreSQL serial/identity sequences to the restored IDs.
        if is_postgres():
            for item in restored:
                table_name = item["table"]
                if "id" not in _portable_columns(table_name):
                    continue

                try:
                    cur.execute(
                        """
                        SELECT pg_get_serial_sequence(?, 'id')
                        """.replace("?", "%s"),
                        (table_name,),
                    )
                    row = cur.fetchone()
                    sequence_name = row[0] if row else None

                    if sequence_name:
                        cur.execute(
                            "SELECT MAX(id) FROM " + table_name
                        )
                        max_row = cur.fetchone()
                        max_id = max_row[0] if max_row else None

                        if max_id is None:
                            cur.execute(
                                "SELECT setval(%s, 1, false)",
                                (sequence_name,),
                            )
                        else:
                            cur.execute(
                                "SELECT setval(%s, %s, true)",
                                (sequence_name, int(max_id)),
                            )
                except Exception:
                    # Sequence reset is a convenience; do not make a valid
                    # data restore fail because a legacy id is not serial.
                    pass

        conn.commit()

        return restored, skipped

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


@app.route("/data-sync")
@admin_required
def data_sync():
    return safe_render_template(
        "data_sync.html",
        settings=get_settings(),
        is_online=is_postgres(),
    )


@app.route("/data-sync/create", methods=["POST"])
@admin_required
def create_portable_backup():
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"reedoy_portable_backup_{timestamp}.rdb"
    path = os.path.join(BACKUP_DIR, filename)

    os.makedirs(BACKUP_DIR, exist_ok=True)

    try:
        _write_portable_backup(path)
        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Portable backup created: {filename}",
        )
        return send_file(
            path,
            as_attachment=True,
            download_name=filename,
            mimetype="application/octet-stream",
        )
    except Exception as exc:
        app.logger.exception("Portable backup creation failed")
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception:
            pass
        flash(f"Portable backup failed: {exc}", "danger")
        return redirect(url_for("data_sync"))


@app.route("/data-sync/restore", methods=["POST"])
@admin_required
def restore_portable_backup():
    upload = request.files.get("backup_file")

    if not upload or not upload.filename:
        flash("Please select a Reedoy portable backup file.", "danger")
        return redirect(url_for("data_sync"))

    filename = os.path.basename(upload.filename)
    if not filename.lower().endswith((".rdb", ".zip")):
        flash("Please select a .rdb portable backup file.", "danger")
        return redirect(url_for("data_sync"))

    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    uploaded_path = os.path.join(
        BACKUP_DIR,
        f"portable_upload_{stamp}.rdb",
    )
    safety_path = os.path.join(
        BACKUP_DIR,
        f"reedoy_safety_before_sync_{stamp}.rdb",
    )

    try:
        upload.save(uploaded_path)

        # Validate the backup before touching the active database.
        snapshot = _read_portable_backup(uploaded_path)

        # Safety copy of the current database in the same portable format.
        _write_portable_backup(safety_path)

        restored, skipped = _portable_restore_snapshot(snapshot)

        log_activity(
            current_user().get("username") if current_user() else "system",
            "Portable backup restored "
            f"from {filename}; tables restored: {len(restored)}",
        )

        total_rows = sum(item["rows"] for item in restored)
        msg = (
            f"Restore completed successfully. "
            f"{len(restored)} tables / {total_rows} rows restored."
        )
        if skipped:
            msg += " Skipped tables: " + ", ".join(skipped) + "."

        flash(msg, "success")

    except Exception as exc:
        app.logger.exception("Portable backup restore failed")
        flash(f"Restore failed. No completed restore was committed: {exc}", "danger")

    finally:
        try:
            if os.path.exists(uploaded_path):
                os.remove(uploaded_path)
        except Exception:
            pass

    return redirect(url_for("data_sync"))


@app.route("/backup")
@admin_required
def backup_maintenance():

    return safe_render_template(
        "backup.html",
        settings=get_settings(),
        db_path=DB_PATH,
        backups=get_database_backups(),
    )


@app.route(
    "/backup/create",
    methods=["POST"]
)
@admin_required
def create_database_backup():

    # This module is intended for the local SQLite source-of-truth
    # database. Do not silently create a misleading SQLite file when
    # the application is connected to PostgreSQL on Render.
    if is_postgres():
        flash(
            "Database backup from this page is available for the local SQLite database. "
            "Use the PostgreSQL backup/export process for the Render database.",
            "danger",
        )
        return redirect(
            url_for("backup_maintenance")
        )

    if not os.path.isfile(DB_PATH):
        flash(
            f"Database file not found: {DB_PATH}",
            "danger",
        )
        return redirect(
            url_for("backup_maintenance")
        )

    os.makedirs(
        BACKUP_DIR,
        exist_ok=True,
    )

    timestamp = datetime.datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    backup_path = os.path.join(
        BACKUP_DIR,
        f"factory_payroll_backup_{timestamp}.db",
    )

    source = None
    destination = None

    try:
        source = sqlite3.connect(
            DB_PATH,
            timeout=30,
        )

        destination = sqlite3.connect(
            backup_path,
            timeout=30,
        )

        with destination:
            source.backup(
                destination
            )

        destination.close()
        destination = None

        source.close()
        source = None

        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Database backup created: {os.path.basename(backup_path)}"
        )
        flash(
            "Database backup created successfully.",
            "success",
        )

    except Exception as exc:
        app.logger.exception(
            "Database backup failed"
        )

        try:
            if destination is not None:
                destination.close()
        except Exception:
            pass

        try:
            if source is not None:
                source.close()
        except Exception:
            pass

        try:
            if os.path.exists(backup_path):
                os.remove(backup_path)
        except Exception:
            pass

        flash(
            f"Backup failed: {exc}",
            "danger",
        )

    return redirect(
        url_for("backup_maintenance")
    )


@app.route(
    "/backup/restore",
    methods=["POST"]
)
@admin_required
def restore_database_backup():

    # Restore is deliberately local-SQLite only. Never modify Render/PostgreSQL
    # from this page.
    if is_postgres():
        flash(
            "Restore from this page is available only for the local SQLite database.",
            "danger",
        )
        return redirect(url_for("backup_maintenance"))

    filename = str(request.form.get("filename", "")).strip()
    safe_name = os.path.basename(filename)

    if (
        not filename
        or safe_name != filename
        or not safe_name.startswith("factory_payroll_backup_")
        or not safe_name.lower().endswith(".db")
    ):
        abort(404)

    backup_path = os.path.join(BACKUP_DIR, safe_name)

    if not os.path.isfile(backup_path):
        flash("Selected backup file was not found.", "danger")
        return redirect(url_for("backup_maintenance"))

    # Step 1: verify the selected backup before touching the live database.
    check_conn = None
    try:
        check_conn = sqlite3.connect(
            backup_path,
            timeout=30,
        )
        integrity = check_conn.execute(
            "PRAGMA integrity_check"
        ).fetchone()
        if not integrity or str(integrity[0]).lower() != "ok":
            flash(
                "Restore stopped: the selected backup failed SQLite integrity check.",
                "danger",
            )
            return redirect(url_for("backup_maintenance"))
    except Exception as exc:
        app.logger.exception("Backup integrity check failed")
        flash(f"Restore stopped: could not verify backup: {exc}", "danger")
        return redirect(url_for("backup_maintenance"))
    finally:
        try:
            if check_conn is not None:
                check_conn.close()
        except Exception:
            pass

    # Step 2: create a safety backup of the CURRENT database before restore.
    os.makedirs(BACKUP_DIR, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    safety_path = os.path.join(
        BACKUP_DIR,
        f"factory_payroll_backup_before_restore_{timestamp}.db",
    )
    temp_restore = os.path.join(
        BACKUP_DIR,
        f"factory_payroll_restore_temp_{timestamp}.db",
    )

    source = None
    safety = None
    restore_source = None
    restore_temp = None

    try:
        if not os.path.isfile(DB_PATH):
            raise FileNotFoundError(f"Current database not found: {DB_PATH}")

        # Safety backup of the current database.
        source = sqlite3.connect(DB_PATH, timeout=30)
        safety = sqlite3.connect(safety_path, timeout=30)
        with safety:
            source.backup(safety)
        safety.close()
        safety = None
        source.close()
        source = None

        # Copy the selected backup into a temporary DB first. This prevents a
        # partially copied file from ever becoming the live database.
        restore_source = sqlite3.connect(backup_path, timeout=30)
        restore_temp = sqlite3.connect(temp_restore, timeout=30)
        with restore_temp:
            restore_source.backup(restore_temp)
        restore_temp.close()
        restore_temp = None
        restore_source.close()
        restore_source = None

        # Final integrity check on the temporary restored DB.
        verify = sqlite3.connect(temp_restore, timeout=30)
        try:
            integrity = verify.execute("PRAGMA integrity_check").fetchone()
            if not integrity or str(integrity[0]).lower() != "ok":
                raise RuntimeError("Temporary restored database failed integrity check")
        finally:
            verify.close()

        # Replace only after all checks succeed.
        os.replace(temp_restore, DB_PATH)
        temp_restore = None

        log_activity(
            current_user().get("username") if current_user() else "system",
            f"Database restored: {safe_name}"
        )
        flash(
            "Database restored successfully. A safety backup of the previous database was also created.",
            "success",
        )

    except Exception as exc:
        app.logger.exception("Database restore failed")
        flash(
            f"Restore failed. The original database was not replaced. Error: {exc}",
            "danger",
        )

    finally:
        for conn in (restore_temp, restore_source, safety, source):
            try:
                if conn is not None:
                    conn.close()
            except Exception:
                pass

        try:
            if os.path.exists(temp_restore):
                os.remove(temp_restore)
        except Exception:
            pass

    return redirect(url_for("backup_maintenance"))


@app.route(
    "/backup/download/<path:filename>"
)
@admin_required
def download_database_backup(filename):

    safe_name = os.path.basename(
        filename
    )

    if safe_name != filename:
        abort(404)

    if not (
        safe_name.startswith("factory_payroll_backup_")
        and safe_name.lower().endswith(".db")
    ):
        abort(404)

    backup_path = os.path.join(
        BACKUP_DIR,
        safe_name,
    )

    if not os.path.isfile(backup_path):
        abort(404)

    return send_file(
        backup_path,
        as_attachment=True,
        download_name=safe_name,
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
                "openpyxl is not installed."
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
            "application/"
            "vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# ============================================================
# WORKERS EXCEL
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
        "reedoy_workers.xlsx"
    )


# ============================================================
# ADVANCES EXCEL
# ============================================================

@app.route(
    "/export/advances.xlsx"
)
@login_required
def export_advances():

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    return export_rows(
        advance_rows(month),
        "reedoy_advances.xlsx"
    )


# ============================================================
# DEPARTMENT EXCEL
# ============================================================

@app.route(
    "/export/department.xlsx"
)
@login_required
def export_department():

    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

    department_name = request.args.get(
        "department",
        ""
    ).strip()

    if department_name:

        worker_list = fetch_all(
            """
            SELECT *
            FROM workers
            WHERE department=?
            ORDER BY id
            """,
            (department_name,),
        )

    else:

        worker_list = fetch_all(
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    salary_map = (
        calculate_salary_bulk(
            worker_list,
            month
        )
    )

    rows = []

    for worker in worker_list:

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
            "Salary Basis": "30 Days",
            "Weekly Holiday": "Friday",
            "Present": salary.get(
                "present",
                0
            ),
            "Absent": salary.get(
                "absent",
                0
            ),
            "OT": salary.get(
                "ot",
                0
            ),
            "Absent Deduction": salary.get(
                "absent_cut",
                0
            ),
            "OT Amount": salary.get(
                "ot_amt",
                0
            ),
            "Nasta": salary.get(
                "nasta",
                0
            ),
            "Gross": salary.get(
                "gross",
                0
            ),
            "Advance": salary.get(
                "advance",
                0
            ),
            "Net Payable": salary.get(
                "net",
                0
            ),
        })

    return export_rows(
        rows,
        "reedoy_department_salary.xlsx"
    )


# ============================================================
# PAYSLIP PDF
# ============================================================

@app.route(
    "/export/payslip.pdf"
)
@login_required
def export_payslip_pdf():

    if (
        SimpleDocTemplate is None
        or Paragraph is None
        or Table is None
        or TableStyle is None
    ):

        abort(
            503,
            description=(
                "reportlab is not installed."
            )
        )

    worker_id = request.args.get(
        "worker_id"
    )

    # Keep PDF month consistent with the selected payslip period.
    month = month_name_year(
        request.args.get("month"),
        request.args.get("year"),
    )

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

    styles = getSampleStyleSheet()

    story = [
        Paragraph(
            DEFAULT_COMPANY_NAME,
            styles["Title"]
        ),
        Paragraph(
            "Payroll Payslip - " + month,
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
            "English Name",
            worker["name"]
        ],
        [
            "Bangla Name",
            worker.get("bangla_name") or ""
        ],
        [
            "Department",
            worker.get(
                "department"
            ) or ""
        ],
        [
            "Basic Salary (30-Day Basis)",
            f"BDT {salary['earned_basic']:,.2f}"
        ],
        [
            "Present / Absent",
            f"{salary.get('present', 0)} / {salary.get('absent', 0)} days"
        ],
        [
            "OT Hours",
            f"{salary.get('ot', 0)} hours"
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
                ),
            ]),
        )
    )

    document.build(
        story
    )

    output.seek(0)

    filename = (
        f"payslip_"
        f"{worker['id']}_"
        f"{month.replace(' ', '_')}.pdf"
    )

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
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

    else:

        output.write(
            "No records\n"
        )

    data = io.BytesIO(
        output.getvalue().encode(
            "utf-8-sig"
        )
    )

    return send_file(
        data,
        as_attachment=True,
        download_name=(
            "reedoy_workers.csv"
        ),
        mimetype="text/csv",
    )


# ============================================================
# 404
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


# ============================================================
# 500
# ============================================================

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
            f"500 - Internal Server Error\n"
            f"{error}",
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
# LOCAL DEVELOPMENT
# ============================================================



# ============================================================
# TEMPORARY READ-ONLY ONLINE WORKERS DIAGNOSTIC
# ============================================================

@app.route("/diagnostic/online-workers")
@login_required
def diagnostic_online_workers():
    if not is_postgres():
        return """
        <h2>Online Workers Diagnostic</h2>
        <p>DATABASE_URL is not configured. This app is currently using SQLite.</p>
        """, 400

    try:
        conn = db_connect()
        cur = conn.cursor()

        cur.execute("""
            SELECT column_name, data_type
            FROM information_schema.columns
            WHERE table_schema='public'
              AND table_name='workers'
            ORDER BY ordinal_position
        """)

        columns = cur.fetchall()

        cur.execute("SELECT COUNT(*) FROM workers")
        worker_count = cur.fetchone()[0]

        cur.execute("""
            SELECT id, name
            FROM workers
            ORDER BY id
            LIMIT 10
        """)

        samples = cur.fetchall()

        db_release(conn)

        html = """
        <!doctype html>
        <html>
        <head>
            <meta charset="utf-8">
            <title>Online Workers Diagnostic</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 30px;
                    line-height: 1.6;
                }
                table {
                    border-collapse: collapse;
                    margin-top: 10px;
                }
                th, td {
                    border: 1px solid #ccc;
                    padding: 7px 12px;
                    text-align: left;
                }
                th {
                    background: #eee;
                }
                .ok {
                    color: green;
                    font-weight: bold;
                }
            </style>
        </head>
        <body>
            <h2>Reedoy v49 - Online Workers Diagnostic</h2>
            <p class="ok">PostgreSQL connection: OK</p>
            <p><b>Online Worker Count:</b> {{ worker_count }}</p>

            <h3>Workers Table Columns</h3>
            <table>
                <tr>
                    <th>Column</th>
                    <th>Data Type</th>
                </tr>
                {% for row in columns %}
                <tr>
                    <td>{{ row[0] }}</td>
                    <td>{{ row[1] }}</td>
                </tr>
                {% endfor %}
            </table>

            <h3>First 10 Workers</h3>
            <table>
                <tr>
                    <th>ID</th>
                    <th>Name</th>
                </tr>
                {% for row in samples %}
                <tr>
                    <td>{{ row[0] }}</td>
                    <td>{{ row[1] }}</td>
                </tr>
                {% endfor %}
            </table>

            <p><b>READ-ONLY CHECK — NO DATA WAS MODIFIED.</b></p>
        </body>
        </html>
        """

        return render_template_string(
            html,
            worker_count=worker_count,
            columns=columns,
            samples=samples,
        )

    except Exception as e:
        try:
            db_release(conn)
        except Exception:
            pass

        return f"""
        <h2>Online Workers Diagnostic</h2>
        <p><b>ERROR:</b> {type(e).__name__}: {e}</p>
        """, 500




# ============================================================
# TEMPORARY ONLINE WORKERS UUID MIGRATION
# ============================================================

@app.route("/sync/prepare-online-workers")
@login_required
def prepare_online_workers_sync():
    import json
    import hashlib
    from pathlib import Path

    if not is_postgres():
        return """
        <h2>Online Worker Sync Preparation</h2>
        <p>ERROR: This route must run against PostgreSQL.</p>
        """, 400

    seed_path = Path(__file__).resolve().parent / "workers_sync_seed.json"

    if not seed_path.exists():
        return """
        <h2>Online Worker Sync Preparation</h2>
        <p>ERROR: workers_sync_seed.json not found.</p>
        """, 500

    conn = None

    try:
        with open(seed_path, "r", encoding="utf-8") as f:
            seed = json.load(f)

        conn = db_connect()
        cur = conn.cursor()

        # ----------------------------------------------------
        # Read current Online workers
        # ----------------------------------------------------
        cur.execute("""
            SELECT id, name
            FROM workers
            ORDER BY id
        """)

        online_rows = cur.fetchall()

        if len(online_rows) != len(seed):
            raise RuntimeError(
                f"Worker count mismatch. "
                f"Online={len(online_rows)}, Seed={len(seed)}. "
                f"NO CHANGES WERE MADE."
            )

        seed_map = {
            int(row["id"]): row
            for row in seed
        }

        mismatches = []

        for row in online_rows:
            worker_id = int(row[0])
            online_name = row[1] or ""

            if worker_id not in seed_map:
                mismatches.append(
                    f"ID {worker_id}: not found in local seed"
                )
                continue

            identity = f"{worker_id}|{online_name}"

            online_hash = hashlib.sha256(
                identity.encode("utf-8")
            ).hexdigest()

            if online_hash != seed_map[worker_id]["identity_hash"]:
                mismatches.append(
                    f"ID {worker_id}: name mismatch"
                )

        if mismatches:
            conn.rollback()

            html = """
            <h2>Online Worker Sync Preparation</h2>
            <h3 style="color:red">
                STOP - NO DATABASE CHANGES WERE MADE
            </h3>
            <p>
                The Online workers do not exactly match the Local
                workers.
            </p>
            <p>Mismatches: {{ count }}</p>
            <ul>
            {% for item in mismatches %}
                <li>{{ item }}</li>
            {% endfor %}
            </ul>
            """

            return render_template_string(
                html,
                count=len(mismatches),
                mismatches=mismatches
            ), 409

        # ----------------------------------------------------
        # All workers match
        # ----------------------------------------------------
        cur.execute("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema='public'
              AND table_name='workers'
              AND column_name='sync_uuid'
        """)

        has_sync_uuid = cur.fetchone() is not None

        if not has_sync_uuid:
            cur.execute("""
                ALTER TABLE workers
                ADD COLUMN sync_uuid TEXT
            """)

        # ----------------------------------------------------
        # Make sure existing UUIDs are not conflicting
        # ----------------------------------------------------
        cur.execute("""
            SELECT id, sync_uuid
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
        """)

        existing_uuid_rows = cur.fetchall()

        existing_uuid_map = {
            int(row[0]): str(row[1])
            for row in existing_uuid_rows
        }

        conflicts = []

        for worker_id, local_data in seed_map.items():
            existing = existing_uuid_map.get(worker_id)

            if existing and existing != local_data["sync_uuid"]:
                conflicts.append(
                    f"ID {worker_id}: existing sync_uuid differs"
                )

        if conflicts:
            conn.rollback()

            html = """
            <h2>Online Worker Sync Preparation</h2>
            <h3 style="color:red">
                STOP - UUID CONFLICT
            </h3>
            <p>NO DATABASE CHANGES WERE MADE.</p>
            <ul>
            {% for item in conflicts %}
                <li>{{ item }}</li>
            {% endfor %}
            </ul>
            """

            return render_template_string(
                html,
                conflicts=conflicts
            ), 409

        # ----------------------------------------------------
        # Assign UUIDs
        # ----------------------------------------------------
        updated = 0

        for worker_id, local_data in seed_map.items():

            cur.execute("""
                UPDATE workers
                SET sync_uuid=%s
                WHERE id=%s
                  AND (
                      sync_uuid IS NULL
                      OR TRIM(sync_uuid)=''
                  )
            """, (
                local_data["sync_uuid"],
                worker_id
            ))

            updated += cur.rowcount

        # ----------------------------------------------------
        # Unique index
        # ----------------------------------------------------
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS
            idx_workers_sync_uuid
            ON workers(sync_uuid)
        """)

        # ----------------------------------------------------
        # Sync metadata table
        # ----------------------------------------------------
        cur.execute("""
            CREATE TABLE IF NOT EXISTS reedoy_sync_meta (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT
            )
        """)

        # ----------------------------------------------------
        # Workers baseline table
        # ----------------------------------------------------
        cur.execute("""
            CREATE TABLE IF NOT EXISTS reedoy_sync_worker_base (
                sync_uuid TEXT PRIMARY KEY,
                worker_id INTEGER,
                fingerprint TEXT,
                snapshot_json TEXT,
                last_synced_at TEXT
            )
        """)

        conn.commit()

        # ----------------------------------------------------
        # Final verification
        # ----------------------------------------------------
        cur.execute("""
            SELECT COUNT(*)
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
        """)

        uuid_count = int(cur.fetchone()[0])

        cur.execute("SELECT COUNT(*) FROM workers")
        worker_count = int(cur.fetchone()[0])

        db_release(conn)
        conn = None

        html = """
        <h2>Online Worker Sync Preparation</h2>

        <h3 style="color:green">
            SUCCESS
        </h3>

        <table border="1" cellpadding="8">
            <tr>
                <th>Item</th>
                <th>Result</th>
            </tr>
            <tr>
                <td>Online Workers</td>
                <td>{{ worker_count }}</td>
            </tr>
            <tr>
                <td>Workers with sync_uuid</td>
                <td>{{ uuid_count }}</td>
            </tr>
            <tr>
                <td>New UUIDs assigned</td>
                <td>{{ updated }}</td>
            </tr>
        </table>

        <p>
        Existing worker IDs, names, salary, OT, department,
        attendance, advance and payment data were not modified.
        </p>

        <p>
        <b>Workers two-way sync foundation is now ready.</b>
        </p>
        """

        return render_template_string(
            html,
            worker_count=worker_count,
            uuid_count=uuid_count,
            updated=updated
        )

    except Exception as e:

        if conn is not None:
            try:
                conn.rollback()
            except Exception:
                pass

            try:
                db_release(conn)
            except Exception:
                pass

        return f"""
        <h2>Online Worker Sync Preparation</h2>
        <h3 style="color:red">ERROR - NO CHANGES WERE COMMITTED</h3>
        <p>{type(e).__name__}: {e}</p>
        """, 500





@app.route("/diagnostic/online-worker-337")
@login_required
def diagnostic_online_worker_337():
    if not is_postgres():
        return "ERROR: PostgreSQL connection is not active.", 400

    conn = None

    try:
        conn = db_connect()
        cur = conn.cursor()

        cur.execute("""
            SELECT id, name, bangla_name, department,
                   designation, basic_salary, ot_rate,
                   refreshment_bill, status
            FROM workers
            WHERE id = %s
        """, (337,))

        raw = cur.fetchone()

        if raw is None:
            db_release(conn)
            conn = None
            return "<h2>Online Worker ID 337 NOT FOUND</h2>", 404

        # Read values by column position from the actual PostgreSQL cursor.
        values = list(raw)

        fields = [
            "ID",
            "English Name",
            "Bangla Name",
            "Department",
            "Designation",
            "Basic Salary",
            "OT Rate",
            "Refreshment Bill",
            "Status"
        ]

        data = []

        for i, field in enumerate(fields):
            value = values[i] if i < len(values) else None

            if value is None:
                value = ""

            data.append((field, value))

        db_release(conn)
        conn = None

        return render_template_string("""
        <h2>Reedoy v49 - Online Worker ID 337</h2>

        <table border="1" cellpadding="8">
            <tr>
                <th>Field</th>
                <th>Value</th>
            </tr>

            {% for field, value in data %}
            <tr>
                <td><b>{{ field }}</b></td>
                <td>{{ value }}</td>
            </tr>
            {% endfor %}
        </table>

        <p>
            <b>READ ONLY — no database changes were made.</b>
        </p>
        """, data=data)

    except Exception as e:

        if conn is not None:
            try:
                db_release(conn)
            except Exception:
                pass

        return f"""
        <h2>Worker 337 Diagnostic Error</h2>
        <p>{type(e).__name__}: {e}</p>
        <p><b>READ ONLY — no database changes were made.</b></p>
        """, 500


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






