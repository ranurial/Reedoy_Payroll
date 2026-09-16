import os
import csv
import io
import sqlite3
import hashlib

from datetime import datetime, date
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
    abort
)


# =========================================================
# APP CONFIG
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "CHANGE-ME-IN-RENDER"
)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

DB_PATH = os.environ.get(
    "SQLITE_PATH",
    os.path.join(os.path.dirname(__file__), "reedoy_payroll.db")
)


# =========================================================
# BANGLA TRANSLATION
# =========================================================

BN = {
    "Dashboard": "ড্যাশবোর্ড",
    "Workers": "কর্মী",
    "Attendance": "উপস্থিতি",
    "Advances": "অগ্রিম বেতন",
    "Payroll Report": "বেতন রিপোর্ট",
    "Department Salary": "বিভাগভিত্তিক বেতন",
    "Users": "ইউজার",
    "Activity Log": "কার্যক্রম লগ",
    "Company Settings": "কোম্পানি সেটিংস",
    "Logout": "লগআউট",
    "Login": "লগইন",
    "Username": "ইউজার আইডি",
    "Password": "পাসওয়ার্ড",
    "Name": "নাম",
    "Bangla Name": "বাংলা নাম",
    "Department": "বিভাগ",
    "Designation": "পদবী",
    "Basic Salary": "মূল বেতন",
    "OT Rate/Hour": "ওটি হার/ঘণ্টা",
    "Refreshment": "নাস্তা",
    "Save": "সংরক্ষণ",
    "Search": "খুঁজুন",
    "Month": "মাস",
    "Present": "উপস্থিত",
    "Absent": "অনুপস্থিত",
    "OT Hours": "ওটি ঘণ্টা",
    "Advance": "অগ্রিম",
    "Net Payable": "নেট প্রদেয়",
    "Gross Salary": "মোট বেতন",
    "Add Worker": "কর্মী যোগ করুন",
    "Edit": "সম্পাদনা",
    "Delete": "মুছুন",
    "Generate": "তৈরি করুন",
    "Export CSV": "CSV ডাউনলোড",
    "Export Excel": "Excel ডাউনলোড",
    "Export PDF": "PDF ডাউনলোড",
    "Add Advance": "অগ্রিম যোগ করুন",
    "Date": "তারিখ",
    "Amount": "পরিমাণ",
    "Note": "নোট",
    "Action": "কার্যক্রম",
    "Total": "মোট",
    "Worker ID": "কর্মী আইডি",
    "Company Name": "কোম্পানির নাম",
    "Address": "ঠিকানা",
    "Phone": "ফোন",
    "Email": "ইমেইল",
    "Save Settings": "সেটিংস সংরক্ষণ",
    "Active": "সক্রিয়",
    "Role": "রোল",
    "No data": "কোনো তথ্য নেই",
    "Current Language": "বর্তমান ভাষা",
    "English": "ইংরেজি",
    "Bangla": "বাংলা",
    "Salary Report": "বেতন রিপোর্ট",
    "Payslip": "পে-স্লিপ",
    "New Password": "নতুন পাসওয়ার্ড",
    "Full Name": "পূর্ণ নাম",
    "Create User": "ইউজার তৈরি করুন",
    "User Management": "ইউজার ম্যানেজমেন্ট",
    "Status": "স্ট্যাটাস",
    "Department Filter": "বিভাগ নির্বাচন"
}


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


def tr(value):
    if session.get("lang", "en") == "bn":
        return BN.get(value, value)
    return value


def dept(value):
    if session.get("lang", "en") == "bn":
        return DEPT_BN.get(value, value)
    return value


app.jinja_env.globals.update(
    tr=tr,
    dept=dept
)


# =========================================================
# PASSWORD HASH
# =========================================================

def sha(value):
    return hashlib.sha256(
        value.encode("utf-8")
    ).hexdigest()


# =========================================================
# DATABASE CLASS
# =========================================================

class DB:

    def __init__(self):
        self.pg = bool(DATABASE_URL)

    def connect(self):

        if self.pg:

            import psycopg
            from psycopg.rows import dict_row

            db_url = DATABASE_URL.replace(
                "postgres://",
                "postgresql://",
                1
            )

            return psycopg.connect(
                db_url,
                row_factory=dict_row
            )

        connection = sqlite3.connect(DB_PATH)
        connection.row_factory = sqlite3.Row

        return connection

    def execute(
        self,
        sql,
        params=(),
        fetch=False,
        many=False
    ):

        connection = self.connect()
        cursor = connection.cursor()

        if self.pg:
            sql = sql.replace("?", "%s")

        if many:
            cursor.executemany(sql, params)
        else:
            cursor.execute(sql, params)

        result = cursor.fetchall() if fetch else None

        connection.commit()

        cursor.close()
        connection.close()

        return result


db = DB()


def rows(sql, params=()):
    return db.execute(
        sql,
        params,
        fetch=True
    )


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

def init_db():

    if db.pg:

        statements = [

            """
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT,
                role TEXT NOT NULL DEFAULT 'User',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                last_login TEXT
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS company_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS workers (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                bangla_name TEXT,
                department TEXT NOT NULL,
                designation TEXT,
                basic_salary DOUBLE PRECISION NOT NULL DEFAULT 0,
                ot_rate DOUBLE PRECISION NOT NULL DEFAULT 0,
                refreshment_bill DOUBLE PRECISION NOT NULL DEFAULT 0
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS attendance (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER NOT NULL,
                month_year TEXT NOT NULL,
                present_days INTEGER DEFAULT 0,
                absent_days INTEGER DEFAULT 0,
                ot_hours DOUBLE PRECISION DEFAULT 0,
                UNIQUE(worker_id, month_year)
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER NOT NULL,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount DOUBLE PRECISION NOT NULL,
                note TEXT
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS activity_log (
                id SERIAL PRIMARY KEY,
                username TEXT,
                action TEXT,
                log_time TEXT NOT NULL
            )
            """

        ]

    else:

        statements = [

            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT,
                role TEXT NOT NULL DEFAULT "User",
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                last_login TEXT
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS company_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS workers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                bangla_name TEXT,
                department TEXT NOT NULL,
                designation TEXT,
                basic_salary REAL NOT NULL DEFAULT 0,
                ot_rate REAL NOT NULL DEFAULT 0,
                refreshment_bill REAL NOT NULL DEFAULT 0
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER NOT NULL,
                month_year TEXT NOT NULL,
                present_days INTEGER DEFAULT 0,
                absent_days INTEGER DEFAULT 0,
                ot_hours REAL DEFAULT 0,
                UNIQUE(worker_id, month_year)
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER NOT NULL,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount REAL NOT NULL,
                note TEXT
            )
            """,

            """
            CREATE TABLE IF NOT EXISTS activity_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                action TEXT,
                log_time TEXT NOT NULL
            )
            """

        ]

    for statement in statements:
        db.execute(statement)


    # -----------------------------------------------------
    # Older database compatibility
    # -----------------------------------------------------

    migrations = [

        "ALTER TABLE workers ADD COLUMN bangla_name TEXT",

        "ALTER TABLE users ADD COLUMN password_hash TEXT",

        "ALTER TABLE users ADD COLUMN full_name TEXT",

        "ALTER TABLE users ADD COLUMN role TEXT",

        "ALTER TABLE users ADD COLUMN active INTEGER",

        "ALTER TABLE users ADD COLUMN created_at TEXT",

        "ALTER TABLE users ADD COLUMN last_login TEXT"

    ]

    for statement in migrations:

        try:
            db.execute(statement)
        except Exception:
            pass


    # -----------------------------------------------------
    # Default admin
    # -----------------------------------------------------

    existing_admin = rows(
        "SELECT id FROM users WHERE username=?",
        ("admin",)
    )

    if not existing_admin:

        db.execute(
            """
            INSERT INTO users
            (
                username,
                password_hash,
                full_name,
                role,
                active,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                "admin",
                sha("admin123"),
                "System Administrator",
                "Administrator",
                1,
                datetime.now().isoformat()
            )
        )


    # -----------------------------------------------------
    # Company settings
    # -----------------------------------------------------

    defaults = {
        "company_name":
            "Reedoy Textile Dyeing Printing & Finishing",

        "company_address":
            "",

        "company_phone":
            "",

        "company_email":
            ""
    }

    for key, value in defaults.items():

        existing = rows(
            "SELECT key FROM company_settings WHERE key=?",
            (key,)
        )

        if not existing:

            db.execute(
                """
                INSERT INTO company_settings(key,value)
                VALUES(?,?)
                """,
                (key, value)
            )


# =========================================================
# SETTINGS
# =========================================================

def setting(key):

    result = rows(
        "SELECT value FROM company_settings WHERE key=?",
        (key,)
    )

    if result:
        return result[0]["value"]

    return ""


# =========================================================
# ACTIVITY LOG
# =========================================================

def log(action):

    if session.get("user"):

        db.execute(
            """
            INSERT INTO activity_log
            (username, action, log_time)
            VALUES (?, ?, ?)
            """,
            (
                session["user"],
                action,
                datetime.now().isoformat()
            )
        )


# =========================================================
# LOGIN DECORATORS
# =========================================================

def login_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        if not session.get("user"):
            return redirect(url_for("login"))

        return function(*args, **kwargs)

    return wrapper


def admin_required(function):

    @wraps(function)
    def wrapper(*args, **kwargs):

        if not session.get("user"):
            return redirect(url_for("login"))

        if session.get("role") != "Administrator":
            abort(403)

        return function(*args, **kwargs)

    return wrapper


# =========================================================
# SALARY CALCULATION
# =========================================================

def calculate_salary(worker):

    basic = float(
        worker.get("basic_salary", 0) or 0
    )

    present = float(
        worker.get("present_days", 0) or 0
    )

    absent = float(
        worker.get("absent_days", 0) or 0
    )

    ot_hours = float(
        worker.get("ot_hours", 0) or 0
    )

    ot_rate = float(
        worker.get("ot_rate", 0) or 0
    )

    refreshment = float(
        worker.get("refreshment_bill", 0) or 0
    )

    advance = float(
        worker.get("advance", 0) or 0
    )

    daily_basic = basic / 30.0

    earned_basic = daily_basic * present

    ot_amount = ot_hours * ot_rate

    gross = (
        earned_basic
        + ot_amount
        + refreshment
    )

    net = gross - advance

    return {
        "earned_basic": round(
            earned_basic,
            2
        ),

        "ot_amount": round(
            ot_amount,
            2
        ),

        "refreshment": round(
            refreshment,
            2
        ),

        "gross": round(
            gross,
            2
        ),

        "advance": round(
            advance,
            2
        ),

        "net": round(
            net,
            2
        )
    }


# =========================================================
# PAYROLL ROWS
# =========================================================

def payroll_rows(
    month,
    search="",
    department="All Departments"
):

    sql = """
        SELECT
            w.*,

            COALESCE(
                a.present_days,
                0
            ) AS present_days,

            COALESCE(
                a.absent_days,
                0
            ) AS absent_days,

            COALESCE(
                a.ot_hours,
                0
            ) AS ot_hours,

            COALESCE(
                (
                    SELECT SUM(x.amount)
                    FROM worker_advances x
                    WHERE x.worker_id = w.id
                    AND x.month_year = ?
                ),
                0
            ) AS advance

        FROM workers w

        LEFT JOIN attendance a
        ON a.worker_id = w.id
        AND a.month_year = ?

        WHERE
            (
                w.name LIKE ?
                OR COALESCE(w.bangla_name,'') LIKE ?
            )

        AND
            (
                ? = 'All Departments'
                OR w.department = ?
            )

        ORDER BY
            w.department,
            w.id
    """

    result = []

    search_pattern = "%" + search + "%"

    data = rows(
        sql,
        (
            month,
            month,
            search_pattern,
            search_pattern,
            department,
            department
        )
    )

    for row in data:

        worker = dict(row)

        salary = calculate_salary(worker)

        # -------------------------------------------------
        # IMPORTANT:
        # report.html needs ot_amount
        # -------------------------------------------------

        worker["earned_basic"] = salary["earned_basic"]

        worker["ot_amount"] = salary["ot_amount"]

        worker["refreshment"] = salary["refreshment"]

        worker["gross"] = salary["gross"]

        worker["advance"] = salary["advance"]

        worker["net"] = salary["net"]

        worker["net_salary"] = salary["net"]

        result.append(worker)

    return result


# =========================================================
# COMMON TEMPLATE DATA
# =========================================================

@app.context_processor
def common():

    return {
        "company":
            setting("company_name")
            or "Reedoy Textile",

        "today":
            date.today().isoformat(),

        "departments":
            list(DEPT_BN.keys())[1:]
    }


# =========================================================
# LANGUAGE
# =========================================================

@app.route("/set-language/<lang>")
def set_language(lang):

    if lang == "bn":
        session["lang"] = "bn"
    else:
        session["lang"] = "en"

    return redirect(
        request.referrer
        or url_for("dashboard")
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

        result = rows(
            """
            SELECT *
            FROM users
            WHERE username=?
            AND active=1
            """,
            (username,)
        )

        if (
            result
            and result[0]["password_hash"]
            == sha(password)
        ):

            session.update(
                user=username,
                role=result[0]["role"],
                lang=session.get("lang", "en")
            )

            db.execute(
                """
                UPDATE users
                SET last_login=?
                WHERE username=?
                """,
                (
                    datetime.now().isoformat(),
                    username
                )
            )

            log("Login")

            return redirect(
                url_for("dashboard")
            )

        flash(
            "Invalid username or password",
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

    session.clear()

    return redirect(
        url_for("login")
    )


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/")
@login_required
def dashboard():

    worker_count = rows(
        "SELECT COUNT(*) c FROM workers"
    )[0]["c"]

    advance_total = rows(
        """
        SELECT COALESCE(SUM(amount),0) s
        FROM worker_advances
        """
    )[0]["s"]

    department_count = rows(
        """
        SELECT COUNT(DISTINCT department) c
        FROM workers
        """
    )[0]["c"]

    return render_template(
        "dashboard.html",

        workers=worker_count,

        advances=advance_total,

        departments=department_count,

        month=datetime.now().strftime(
            "%Y-%m"
        )
    )


# =========================================================
# WORKERS
# =========================================================

@app.route(
    "/workers",
    methods=["GET", "POST"]
)
@login_required
def workers():

    if request.method == "POST":

        form = request.form

        worker_id = form.get("id")

        values = (
            form["name"].strip(),

            form.get(
                "bangla_name",
                ""
            ).strip(),

            form["department"],

            form.get(
                "designation",
                ""
            ).strip(),

            float(
                form.get(
                    "basic_salary"
                ) or 0
            ),

            float(
                form.get(
                    "ot_rate"
                ) or 0
            ),

            float(
                form.get(
                    "refreshment_bill"
                ) or 0
            )
        )

        if worker_id:

            db.execute(
                """
                UPDATE workers
                SET
                    name=?,
                    bangla_name=?,
                    department=?,
                    designation=?,
                    basic_salary=?,
                    ot_rate=?,
                    refreshment_bill=?
                WHERE id=?
                """,
                values + (int(worker_id),)
            )

            log(
                "Updated worker: "
                + form["name"]
            )

            flash(
                "Worker updated",
                "success"
            )

        else:

            db.execute(
                """
                INSERT INTO workers
                (
                    name,
                    bangla_name,
                    department,
                    designation,
                    basic_salary,
                    ot_rate,
                    refreshment_bill
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                values
            )

            log(
                "Added worker: "
                + form["name"]
            )

            flash(
                "Worker saved",
                "success"
            )

        return redirect(
            url_for("workers")
        )


    query = request.args.get(
        "q",
        ""
    ).strip()

    edit_id = request.args.get(
        "edit"
    )

    edit = None

    if edit_id:

        result = rows(
            """
            SELECT *
            FROM workers
            WHERE id=?
            """,
            (edit_id,)
        )

        if result:
            edit = result[0]


    data = rows(
        """
        SELECT *
        FROM workers
        WHERE
            name LIKE ?
            OR COALESCE(bangla_name,'') LIKE ?
        ORDER BY id DESC
        """,
        (
            "%" + query + "%",
            "%" + query + "%"
        )
    )


    return render_template(
        "workers.html",
        workers=data,
        q=query,
        edit=edit
    )


# =========================================================
# DELETE WORKER
# =========================================================

@app.route(
    "/workers/delete/<int:wid>"
)
@login_required
def delete_worker(wid):

    db.execute(
        """
        DELETE FROM attendance
        WHERE worker_id=?
        """,
        (wid,)
    )

    db.execute(
        """
        DELETE FROM worker_advances
        WHERE worker_id=?
        """,
        (wid,)
    )

    db.execute(
        """
        DELETE FROM workers
        WHERE id=?
        """,
        (wid,)
    )

    log(
        f"Deleted worker {wid}"
    )

    flash(
        "Worker deleted",
        "success"
    )

    return redirect(
        url_for("workers")
    )


# =========================================================
# ATTENDANCE
# =========================================================

@app.route(
    "/attendance",
    methods=["GET", "POST"]
)
@login_required
def attendance():

    month = (
        request.values.get("month")
        or datetime.now().strftime("%Y-%m")
    )


    if request.method == "POST":

        worker_list = rows(
            "SELECT id FROM workers"
        )

        for worker in worker_list:

            worker_id = str(
                worker["id"]
            )

            present = int(
                request.form.get(
                    "present_" + worker_id,
                    0
                ) or 0
            )

            absent = int(
                request.form.get(
                    "absent_" + worker_id,
                    0
                ) or 0
            )

            ot = float(
                request.form.get(
                    "ot_" + worker_id,
                    0
                ) or 0
            )


            if db.pg:

                db.execute(
                    """
                    INSERT INTO attendance
                    (
                        worker_id,
                        month_year,
                        present_days,
                        absent_days,
                        ot_hours
                    )
                    VALUES (?, ?, ?, ?, ?)

                    ON CONFLICT
                    (
                        worker_id,
                        month_year
                    )

                    DO UPDATE SET
                        present_days =
                            EXCLUDED.present_days,

                        absent_days =
                            EXCLUDED.absent_days,

                        ot_hours =
                            EXCLUDED.ot_hours
                    """,
                    (
                        worker["id"],
                        month,
                        present,
                        absent,
                        ot
                    )
                )

            else:

                db.execute(
                    """
                    INSERT INTO attendance
                    (
                        worker_id,
                        month_year,
                        present_days,
                        absent_days,
                        ot_hours
                    )
                    VALUES (?, ?, ?, ?, ?)

                    ON CONFLICT
                    (
                        worker_id,
                        month_year
                    )

                    DO UPDATE SET
                        present_days =
                            excluded.present_days,

                        absent_days =
                            excluded.absent_days,

                        ot_hours =
                            excluded.ot_hours
                    """,
                    (
                        worker["id"],
                        month,
                        present,
                        absent,
                        ot
                    )
                )


        log(
            "Updated attendance "
            + month
        )

        flash(
            "Attendance saved",
            "success"
        )


    data = rows(
        """
        SELECT
            w.*,

            COALESCE(
                a.present_days,
                0
            ) present_days,

            COALESCE(
                a.absent_days,
                0
            ) absent_days,

            COALESCE(
                a.ot_hours,
                0
            ) ot_hours

        FROM workers w

        LEFT JOIN attendance a
        ON a.worker_id=w.id
        AND a.month_year=?

        ORDER BY w.id
        """,
        (month,)
    )


    return render_template(
        "attendance.html",
        data=data,
        month=month
    )


# =========================================================
# ADVANCES
# =========================================================

@app.route(
    "/advances",
    methods=["GET", "POST"]
)
@login_required
def advances():

    if request.method == "POST":

        form = request.form

        db.execute(
            """
            INSERT INTO worker_advances
            (
                worker_id,
                month_year,
                advance_date,
                amount,
                note
            )
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                int(form["worker_id"]),

                form["month_year"],

                form["advance_date"],

                float(form["amount"]),

                form.get("note", "")
            )
        )

        log(
            "Added advance"
        )

        flash(
            "Advance saved",
            "success"
        )

        return redirect(
            url_for("advances")
        )


    query = request.args.get(
        "q",
        ""
    ).strip()


    data = rows(
        """
        SELECT
            a.*,
            w.name,
            w.bangla_name

        FROM worker_advances a

        LEFT JOIN workers w
        ON w.id=a.worker_id

        WHERE
            w.name LIKE ?
            OR COALESCE(
                w.bangla_name,
                ''
            ) LIKE ?

        ORDER BY a.id DESC
        """,
        (
            "%" + query + "%",
            "%" + query + "%"
        )
    )


    worker_list = rows(
        """
        SELECT *
        FROM workers
        ORDER BY name
        """
    )


    return render_template(
        "advances.html",
        advances=data,
        workers=worker_list,
        q=query
    )


# =========================================================
# DELETE ADVANCE
# =========================================================

@app.route(
    "/advances/delete/<int:aid>"
)
@login_required
def delete_advance(aid):

    db.execute(
        """
        DELETE FROM worker_advances
        WHERE id=?
        """,
        (aid,)
    )

    log(
        f"Deleted advance {aid}"
    )

    return redirect(
        url_for("advances")
    )


# =========================================================
# PAYROLL REPORT
# =========================================================

@app.route("/report")
@login_required
def report():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q",
        ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )


    data = payroll_rows(
        month,
        search,
        department
    )


    return render_template(
        "report.html",

        data=data,

        report_rows=data,

        month=month,

        q=search,

        selected_department=department
    )


# =========================================================
# CSV REPORT
# =========================================================

@app.route("/report.csv")
@login_required
def report_csv():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q",
        ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )


    data = payroll_rows(
        month,
        search,
        department
    )


    output = io.StringIO()

    writer = csv.writer(output)

    writer.writerow([
        "Worker ID",
        "Name",
        "Bangla Name",
        "Department",
        "Designation",
        "Present",
        "Absent",
        "OT Hours",
        "OT Amount",
        "Basic",
        "Refreshment",
        "Gross",
        "Advance",
        "Net Payable"
    ])


    for worker in data:

        writer.writerow([
            worker["id"],
            worker["name"],
            worker.get(
                "bangla_name",
                ""
            ),
            worker["department"],
            worker.get(
                "designation",
                ""
            ),
            worker["present_days"],
            worker["absent_days"],
            worker["ot_hours"],
            round(
                worker["ot_amount"],
                2
            ),
            worker["basic_salary"],
            round(
                worker["refreshment"],
                2
            ),
            round(
                worker["gross"],
                2
            ),
            round(
                worker["advance"],
                2
            ),
            round(
                worker["net"],
                2
            )
        ])


    file_data = io.BytesIO(
        output.getvalue().encode(
            "utf-8-sig"
        )
    )

    file_data.seek(0)


    return send_file(
        file_data,
        as_attachment=True,
        download_name=
            f"salary_{month}.csv",
        mimetype="text/csv"
    )


# =========================================================
# EXCEL REPORT
# =========================================================

@app.route("/report.xlsx")
@login_required
def report_xlsx():

    try:

        from openpyxl import Workbook

        from openpyxl.styles import (
            Font,
            Alignment
        )

    except Exception:

        abort(500)


    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q",
        ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )


    data = payroll_rows(
        month,
        search,
        department
    )


    workbook = Workbook()

    sheet = workbook.active

    sheet.title = "Salary Report"


    sheet.append([
        setting("company_name")
    ])

    sheet.append([
        f"Salary Report - {month}"
    ])


    sheet.append([
        "ID",
        "Name",
        "Bangla Name",
        "Department",
        "Designation",
        "Present",
        "Absent",
        "OT Hours",
        "OT Amount",
        "Basic",
        "Refreshment",
        "Gross",
        "Advance",
        "Net Payable"
    ])


    for cell in sheet[1]:

        cell.font = Font(
            bold=True,
            size=14
        )


    for worker in data:

        sheet.append([
            worker["id"],
            worker["name"],
            worker.get(
                "bangla_name",
                ""
            ),
            dept(
                worker["department"]
            ),
            worker.get(
                "designation",
                ""
            ),
            worker["present_days"],
            worker["absent_days"],
            worker["ot_hours"],
            round(
                worker["ot_amount"],
                2
            ),
            worker["basic_salary"],
            round(
                worker["refreshment"],
                2
            ),
            round(
                worker["gross"],
                2
            ),
            round(
                worker["advance"],
                2
            ),
            round(
                worker["net"],
                2
            )
        ])


    for column in sheet.columns:

        max_length = 0

        for cell in column:

            value = str(
                cell.value or ""
            )

            max_length = max(
                max_length,
                len(value)
            )

        letter = column[0].column_letter

        sheet.column_dimensions[
            letter
        ].width = min(
            max(
                max_length + 2,
                10
            ),
            30
        )


    file_data = io.BytesIO()

    workbook.save(file_data)

    file_data.seek(0)


    return send_file(
        file_data,
        as_attachment=True,
        download_name=
            f"salary_{month}.xlsx",

        mimetype=
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


# =========================================================
# PDF REPORT
# =========================================================

@app.route("/report.pdf")
@login_required
def report_pdf():

    try:

        from reportlab.lib.pagesizes import (
            A4,
            landscape
        )

        from reportlab.pdfgen import canvas

    except Exception:

        abort(500)


    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q",
        ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )


    data = payroll_rows(
        month,
        search,
        department
    )


    file_data = io.BytesIO()

    page_size = landscape(A4)

    canvas_obj = canvas.Canvas(
        file_data,
        pagesize=page_size
    )

    width, height = page_size


    canvas_obj.setFont(
        "Helvetica-Bold",
        14
    )

    canvas_obj.drawString(
        30,
        height - 30,
        setting("company_name")
        or "Reedoy Textile"
    )


    canvas_obj.setFont(
        "Helvetica",
        10
    )

    canvas_obj.drawString(
        30,
        height - 48,
        f"Salary Report - {month}"
    )


    headers = [
        "ID",
        "Name",
        "Department",
        "Present",
        "Absent",
        "OT",
        "OT Amount",
        "Gross",
        "Advance",
        "Net"
    ]


    positions = [
        30,
        60,
        180,
        300,
        350,
        400,
        435,
        500,
        560,
        620
    ]


    y = height - 75


    canvas_obj.setFont(
        "Helvetica-Bold",
        8
    )


    for index, header in enumerate(headers):

        canvas_obj.drawString(
            positions[index],
            y,
            header
        )


    canvas_obj.setFont(
        "Helvetica",
        8
    )


    y -= 15


    for worker in data:

        values = [
            worker["id"],
            worker["name"],
            worker["department"],
            worker["present_days"],
            worker["absent_days"],
            worker["ot_hours"],
            round(
                worker["ot_amount"],
                2
            ),
            round(
                worker["gross"],
                2
            ),
            round(
                worker["advance"],
                2
            ),
            round(
                worker["net"],
                2
            )
        ]


        for index, value in enumerate(values):

            canvas_obj.drawString(
                positions[index],
                y,
                str(value)[:25]
            )


        y -= 13


        if y < 30:

            canvas_obj.showPage()

            y = height - 40

            canvas_obj.setFont(
                "Helvetica",
                8
            )


    canvas_obj.save()

    file_data.seek(0)


    return send_file(
        file_data,
        as_attachment=True,
        download_name=
            f"salary_{month}.pdf",
        mimetype="application/pdf"
    )


# =========================================================
# DEPARTMENT SALARY
# =========================================================

@app.route("/department")
@login_required
def department():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    selected = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )


    data = payroll_rows(
        month,
        "",
        selected
    )


    return render_template(
        "department.html",

        data=data,

        month=month,

        selected=selected
    )


# =========================================================
# PAYSLIP
# =========================================================

@app.route("/payslip/<int:wid>")
@login_required
def payslip(wid):

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )


    data = payroll_rows(
        month
    )


    worker = next(
        (
            item
            for item in data
            if item["id"] == wid
        ),
        None
    )


    if not worker:
        abort(404)


    return render_template(
        "payslip.html",

        w=worker,

        month=month,

        gross=worker["gross"],

        net=worker["net"]
    )


# =========================================================
# COMPANY SETTINGS
# =========================================================

@app.route(
    "/settings",
    methods=["GET", "POST"]
)
@admin_required
def settings():

    if request.method == "POST":

        keys = [
            "company_name",
            "company_address",
            "company_phone",
            "company_email"
        ]


        for key in keys:

            value = request.form.get(
                key,
                ""
            )


            db.execute(
                """
                INSERT INTO company_settings
                (
                    key,
                    value
                )
                VALUES (?, ?)

                ON CONFLICT(key)
                DO UPDATE SET
                    value=excluded.value
                """,
                (
                    key,
                    value
                )
            )


        log(
            "Updated company settings"
        )

        flash(
            "Settings saved",
            "success"
        )


        return redirect(
            url_for("settings")
        )


    settings_data = {

        key: setting(key)

        for key in [
            "company_name",
            "company_address",
            "company_phone",
            "company_email"
        ]
    }


    return render_template(
        "settings.html",
        settings=settings_data
    )


# =========================================================
# USER MANAGEMENT
# =========================================================

@app.route(
    "/users",
    methods=["GET", "POST"]
)
@admin_required
def users():

    if request.method == "POST":

        form = request.form

        username = form[
            "username"
        ].strip()

        password = form[
            "password"
        ]

        role = form.get(
            "role",
            "User"
        )


        if not username or not password:

            flash(
                "Username and password required",
                "danger"
            )

        else:

            try:

                db.execute(
                    """
                    INSERT INTO users
                    (
                        username,
                        password_hash,
                        full_name,
                        role,
                        active,
                        created_at
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    """,
                    (
                        username,
                        sha(password),
                        form.get(
                            "full_name",
                            ""
                        ),
                        role,
                        1,
                        datetime.now().isoformat()
                    )
                )


                log(
                    "Created user "
                    + username
                )


                flash(
                    "User created",
                    "success"
                )


            except Exception:

                flash(
                    "Username already exists",
                    "danger"
                )


    user_list = rows(
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


    return render_template(
        "users.html",
        users=user_list
    )


# =========================================================
# TOGGLE USER
# =========================================================

@app.route(
    "/users/toggle/<int:uid>"
)
@admin_required
def toggle_user(uid):

    db.execute(
        """
        UPDATE users

        SET active =
            CASE
                WHEN active=1
                THEN 0
                ELSE 1
            END

        WHERE id=?
        """,
        (uid,)
    )


    log(
        f"Toggled user {uid}"
    )


    return redirect(
        url_for("users")
    )


# =========================================================
# CHANGE PASSWORD
# =========================================================

@app.route(
    "/users/password/<int:uid>",
    methods=["POST"]
)
@admin_required
def change_password(uid):

    password = request.form.get(
        "password",
        ""
    )


    if not password:

        flash(
            "Password required",
            "danger"
        )

        return redirect(
            url_for("users")
        )


    db.execute(
        """
        UPDATE users

        SET password_hash=?

        WHERE id=?
        """,
        (
            sha(password),
            uid
        )
    )


    log(
        f"Changed password for user {uid}"
    )


    flash(
        "Password changed",
        "success"
    )


    return redirect(
        url_for("users")
    )


# =========================================================
# ACTIVITY LOG
# =========================================================

@app.route("/activity")
@admin_required
def activity():

    logs = rows(
        """
        SELECT *
        FROM activity_log
        ORDER BY id DESC
        LIMIT 500
        """
    )


    return render_template(
        "activity.html",
        logs=logs
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(403)
def forbidden(error):

    return (
        "403 - Access Forbidden",
        403
    )


@app.errorhandler(404)
def not_found(error):

    return (
        "404 - Page Not Found",
        404
    )


@app.errorhandler(500)
def internal_error(error):

    return (
        "500 - Internal Server Error",
        500
    )


# =========================================================
# INITIALIZE DATABASE
# =========================================================

init_db()


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
