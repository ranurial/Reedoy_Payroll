```python
import os
import csv
import io
import sqlite3
import hashlib
from datetime import datetime
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
    "REEDOY-CHANGE-THIS-SECRET"
)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

SQLITE_PATH = os.environ.get(
    "SQLITE_PATH",
    os.path.join(os.path.dirname(__file__), "reedoy_payroll.db")
)

# =========================================================
# TRANSLATION
# =========================================================

BN = {
    "Dashboard": "ড্যাশবোর্ড",
    "Workers": "শ্রমিক",
    "Attendance": "উপস্থিতি",
    "Report": "রিপোর্ট",
    "Department Salary": "বিভাগভিত্তিক বেতন",
    "Payslip": "বেতন স্লিপ",
    "Advance Salary": "অগ্রিম বেতন",
    "Settings": "সেটিংস",
    "Users": "ইউজার",
    "Activity": "কার্যক্রম",
    "Logout": "লগআউট",
    "Login": "লগইন",
    "Save": "সংরক্ষণ",
    "Update": "আপডেট",
    "Delete": "মুছে ফেলুন",
    "Search": "অনুসন্ধান",
    "Name": "নাম",
    "Bangla Name": "বাংলা নাম",
    "Department": "বিভাগ",
    "Basic Salary": "মূল বেতন",
    "Present Days": "উপস্থিত দিন",
    "Absent Days": "অনুপস্থিত দিন",
    "OT Hours": "ওভারটাইম ঘণ্টা",
    "OT Rate": "ওভারটাইম রেট",
    "OT Amount": "ওভারটাইম টাকা",
    "Refreshment": "নাস্তা",
    "Advance": "অগ্রিম",
    "Gross Salary": "মোট বেতন",
    "Net Salary": "প্রাপ্য বেতন",
    "Month": "মাস",
    "All Departments": "সকল বিভাগ",
    "No records found": "কোনো রেকর্ড পাওয়া যায়নি",
    "No department records found": "কোনো বিভাগীয় রেকর্ড পাওয়া যায়নি",
}

DEPT_BN = {
    "All Departments": "সকল বিভাগ"
}


# =========================================================
# DATABASE
# =========================================================

def get_db():
    """
    PostgreSQL on Render.
    SQLite locally.
    """

    if DATABASE_URL:
        import psycopg2
        from psycopg2.extras import RealDictCursor

        conn = psycopg2.connect(DATABASE_URL)

        return conn

    conn = sqlite3.connect(SQLITE_PATH)
    conn.row_factory = sqlite3.Row

    return conn


def is_postgres():
    return bool(DATABASE_URL)


def close_db(conn):
    try:
        conn.close()
    except Exception:
        pass


# =========================================================
# DATABASE QUERY HELPERS
# =========================================================

def fetch_all(sql, params=()):
    conn = get_db()

    try:
        if is_postgres():
            from psycopg2.extras import RealDictCursor

            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute(sql, params)
            rows = cur.fetchall()

        else:
            cur = conn.cursor()
            cur.execute(sql, params)
            rows = cur.fetchall()

        return rows

    finally:
        close_db(conn)


def fetch_one(sql, params=()):
    conn = get_db()

    try:
        if is_postgres():
            from psycopg2.extras import RealDictCursor

            cur = conn.cursor(cursor_factory=RealDictCursor)
            cur.execute(sql, params)
            row = cur.fetchone()

        else:
            cur = conn.cursor()
            cur.execute(sql, params)
            row = cur.fetchone()

        return row

    finally:
        close_db(conn)


def execute_sql(sql, params=()):
    conn = get_db()

    try:
        cur = conn.cursor()
        cur.execute(sql, params)
        conn.commit()

        try:
            return cur.lastrowid
        except Exception:
            return None

    finally:
        close_db(conn)


# =========================================================
# PASSWORD
# =========================================================

def hash_password(password):
    return hashlib.sha256(
        password.encode("utf-8")
    ).hexdigest()


def check_password(password, password_hash):
    return hash_password(password) == password_hash


# =========================================================
# INIT DATABASE
# =========================================================

def init_db():

    conn = get_db()

    try:
        cur = conn.cursor()

        if is_postgres():

            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id SERIAL PRIMARY KEY,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT DEFAULT 'User'
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS company_settings (
                    id SERIAL PRIMARY KEY,
                    setting_key TEXT UNIQUE NOT NULL,
                    setting_value TEXT
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS workers (
                    id INTEGER PRIMARY KEY,
                    name TEXT,
                    bangla_name TEXT,
                    department TEXT,
                    designation TEXT,
                    basic_salary DOUBLE PRECISION DEFAULT 0,
                    ot_rate DOUBLE PRECISION DEFAULT 0,
                    refreshment_bill DOUBLE PRECISION DEFAULT 0,
                    phone TEXT,
                    address TEXT,
                    joining_date TEXT,
                    status TEXT DEFAULT 'Active'
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS attendance (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER,
                    month TEXT,
                    present_days DOUBLE PRECISION DEFAULT 0,
                    absent_days DOUBLE PRECISION DEFAULT 0,
                    ot_hours DOUBLE PRECISION DEFAULT 0,
                    UNIQUE(worker_id, month)
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS worker_advances (
                    id SERIAL PRIMARY KEY,
                    worker_id INTEGER,
                    advance_date TEXT,
                    amount DOUBLE PRECISION DEFAULT 0,
                    note TEXT
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS activity_log (
                    id SERIAL PRIMARY KEY,
                    username TEXT,
                    action TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

        else:

            cur.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    role TEXT DEFAULT 'User'
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS company_settings (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    setting_key TEXT UNIQUE NOT NULL,
                    setting_value TEXT
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS workers (
                    id INTEGER PRIMARY KEY,
                    name TEXT,
                    bangla_name TEXT,
                    department TEXT,
                    designation TEXT,
                    basic_salary REAL DEFAULT 0,
                    ot_rate REAL DEFAULT 0,
                    refreshment_bill REAL DEFAULT 0,
                    phone TEXT,
                    address TEXT,
                    joining_date TEXT,
                    status TEXT DEFAULT 'Active'
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS attendance (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER,
                    month TEXT,
                    present_days REAL DEFAULT 0,
                    absent_days REAL DEFAULT 0,
                    ot_hours REAL DEFAULT 0,
                    UNIQUE(worker_id, month)
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS worker_advances (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    worker_id INTEGER,
                    advance_date TEXT,
                    amount REAL DEFAULT 0,
                    note TEXT
                )
            """)

            cur.execute("""
                CREATE TABLE IF NOT EXISTS activity_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT,
                    action TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

        # -------------------------------------------------
        # DEFAULT ADMIN
        # -------------------------------------------------

        if is_postgres():
            cur.execute(
                "SELECT id FROM users WHERE username=%s",
                ("admin",)
            )
        else:
            cur.execute(
                "SELECT id FROM users WHERE username=?",
                ("admin",)
            )

        admin = cur.fetchone()

        if not admin:

            if is_postgres():
                cur.execute("""
                    INSERT INTO users
                    (username, password_hash, role)
                    VALUES (%s, %s, %s)
                """, (
                    "admin",
                    hash_password("admin123"),
                    "Administrator"
                ))

            else:
                cur.execute("""
                    INSERT INTO users
                    (username, password_hash, role)
                    VALUES (?, ?, ?)
                """, (
                    "admin",
                    hash_password("admin123"),
                    "Administrator"
                ))

        # -------------------------------------------------
        # DEFAULT COMPANY SETTINGS
        # -------------------------------------------------

        defaults = {
            "company_name":
                "Reedoy Textile Dyeing Printing & Finishing",
            "address":
                "",
            "phone":
                "",
            "language":
                "English"
        }

        for key, value in defaults.items():

            if is_postgres():
                cur.execute(
                    """
                    SELECT id
                    FROM company_settings
                    WHERE setting_key=%s
                    """,
                    (key,)
                )
            else:
                cur.execute(
                    """
                    SELECT id
                    FROM company_settings
                    WHERE setting_key=?
                    """,
                    (key,)
                )

            exists = cur.fetchone()

            if not exists:

                if is_postgres():
                    cur.execute("""
                        INSERT INTO company_settings
                        (setting_key, setting_value)
                        VALUES (%s, %s)
                    """, (key, value))

                else:
                    cur.execute("""
                        INSERT INTO company_settings
                        (setting_key, setting_value)
                        VALUES (?, ?)
                    """, (key, value))

        conn.commit()

    finally:
        close_db(conn)


# =========================================================
# LOGIN
# =========================================================

def login_required(f):

    @wraps(f)
    def decorated(*args, **kwargs):

        if not session.get("user_id"):
            return redirect(url_for("login"))

        return f(*args, **kwargs)

    return decorated


def admin_required(f):

    @wraps(f)
    def decorated(*args, **kwargs):

        if not session.get("user_id"):
            return redirect(url_for("login"))

        if session.get("role") != "Administrator":
            flash("Administrator access required.", "danger")
            return redirect(url_for("dashboard"))

        return f(*args, **kwargs)

    return decorated


# =========================================================
# ACTIVITY LOG
# =========================================================

def log_activity(action):

    try:

        username = session.get("username", "system")

        if is_postgres():

            execute_sql("""
                INSERT INTO activity_log
                (username, action)
                VALUES (%s, %s)
            """, (username, action))

        else:

            execute_sql("""
                INSERT INTO activity_log
                (username, action)
                VALUES (?, ?)
            """, (username, action))

    except Exception:
        pass


# =========================================================
# COMPANY SETTINGS
# =========================================================

def get_settings():

    rows = fetch_all("""
        SELECT setting_key, setting_value
        FROM company_settings
    """)

    result = {}

    for row in rows:

        if isinstance(row, dict):
            result[row["setting_key"]] = row["setting_value"]

        else:
            result[row["setting_key"]] = row["setting_value"]

    return result


# =========================================================
# SALARY CALCULATION
# =========================================================

def calculate_salary(worker, attendance=None, advance=0):

    attendance = attendance or {}

    try:
        basic = float(
            worker.get("basic_salary", 0) or 0
        )
    except Exception:
        basic = 0

    try:
        present = float(
            attendance.get("present_days", 0) or 0
        )
    except Exception:
        present = 0

    try:
        absent = float(
            attendance.get("absent_days", 0) or 0
        )
    except Exception:
        absent = 0

    try:
        ot_hours = float(
            attendance.get("ot_hours", 0) or 0
        )
    except Exception:
        ot_hours = 0

    try:
        ot_rate = float(
            worker.get("ot_rate", 0) or 0
        )
    except Exception:
        ot_rate = 0

    try:
        refreshment = float(
            worker.get("refreshment_bill", 0) or 0
        )
    except Exception:
        refreshment = 0

    try:
        advance = float(advance or 0)
    except Exception:
        advance = 0

    # 30-day salary calculation
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
        "basic_salary": basic,
        "present_days": present,
        "absent_days": absent,
        "ot_hours": ot_hours,
        "ot_rate": ot_rate,
        "earned_basic": earned_basic,
        "ot_amount": ot_amount,
        "refreshment": refreshment,
        "gross": gross,
        "advance": advance,
        "net": net,
        "net_salary": net
    }


# =========================================================
# DEPARTMENTS
# =========================================================

def get_departments():

    rows = fetch_all("""
        SELECT DISTINCT department
        FROM workers
        WHERE department IS NOT NULL
          AND TRIM(department) <> ''
        ORDER BY department
    """)

    departments = ["All Departments"]

    for row in rows:

        if isinstance(row, dict):
            dep = row["department"]
        else:
            dep = row["department"]

        if dep and dep not in departments:
            departments.append(dep)

    return departments


# =========================================================
# PAYROLL ROWS
# =========================================================

def payroll_rows(
    month,
    search="",
    department="All Departments"
):

    search = search or ""

    # -----------------------------------------------------
    # WORKERS
    # -----------------------------------------------------

    if is_postgres():

        sql = """
            SELECT
                w.*,
                a.present_days,
                a.absent_days,
                a.ot_hours
            FROM workers w
            LEFT JOIN attendance a
                ON a.worker_id = w.id
                AND a.month = %s
            WHERE
                (
                    %s = ''
                    OR w.name ILIKE %s
                    OR w.bangla_name ILIKE %s
                    OR w.department ILIKE %s
                )
                AND
                (
                    %s = 'All Departments'
                    OR w.department = %s
                )
            ORDER BY w.id
        """

        like = "%" + search + "%"

        rows = fetch_all(
            sql,
            (
                month,
                search,
                like,
                like,
                like,
                department,
                department
            )
        )

    else:

        sql = """
            SELECT
                w.*,
                a.present_days,
                a.absent_days,
                a.ot_hours
            FROM workers w
            LEFT JOIN attendance a
                ON a.worker_id = w.id
                AND a.month = ?
            WHERE
                (
                    ? = ''
                    OR w.name LIKE ?
                    OR w.bangla_name LIKE ?
                    OR w.department LIKE ?
                )
                AND
                (
                    ? = 'All Departments'
                    OR w.department = ?
                )
            ORDER BY w.id
        """

        like = "%" + search + "%"

        rows = fetch_all(
            sql,
            (
                month,
                search,
                like,
                like,
                like,
                department,
                department
            )
        )

    # -----------------------------------------------------
    # BUILD RESULT
    # -----------------------------------------------------

    data = []

    for row in rows:

        if isinstance(row, dict):
            worker = dict(row)
        else:
            worker = dict(row)

        # Attendance
        attendance = {
            "present_days":
                worker.get("present_days", 0) or 0,

            "absent_days":
                worker.get("absent_days", 0) or 0,

            "ot_hours":
                worker.get("ot_hours", 0) or 0
        }

        worker_id = worker.get("id")

        # -------------------------------------------------
        # ADVANCE
        # -------------------------------------------------

        advance = 0

        if worker_id:

            if is_postgres():

                adv_row = fetch_one("""
                    SELECT COALESCE(SUM(amount), 0) AS total
                    FROM worker_advances
                    WHERE worker_id = %s
                      AND (
                          advance_date IS NULL
                          OR LEFT(advance_date, 7) = %s
                      )
                """, (
                    worker_id,
                    month
                ))

            else:

                adv_row = fetch_one("""
                    SELECT COALESCE(SUM(amount), 0) AS total
                    FROM worker_advances
                    WHERE worker_id = ?
                      AND (
                          advance_date IS NULL
                          OR substr(advance_date, 1, 7) = ?
                      )
                """, (
                    worker_id,
                    month
                ))

            if adv_row:

                if isinstance(adv_row, dict):
                    advance = adv_row.get("total", 0) or 0
                else:
                    advance = adv_row["total"] or 0

        # -------------------------------------------------
        # SALARY
        # -------------------------------------------------

        salary = calculate_salary(
            worker,
            attendance,
            advance
        )

        # Add ALL fields explicitly.
        worker.update({
            "earned_basic":
                salary["earned_basic"],

            "ot_amount":
                salary["ot_amount"],

            "refreshment":
                salary["refreshment"],

            "gross":
                salary["gross"],

            "advance":
                salary["advance"],

            "net":
                salary["net"],

            "net_salary":
                salary["net_salary"],

            "present_days":
                salary["present_days"],

            "absent_days":
                salary["absent_days"],

            "ot_hours":
                salary["ot_hours"],

            "ot_rate":
                salary["ot_rate"],

            "basic_salary":
                salary["basic_salary"]
        })

        data.append(worker)

    return data


# =========================================================
# HOME
# =========================================================

@app.route("/")
def home():

    if session.get("user_id"):
        return redirect(url_for("dashboard"))

    return redirect(url_for("login"))


# =========================================================
# LOGIN
# =========================================================

@app.route("/login", methods=["GET", "POST"])
def login():

    if request.method == "POST":

        username = request.form.get(
            "username", ""
        ).strip()

        password = request.form.get(
            "password", ""
        )

        if is_postgres():

            user = fetch_one("""
                SELECT *
                FROM users
                WHERE username=%s
            """, (username,))

        else:

            user = fetch_one("""
                SELECT *
                FROM users
                WHERE username=?
            """, (username,))

        if user:

            if isinstance(user, dict):
                password_hash = user["password_hash"]
                user_id = user["id"]
                role = user["role"]
            else:
                password_hash = user["password_hash"]
                user_id = user["id"]
                role = user["role"]

            if check_password(
                password,
                password_hash
            ):

                session["user_id"] = user_id
                session["username"] = username
                session["role"] = role

                log_activity("Login")

                return redirect(
                    url_for("dashboard")
                )

        flash(
            "Invalid username or password.",
            "danger"
        )

    return render_template("login.html")


# =========================================================
# LOGOUT
# =========================================================

@app.route("/logout")
def logout():

    log_activity("Logout")

    session.clear()

    return redirect(url_for("login"))


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/dashboard")
@login_required
def dashboard():

    worker_row = fetch_one(
        "SELECT COUNT(*) AS total FROM workers"
    )

    if isinstance(worker_row, dict):
        total_workers = worker_row["total"]
    else:
        total_workers = worker_row["total"]

    return render_template(
        "dashboard.html",
        total_workers=total_workers,
        settings=get_settings()
    )


# =========================================================
# WORKERS
# =========================================================

@app.route("/workers")
@login_required
def workers():

    rows = fetch_all("""
        SELECT *
        FROM workers
        ORDER BY id
    """)

    workers_data = [
        dict(row) if not isinstance(row, dict)
        else dict(row)
        for row in rows
    ]

    return render_template(
        "workers.html",
        workers=workers_data
    )


# =========================================================
# ADD WORKER
# =========================================================

@app.route("/workers/add", methods=["GET", "POST"])
@login_required
def add_worker():

    if request.method == "POST":

        name = request.form.get(
            "name", ""
        ).strip()

        bangla_name = request.form.get(
            "bangla_name", ""
        ).strip()

        department = request.form.get(
            "department", ""
        ).strip()

        designation = request.form.get(
            "designation", ""
        ).strip()

        basic_salary = float(
            request.form.get(
                "basic_salary", 0
            ) or 0
        )

        ot_rate = float(
            request.form.get(
                "ot_rate", 0
            ) or 0
        )

        refreshment_bill = float(
            request.form.get(
                "refreshment_bill", 0
            ) or 0
        )

        phone = request.form.get(
            "phone", ""
        ).strip()

        address = request.form.get(
            "address", ""
        ).strip()

        joining_date = request.form.get(
            "joining_date", ""
        ).strip()

        status = request.form.get(
            "status", "Active"
        )

        if is_postgres():

            execute_sql("""
                INSERT INTO workers
                (
                    name,
                    bangla_name,
                    department,
                    designation,
                    basic_salary,
                    ot_rate,
                    refreshment_bill,
                    phone,
                    address,
                    joining_date,
                    status
                )
                VALUES
                (
                    %s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s
                )
            """, (
                name,
                bangla_name,
                department,
                designation,
                basic_salary,
                ot_rate,
                refreshment_bill,
                phone,
                address,
                joining_date,
                status
            ))

        else:

            execute_sql("""
                INSERT INTO workers
                (
                    name,
                    bangla_name,
                    department,
                    designation,
                    basic_salary,
                    ot_rate,
                    refreshment_bill,
                    phone,
                    address,
                    joining_date,
                    status
                )
                VALUES
                (
                    ?,?,?,?,?,?,?,?,?,?,?
                )
            """, (
                name,
                bangla_name,
                department,
                designation,
                basic_salary,
                ot_rate,
                refreshment_bill,
                phone,
                address,
                joining_date,
                status
            ))

        log_activity(
            "Added worker: " + name
        )

        flash(
            "Worker added successfully.",
            "success"
        )

        return redirect(
            url_for("workers")
        )

    return render_template(
        "worker_form.html",
        worker=None
    )


# =========================================================
# EDIT WORKER
# =========================================================

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
        WHERE id=%s
        """ if is_postgres() else """
        SELECT *
        FROM workers
        WHERE id=?
        """,
        (worker_id,)
    )

    if not worker:
        abort(404)

    worker = (
        dict(worker)
        if not isinstance(worker, dict)
        else dict(worker)
    )

    if request.method == "POST":

        values = (
            request.form.get("name", "").strip(),
            request.form.get("bangla_name", "").strip(),
            request.form.get("department", "").strip(),
            request.form.get("designation", "").strip(),
            float(request.form.get("basic_salary", 0) or 0),
            float(request.form.get("ot_rate", 0) or 0),
            float(request.form.get("refreshment_bill", 0) or 0),
            request.form.get("phone", "").strip(),
            request.form.get("address", "").strip(),
            request.form.get("joining_date", "").strip(),
            request.form.get("status", "Active"),
            worker_id
        )

        if is_postgres():

            execute_sql("""
                UPDATE workers
                SET
                    name=%s,
                    bangla_name=%s,
                    department=%s,
                    designation=%s,
                    basic_salary=%s,
                    ot_rate=%s,
                    refreshment_bill=%s,
                    phone=%s,
                    address=%s,
                    joining_date=%s,
                    status=%s
                WHERE id=%s
            """, values)

        else:

            execute_sql("""
                UPDATE workers
                SET
                    name=?,
                    bangla_name=?,
                    department=?,
                    designation=?,
                    basic_salary=?,
                    ot_rate=?,
                    refreshment_bill=?,
                    phone=?,
                    address=?,
                    joining_date=?,
                    status=?
                WHERE id=?
            """, values)

        log_activity(
            f"Edited worker ID {worker_id}"
        )

        flash(
            "Worker updated successfully.",
            "success"
        )

        return redirect(
            url_for("workers")
        )

    return render_template(
        "worker_form.html",
        worker=worker
    )


# =========================================================
# DELETE WORKER
# =========================================================

@app.route(
    "/workers/delete/<int:worker_id>",
    methods=["POST"]
)
@admin_required
def delete_worker(worker_id):

    if is_postgres():

        execute_sql(
            "DELETE FROM workers WHERE id=%s",
            (worker_id,)
        )

    else:

        execute_sql(
            "DELETE FROM workers WHERE id=?",
            (worker_id,)
        )

    log_activity(
        f"Deleted worker ID {worker_id}"
    )

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

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    data = payroll_rows(
        month=month
    )

    return render_template(
        "attendance.html",
        data=data,
        month=month
    )


# =========================================================
# SAVE ATTENDANCE
# =========================================================

@app.route(
    "/attendance/save",
    methods=["POST"]
)
@login_required
def save_attendance():

    month = request.form.get(
        "month"
    ) or datetime.now().strftime("%Y-%m")

    worker_ids = request.form.getlist(
        "worker_id"
    )

    for worker_id in worker_ids:

        present = float(
            request.form.get(
                f"present_{worker_id}",
                0
            ) or 0
        )

        absent = float(
            request.form.get(
                f"absent_{worker_id}",
                0
            ) or 0
        )

        ot_hours = float(
            request.form.get(
                f"ot_{worker_id}",
                0
            ) or 0
        )

        if is_postgres():

            exists = fetch_one("""
                SELECT id
                FROM attendance
                WHERE worker_id=%s
                  AND month=%s
            """, (
                worker_id,
                month
            ))

        else:

            exists = fetch_one("""
                SELECT id
                FROM attendance
                WHERE worker_id=?
                  AND month=?
            """, (
                worker_id,
                month
            ))

        if exists:

            attendance_id = (
                exists["id"]
                if isinstance(exists, dict)
                else exists["id"]
            )

            if is_postgres():

                execute_sql("""
                    UPDATE attendance
                    SET
                        present_days=%s,
                        absent_days=%s,
                        ot_hours=%s
                    WHERE id=%s
                """, (
                    present,
                    absent,
                    ot_hours,
                    attendance_id
                ))

            else:

                execute_sql("""
                    UPDATE attendance
                    SET
                        present_days=?,
                        absent_days=?,
                        ot_hours=?
                    WHERE id=?
                """, (
                    present,
                    absent,
                    ot_hours,
                    attendance_id
                ))

        else:

            if is_postgres():

                execute_sql("""
                    INSERT INTO attendance
                    (
                        worker_id,
                        month,
                        present_days,
                        absent_days,
                        ot_hours
                    )
                    VALUES (%s,%s,%s,%s,%s)
                """, (
                    worker_id,
                    month,
                    present,
                    absent,
                    ot_hours
                ))

            else:

                execute_sql("""
                    INSERT INTO attendance
                    (
                        worker_id,
                        month,
                        present_days,
                        absent_days,
                        ot_hours
                    )
                    VALUES (?,?,?,?,?)
                """, (
                    worker_id,
                    month,
                    present,
                    absent,
                    ot_hours
                ))

    log_activity(
        f"Saved attendance for {month}"
    )

    flash(
        "Attendance saved successfully.",
        "success"
    )

    return redirect(
        url_for(
            "attendance",
            month=month
        )
    )


# =========================================================
# REPORT
# =========================================================

@app.route("/report")
@login_required
def report():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q", ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )

    data = payroll_rows(
        month=month,
        search=search,
        department=department
    )

    return render_template(
        "report.html",
        data=data,
        report_rows=data,
        month=month,
        q=search,
        search=search,
        selected=department,
        selected_department=department,
        department=department,
        departments=get_departments()
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
            "department"
        )
        or "All Departments"
    )

    data = payroll_rows(
        month=month,
        search="",
        department=selected
    )

    departments = get_departments()

    return render_template(
        "department.html",

        # Main data
        data=data,

        # Alternative names for compatibility
        report_rows=data,
        rows=data,
        workers=data,

        # Date / month
        month=month,

        # Department
        selected=selected,
        selected_department=selected,
        department=selected,

        # Department list
        departments=departments
    )


# =========================================================
# PAYSLIP
# =========================================================

@app.route("/payslip/<int:worker_id>")
@login_required
def payslip(worker_id):

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    data = payroll_rows(
        month=month
    )

    worker = None

    for row in data:

        if int(row["id"]) == worker_id:
            worker = row
            break

    if not worker:
        abort(404)

    settings = get_settings()

    return render_template(
        "payslip.html",
        worker=worker,
        data=worker,
        month=month,
        settings=settings
    )


# =========================================================
# ADVANCE SALARY
# =========================================================

@app.route("/advance")
@login_required
def advance():

    workers_rows = fetch_all("""
        SELECT id, name, bangla_name
        FROM workers
        ORDER BY id
    """)

    workers_data = [
        dict(row) if not isinstance(row, dict)
        else dict(row)
        for row in workers_rows
    ]

    advances = fetch_all("""
        SELECT
            a.*,
            w.name,
            w.bangla_name
        FROM worker_advances a
        LEFT JOIN workers w
            ON w.id = a.worker_id
        ORDER BY a.id DESC
    """)

    advances_data = [
        dict(row) if not isinstance(row, dict)
        else dict(row)
        for row in advances
    ]

    return render_template(
        "advance.html",
        workers=workers_data,
        advances=advances_data
    )


# =========================================================
# SAVE ADVANCE
# =========================================================

@app.route(
    "/advance/save",
    methods=["POST"]
)
@login_required
def save_advance():

    worker_id = request.form.get(
        "worker_id"
    )

    amount = float(
        request.form.get(
            "amount", 0
        ) or 0
    )

    advance_date = request.form.get(
        "advance_date"
    ) or datetime.now().strftime("%Y-%m-%d")

    note = request.form.get(
        "note", ""
    ).strip()

    if is_postgres():

        execute_sql("""
            INSERT INTO worker_advances
            (
                worker_id,
                advance_date,
                amount,
                note
            )
            VALUES (%s,%s,%s,%s)
        """, (
            worker_id,
            advance_date,
            amount,
            note
        ))

    else:

        execute_sql("""
            INSERT INTO worker_advances
            (
                worker_id,
                advance_date,
                amount,
                note
            )
            VALUES (?,?,?,?)
        """, (
            worker_id,
            advance_date,
            amount,
            note
        ))

    log_activity(
        f"Added advance for worker {worker_id}"
    )

    flash(
        "Advance salary saved.",
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

        company_name = request.form.get(
            "company_name", ""
        ).strip()

        address = request.form.get(
            "address", ""
        ).strip()

        phone = request.form.get(
            "phone", ""
        ).strip()

        language = request.form.get(
            "language", "English"
        )

        values = {
            "company_name": company_name,
            "address": address,
            "phone": phone,
            "language": language
        }

        for key, value in values.items():

            if is_postgres():

                execute_sql("""
                    INSERT INTO company_settings
                    (setting_key, setting_value)
                    VALUES (%s,%s)
                    ON CONFLICT(setting_key)
                    DO UPDATE SET
                        setting_value=EXCLUDED.setting_value
                """, (
                    key,
                    value
                ))

            else:

                exists = fetch_one("""
                    SELECT id
                    FROM company_settings
                    WHERE setting_key=?
                """, (key,))

                if exists:

                    execute_sql("""
                        UPDATE company_settings
                        SET setting_value=?
                        WHERE setting_key=?
                    """, (
                        value,
                        key
                    ))

                else:

                    execute_sql("""
                        INSERT INTO company_settings
                        (setting_key, setting_value)
                        VALUES (?,?)
                    """, (
                        key,
                        value
                    ))

        log_activity(
            "Updated company settings"
        )

        flash(
            "Settings updated.",
            "success"
        )

        return redirect(
            url_for("settings")
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

    rows = fetch_all("""
        SELECT id, username, role
        FROM users
        ORDER BY id
    """)

    users_data = [
        dict(row) if not isinstance(row, dict)
        else dict(row)
        for row in rows
    ]

    return render_template(
        "users.html",
        users=users_data
    )


# =========================================================
# ADD USER
# =========================================================

@app.route(
    "/users/add",
    methods=["POST"]
)
@admin_required
def add_user():

    username = request.form.get(
        "username", ""
    ).strip()

    password = request.form.get(
        "password", ""
    )

    role = request.form.get(
        "role", "User"
    )

    if not username or not password:

        flash(
            "Username and password are required.",
            "danger"
        )

        return redirect(
            url_for("users")
        )

    try:

        if is_postgres():

            execute_sql("""
                INSERT INTO users
                (
                    username,
                    password_hash,
                    role
                )
                VALUES (%s,%s,%s)
            """, (
                username,
                hash_password(password),
                role
            ))

        else:

            execute_sql("""
                INSERT INTO users
                (
                    username,
                    password_hash,
                    role
                )
                VALUES (?,?,?)
            """, (
                username,
                hash_password(password),
                role
            ))

        log_activity(
            f"Added user {username}"
        )

        flash(
            "User added successfully.",
            "success"
        )

    except Exception:

        flash(
            "Username already exists.",
            "danger"
        )

    return redirect(
        url_for("users")
    )


# =========================================================
# DELETE USER
# =========================================================

@app.route(
    "/users/delete/<int:user_id>",
    methods=["POST"]
)
@admin_required
def delete_user(user_id):

    # Never delete currently logged-in admin
    if user_id == session.get("user_id"):

        flash(
            "You cannot delete your own account.",
            "danger"
        )

        return redirect(
            url_for("users")
        )

    if is_postgres():

        execute_sql(
            "DELETE FROM users WHERE id=%s",
            (user_id,)
        )

    else:

        execute_sql(
            "DELETE FROM users WHERE id=?",
            (user_id,)
        )

    log_activity(
        f"Deleted user ID {user_id}"
    )

    flash(
        "User deleted.",
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

    rows = fetch_all("""
        SELECT *
        FROM activity_log
        ORDER BY id DESC
        LIMIT 500
    """)

    logs = [
        dict(row) if not isinstance(row, dict)
        else dict(row)
        for row in rows
    ]

    return render_template(
        "activity.html",
        logs=logs,
        activity=logs
    )


# =========================================================
# CSV EXPORT
# =========================================================

@app.route("/report/csv")
@login_required
def report_csv():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q", ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )

    data = payroll_rows(
        month=month,
        search=search,
        department=department
    )

    output = io.StringIO()

    writer = csv.writer(output)

    writer.writerow([
        "ID",
        "Name",
        "Bangla Name",
        "Department",
        "Basic Salary",
        "Present Days",
        "Absent Days",
        "OT Hours",
        "OT Amount",
        "Refreshment",
        "Advance",
        "Gross Salary",
        "Net Salary"
    ])

    for row in data:

        writer.writerow([
            row.get("id", ""),
            row.get("name", ""),
            row.get("bangla_name", ""),
            row.get("department", ""),
            row.get("basic_salary", 0),
            row.get("present_days", 0),
            row.get("absent_days", 0),
            row.get("ot_hours", 0),
            row.get("ot_amount", 0),
            row.get("refreshment", 0),
            row.get("advance", 0),
            row.get("gross", 0),
            row.get("net_salary", 0)
        ])

    output.seek(0)

    filename = (
        f"Reedoy_Payroll_{month}.csv"
    )

    return send_file(
        io.BytesIO(
            output.getvalue().encode("utf-8-sig")
        ),
        mimetype="text/csv",
        as_attachment=True,
        download_name=filename
    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route("/report/excel")
@login_required
def report_excel():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q", ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )

    data = payroll_rows(
        month=month,
        search=search,
        department=department
    )

    try:

        from openpyxl import Workbook

    except ImportError:

        flash(
            "openpyxl is not installed.",
            "danger"
        )

        return redirect(
            url_for("report")
        )

    wb = Workbook()

    ws = wb.active

    ws.title = "Payroll"

    ws.append([
        "ID",
        "Name",
        "Bangla Name",
        "Department",
        "Basic Salary",
        "Present Days",
        "Absent Days",
        "OT Hours",
        "OT Amount",
        "Refreshment",
        "Advance",
        "Gross Salary",
        "Net Salary"
    ])

    for row in data:

        ws.append([
            row.get("id", ""),
            row.get("name", ""),
            row.get("bangla_name", ""),
            row.get("department", ""),
            row.get("basic_salary", 0),
            row.get("present_days", 0),
            row.get("absent_days", 0),
            row.get("ot_hours", 0),
            row.get("ot_amount", 0),
            row.get("refreshment", 0),
            row.get("advance", 0),
            row.get("gross", 0),
            row.get("net_salary", 0)
        ])

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    filename = (
        f"Reedoy_Payroll_{month}.xlsx"
    )

    return send_file(
        output,
        mimetype=(
            "application/vnd.openxmlformats-"
            "officedocument.spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=filename
    )


# =========================================================
# PDF EXPORT
# =========================================================

@app.route("/report/pdf")
@login_required
def report_pdf():

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q", ""
    ).strip()

    department = (
        request.args.get(
            "department",
            "All Departments"
        )
        or "All Departments"
    )

    data = payroll_rows(
        month=month,
        search=search,
        department=department
    )

    try:

        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.platypus import (
            SimpleDocTemplate,
            Table,
            TableStyle,
            Paragraph
        )

    except ImportError:

        flash(
            "reportlab is not installed.",
            "danger"
        )

        return redirect(
            url_for("report")
        )

    output = io.BytesIO()

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )

    styles = getSampleStyleSheet()

    story = []

    story.append(
        Paragraph(
            f"Reedoy Textile Payroll - {month}",
            styles["Title"]
        )
    )

    table_data = [[
        "ID",
        "Name",
        "Department",
        "Basic",
        "Present",
        "Absent",
        "OT",
        "OT Amount",
        "Refreshment",
        "Advance",
        "Net"
    ]]

    for row in data:

        table_data.append([
            row.get("id", ""),
            row.get("name", ""),
            row.get("department", ""),
            f"{float(row.get('basic_salary',0) or 0):.2f}",
            f"{float(row.get('present_days',0) or 0):.1f}",
            f"{float(row.get('absent_days',0) or 0):.1f}",
            f"{float(row.get('ot_hours',0) or 0):.1f}",
            f"{float(row.get('ot_amount',0) or 0):.2f}",
            f"{float(row.get('refreshment',0) or 0):.2f}",
            f"{float(row.get('advance',0) or 0):.2f}",
            f"{float(row.get('net_salary',0) or 0):.2f}"
        ])

    table = Table(
        table_data,
        repeatRows=1
    )

    table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
            ("ALIGN", (3, 1), (-1, -1), "RIGHT"),
        ])
    )

    story.append(table)

    doc.build(story)

    output.seek(0)

    filename = (
        f"Reedoy_Payroll_{month}.pdf"
    )

    return send_file(
        output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=filename
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(404)
def page_not_found(error):

    return render_template(
        "404.html"
    ), 404


@app.errorhandler(500)
def internal_error(error):

    return render_template(
        "500.html"
    ), 500


# =========================================================
# STARTUP
# =========================================================

try:
    init_db()
except Exception as e:
    print("Database initialization error:", e)


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
```
