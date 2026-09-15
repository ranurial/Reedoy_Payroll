import os
import io
import csv
import hashlib
import sqlite3
from datetime import datetime
from functools import wraps

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, send_file, abort
)

# =========================================================
# OPTIONAL PACKAGES
# =========================================================

try:
    import psycopg
    from psycopg.rows import dict_row
except ImportError:
    psycopg = None
    dict_row = None

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    )
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
except ImportError:
    SimpleDocTemplate = None


# =========================================================
# FLASK APP
# =========================================================

app = Flask(__name__)

app.secret_key = os.environ.get(
    "SECRET_KEY",
    "REEDOY-CHANGE-THIS-SECRET-KEY"
)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = "postgresql://" + DATABASE_URL[11:]

SQLITE_PATH = os.environ.get(
    "SQLITE_PATH",
    "reedoy_payroll.db"
)

DEFAULT_COMPANY_NAME = (
    "REEDOY TEXTILE DYEING PRINTING & FINISHING"
)


# =========================================================
# DATABASE HELPERS
# =========================================================

def using_postgres():
    return bool(DATABASE_URL and psycopg)


def db_connect():
    if using_postgres():
        return psycopg.connect(
            DATABASE_URL,
            row_factory=dict_row
        )

    conn = sqlite3.connect(SQLITE_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def execute(conn, sql, params=()):
    """
    Allows the same SQL to work with SQLite and PostgreSQL.
    App code uses ? placeholders.
    PostgreSQL uses %s.
    """
    if using_postgres():
        sql = sql.replace("?", "%s")

    return conn.execute(sql, params)


def fetchone(conn, sql, params=()):
    cur = execute(conn, sql, params)
    return cur.fetchone()


def fetchall(conn, sql, params=()):
    cur = execute(conn, sql, params)
    return cur.fetchall()


# =========================================================
# PASSWORD
# =========================================================

def hash_password(password):
    return hashlib.sha256(
        str(password).encode("utf-8")
    ).hexdigest()


# =========================================================
# DATABASE INITIALIZATION
# =========================================================

def init_db():

    conn = db_connect()

    if using_postgres():

        execute(conn, """
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT,
                full_name TEXT,
                role TEXT NOT NULL DEFAULT 'User',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                last_login TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS company_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS workers (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                basic_salary DOUBLE PRECISION NOT NULL DEFAULT 0,
                ot_rate DOUBLE PRECISION NOT NULL DEFAULT 0,
                department TEXT NOT NULL DEFAULT 'General',
                designation TEXT DEFAULT 'Worker',
                refreshment_bill DOUBLE PRECISION DEFAULT 0,
                bangla_name TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS attendance (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                present_days INTEGER DEFAULT 0,
                absent_days INTEGER DEFAULT 0,
                ot_hours DOUBLE PRECISION DEFAULT 0,
                advance_deduction DOUBLE PRECISION DEFAULT 0,
                UNIQUE(worker_id, month_year)
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS daily_attendance (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                day INTEGER NOT NULL,
                status TEXT DEFAULT 'P',
                ot_hours DOUBLE PRECISION DEFAULT 0,
                UNIQUE(worker_id, month_year, day)
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount DOUBLE PRECISION NOT NULL DEFAULT 0,
                note TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS activity_log (
                id SERIAL PRIMARY KEY,
                username TEXT,
                action TEXT,
                log_time TEXT NOT NULL
            )
        """)

    else:

        execute(conn, """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT,
                full_name TEXT,
                role TEXT NOT NULL DEFAULT 'User',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                last_login TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS company_settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS workers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                basic_salary REAL NOT NULL DEFAULT 0,
                ot_rate REAL NOT NULL DEFAULT 0,
                department TEXT NOT NULL DEFAULT 'General',
                designation TEXT DEFAULT 'Worker',
                refreshment_bill REAL DEFAULT 0,
                bangla_name TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                present_days INTEGER DEFAULT 0,
                absent_days INTEGER DEFAULT 0,
                ot_hours REAL DEFAULT 0,
                advance_deduction REAL DEFAULT 0,
                UNIQUE(worker_id, month_year)
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS daily_attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                day INTEGER NOT NULL,
                status TEXT DEFAULT 'P',
                ot_hours REAL DEFAULT 0,
                UNIQUE(worker_id, month_year, day)
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS worker_advances (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                worker_id INTEGER,
                month_year TEXT NOT NULL,
                advance_date TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                note TEXT
            )
        """)

        execute(conn, """
            CREATE TABLE IF NOT EXISTS activity_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                action TEXT,
                log_time TEXT NOT NULL
            )
        """)

    # -----------------------------------------------------
    # ADMIN USER
    # -----------------------------------------------------

    admin = fetchone(
        conn,
        "SELECT id, password_hash FROM users WHERE username=? LIMIT 1",
        ("admin",)
    )

    if not admin:

        execute(conn, """
            INSERT INTO users
            (username, password_hash, full_name, role, active, created_at)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            "admin",
            hash_password("admin123"),
            "System Administrator",
            "Administrator",
            1,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

    elif not admin.get("password_hash"):

        execute(conn, """
            UPDATE users
            SET password_hash=?,
                full_name=?,
                role=?,
                active=1
            WHERE username=?
        """, (
            hash_password("admin123"),
            "System Administrator",
            "Administrator",
            "admin"
        ))

    # -----------------------------------------------------
    # COMPANY SETTINGS
    # -----------------------------------------------------

    defaults = {
        "company_name": DEFAULT_COMPANY_NAME,
        "company_address": "",
        "company_phone": "",
        "company_email": "",
        "company_logo": ""
    }

    for key, value in defaults.items():

        existing = fetchone(
            conn,
            "SELECT key FROM company_settings WHERE key=?",
            (key,)
        )

        if not existing:

            execute(conn, """
                INSERT INTO company_settings(key, value)
                VALUES (?, ?)
            """, (key, value))

    conn.commit()
    conn.close()


# =========================================================
# COMPANY SETTINGS
# =========================================================

def get_company_settings():

    conn = db_connect()

    rows = fetchall(
        conn,
        "SELECT key, value FROM company_settings"
    )

    conn.close()

    data = {
        "company_name": DEFAULT_COMPANY_NAME,
        "company_address": "",
        "company_phone": "",
        "company_email": "",
        "company_logo": ""
    }

    for row in rows:
        data[row["key"]] = row["value"] or ""

    return data


# =========================================================
# ACTIVITY LOG
# =========================================================

def log_activity(action):

    try:

        username = session.get("username", "system")

        conn = db_connect()

        execute(conn, """
            INSERT INTO activity_log
            (username, action, log_time)
            VALUES (?, ?, ?)
        """, (
            username,
            action,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ))

        conn.commit()
        conn.close()

    except Exception:
        pass


# =========================================================
# LOGIN DECORATORS
# =========================================================

def login_required(func):

    @wraps(func)
    def wrapper(*args, **kwargs):

        if "user_id" not in session:
            return redirect(url_for("login"))

        return func(*args, **kwargs)

    return wrapper


def admin_required(func):

    @wraps(func)
    def wrapper(*args, **kwargs):

        if "user_id" not in session:
            return redirect(url_for("login"))

        if session.get("role") != "Administrator":
            flash(
                "Administrator access required.",
                "danger"
            )
            return redirect(url_for("dashboard"))

        return func(*args, **kwargs)

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

    # Salary is based on 30 days.
    daily_basic = basic / 30.0

    earned_basic = daily_basic * present

    # OT
    ot_amount = ot_hours * ot_rate

    # Gross
    gross = (
        earned_basic
        + ot_amount
        + refreshment
    )

    # Net
    net = gross - advance

    return {
        "earned_basic": round(earned_basic, 2),
        "absent_days": absent,
        "ot_amount": round(ot_amount, 2),
        "refreshment": round(refreshment, 2),
        "gross": round(gross, 2),
        "advance": round(advance, 2),
        "net": round(net, 2)
    }


# =========================================================
# PAYROLL ROWS
# =========================================================

def payroll_rows(
    month,
    search="",
    department="All Departments"
):

    conn = db_connect()

    sql = """
        SELECT
            w.id,
            w.name,
            w.bangla_name,
            w.department,
            w.designation,
            w.basic_salary,
            w.ot_rate,
            w.refreshment_bill,

            COALESCE(a.present_days, 0)
                AS present_days,

            COALESCE(a.absent_days, 0)
                AS absent_days,

            COALESCE(a.ot_hours, 0)
                AS ot_hours,

            COALESCE(
                (
                    SELECT SUM(wa.amount)
                    FROM worker_advances wa
                    WHERE wa.worker_id = w.id
                    AND wa.month_year = ?
                ),
                0
            ) AS advance

        FROM workers w

        LEFT JOIN attendance a
            ON a.worker_id = w.id
            AND a.month_year = ?

        WHERE 1=1
    """

    params = [
        month,
        month
    ]

    if search:

        sql += """
            AND (
                CAST(w.id AS TEXT) LIKE ?
                OR LOWER(COALESCE(w.name, '')) LIKE ?
                OR LOWER(COALESCE(w.bangla_name, '')) LIKE ?
                OR LOWER(COALESCE(w.department, '')) LIKE ?
            )
        """

        q = "%" + search.lower() + "%"

        params.extend([
            q,
            q,
            q,
            q
        ])

    if department and department != "All Departments":

        sql += """
            AND w.department = ?
        """

        params.append(department)

    sql += """
        ORDER BY w.id ASC
    """

    rows = fetchall(
        conn,
        sql,
        tuple(params)
    )

    conn.close()

    result = []

    for row in rows:

        worker = dict(row)

        salary = calculate_salary(worker)

        # IMPORTANT:
        # These fields MUST always exist.
        worker.update(salary)

        worker["ot_amount"] = salary["ot_amount"]
        worker["net_salary"] = salary["net"]

        # Extra safe defaults for Jinja.
        worker.setdefault("present_days", 0)
        worker.setdefault("absent_days", 0)
        worker.setdefault("ot_hours", 0)
        worker.setdefault("refreshment", 0)
        worker.setdefault("advance", 0)
        worker.setdefault("gross", 0)
        worker.setdefault("net", 0)
        worker.setdefault("ot_amount", 0)

        result.append(worker)

    return result


# =========================================================
# LOGIN
# =========================================================

@app.route("/login", methods=["GET", "POST"])
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

        conn = db_connect()

        user = fetchone(
            conn,
            """
            SELECT *
            FROM users
            WHERE username=?
            AND active=1
            LIMIT 1
            """,
            (username,)
        )

        if user:

            stored_hash = user.get("password_hash")

            # Legacy database compatibility.
            if not stored_hash:

                old_password = user.get("password")

                if old_password:
                    stored_hash = hash_password(
                        old_password
                    )

            if stored_hash == hash_password(password):

                execute(conn, """
                    UPDATE users
                    SET last_login=?
                    WHERE id=?
                """, (
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                    user["id"]
                ))

                conn.commit()
                conn.close()

                session["user_id"] = user["id"]
                session["username"] = user["username"]
                session["full_name"] = (
                    user.get("full_name")
                    or user["username"]
                )
                session["role"] = (
                    user.get("role")
                    or "User"
                )

                log_activity("Login")

                return redirect(
                    url_for("dashboard")
                )

        conn.close()

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

    return redirect(
        url_for("login")
    )


# =========================================================
# DASHBOARD
# =========================================================

@app.route("/")
@app.route("/dashboard")
@login_required
def dashboard():

    conn = db_connect()

    total_workers = fetchone(
        conn,
        "SELECT COUNT(*) AS total FROM workers"
    )["total"]

    total_basic = fetchone(
        conn,
        """
        SELECT COALESCE(SUM(basic_salary),0)
        AS total
        FROM workers
        """
    )["total"]

    month = datetime.now().strftime("%Y-%m")

    total_advance = fetchone(
        conn,
        """
        SELECT COALESCE(SUM(amount),0)
        AS total
        FROM worker_advances
        WHERE month_year=?
        """,
        (month,)
    )["total"]

    departments = fetchall(
        conn,
        """
        SELECT department,
               COUNT(*) AS workers,
               COALESCE(SUM(basic_salary),0)
               AS salary
        FROM workers
        GROUP BY department
        ORDER BY department
        """
    )

    conn.close()

    return render_template(
        "dashboard.html",
        total_workers=total_workers,
        total_basic=total_basic,
        total_advance=total_advance,
        departments=departments,
        month=month
    )


# =========================================================
# WORKERS
# =========================================================

@app.route("/workers")
@login_required
def workers():

    search = request.args.get(
        "q",
        ""
    ).strip()

    conn = db_connect()

    if search:

        q = "%" + search.lower() + "%"

        rows = fetchall(
            conn,
            """
            SELECT *
            FROM workers
            WHERE
                CAST(id AS TEXT) LIKE ?
                OR LOWER(COALESCE(name,'')) LIKE ?
                OR LOWER(COALESCE(bangla_name,'')) LIKE ?
                OR LOWER(COALESCE(department,'')) LIKE ?
            ORDER BY id
            """,
            (q, q, q, q)
        )

    else:

        rows = fetchall(
            conn,
            """
            SELECT *
            FROM workers
            ORDER BY id
            """
        )

    conn.close()

    return render_template(
        "workers.html",
        workers=rows,
        q=search
    )


# =========================================================
# ADD WORKER
# =========================================================

@app.route("/workers/add", methods=["GET", "POST"])
@login_required
def add_worker():

    if request.method == "POST":

        name = request.form.get(
            "name",
            ""
        ).strip()

        bangla_name = request.form.get(
            "bangla_name",
            ""
        ).strip()

        department = request.form.get(
            "department",
            "General"
        ).strip()

        designation = request.form.get(
            "designation",
            "Worker"
        ).strip()

        basic_salary = float(
            request.form.get(
                "basic_salary",
                0
            ) or 0
        )

        ot_rate = float(
            request.form.get(
                "ot_rate",
                0
            ) or 0
        )

        refreshment = float(
            request.form.get(
                "refreshment_bill",
                0
            ) or 0
        )

        conn = db_connect()

        execute(conn, """
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
        """, (
            name,
            bangla_name,
            department,
            designation,
            basic_salary,
            ot_rate,
            refreshment
        ))

        conn.commit()
        conn.close()

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
    "/workers/edit/<int:wid>",
    methods=["GET", "POST"]
)
@login_required
def edit_worker(wid):

    conn = db_connect()

    worker = fetchone(
        conn,
        "SELECT * FROM workers WHERE id=?",
        (wid,)
    )

    if not worker:

        conn.close()

        abort(404)

    if request.method == "POST":

        name = request.form.get(
            "name",
            ""
        ).strip()

        bangla_name = request.form.get(
            "bangla_name",
            ""
        ).strip()

        department = request.form.get(
            "department",
            "General"
        ).strip()

        designation = request.form.get(
            "designation",
            "Worker"
        ).strip()

        basic_salary = float(
            request.form.get(
                "basic_salary",
                0
            ) or 0
        )

        ot_rate = float(
            request.form.get(
                "ot_rate",
                0
            ) or 0
        )

        refreshment = float(
            request.form.get(
                "refreshment_bill",
                0
            ) or 0
        )

        execute(conn, """
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
        """, (
            name,
            bangla_name,
            department,
            designation,
            basic_salary,
            ot_rate,
            refreshment,
            wid
        ))

        conn.commit()
        conn.close()

        log_activity(
            "Edited worker ID: "
            + str(wid)
        )

        flash(
            "Worker updated successfully.",
            "success"
        )

        return redirect(
            url_for("workers")
        )

    conn.close()

    return render_template(
        "worker_form.html",
        worker=worker
    )


# =========================================================
# DELETE WORKER
# =========================================================

@app.route(
    "/workers/delete/<int:wid>",
    methods=["POST", "GET"]
)
@admin_required
def delete_worker(wid):

    conn = db_connect()

    execute(
        conn,
        "DELETE FROM worker_advances WHERE worker_id=?",
        (wid,)
    )

    execute(
        conn,
        "DELETE FROM attendance WHERE worker_id=?",
        (wid,)
    )

    execute(
        conn,
        "DELETE FROM daily_attendance WHERE worker_id=?",
        (wid,)
    )

    execute(
        conn,
        "DELETE FROM workers WHERE id=?",
        (wid,)
    )

    conn.commit()
    conn.close()

    log_activity(
        "Deleted worker ID: " + str(wid)
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

        conn = db_connect()

        worker_ids = request.form.getlist(
            "worker_id"
        )

        for wid in worker_ids:

            present = int(
                request.form.get(
                    f"present_{wid}",
                    0
                ) or 0
            )

            absent = int(
                request.form.get(
                    f"absent_{wid}",
                    0
                ) or 0
            )

            ot_hours = float(
                request.form.get(
                    f"ot_{wid}",
                    0
                ) or 0
            )

            advance = float(
                request.form.get(
                    f"advance_{wid}",
                    0
                ) or 0
            )

            if using_postgres():

                execute(conn, """
                    INSERT INTO attendance
                    (
                        worker_id,
                        month_year,
                        present_days,
                        absent_days,
                        ot_hours,
                        advance_deduction
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(worker_id, month_year)
                    DO UPDATE SET
                        present_days=EXCLUDED.present_days,
                        absent_days=EXCLUDED.absent_days,
                        ot_hours=EXCLUDED.ot_hours,
                        advance_deduction=EXCLUDED.advance_deduction
                """, (
                    wid,
                    month,
                    present,
                    absent,
                    ot_hours,
                    advance
                ))

            else:

                execute(conn, """
                    INSERT INTO attendance
                    (
                        worker_id,
                        month_year,
                        present_days,
                        absent_days,
                        ot_hours,
                        advance_deduction
                    )
                    VALUES (?, ?, ?, ?, ?, ?)
                    ON CONFLICT(worker_id, month_year)
                    DO UPDATE SET
                        present_days=excluded.present_days,
                        absent_days=excluded.absent_days,
                        ot_hours=excluded.ot_hours,
                        advance_deduction=excluded.advance_deduction
                """, (
                    wid,
                    month,
                    present,
                    absent,
                    ot_hours,
                    advance
                ))

        conn.commit()
        conn.close()

        log_activity(
            "Updated attendance for "
            + month
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

    conn = db_connect()

    workers_data = fetchall(
        conn,
        """
        SELECT
            w.*,
            COALESCE(a.present_days,0)
            AS present_days,
            COALESCE(a.absent_days,0)
            AS absent_days,
            COALESCE(a.ot_hours,0)
            AS ot_hours,
            COALESCE(a.advance_deduction,0)
            AS advance_deduction
        FROM workers w
        LEFT JOIN attendance a
            ON a.worker_id=w.id
            AND a.month_year=?
        ORDER BY w.id
        """,
        (month,)
    )

    conn.close()

    return render_template(
        "attendance.html",
        data=workers_data,
        workers=workers_data,
        month=month
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

    search = (
        request.args.get("q", "")
        .strip()
    )

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

    # IMPORTANT:
    # report_rows is also supplied because
    # older/newer templates may use this variable.
    return render_template(
        "report.html",
        data=data,
        report_rows=data,
        month=month,
        q=search,
        selected_department=department
    )


# =========================================================
# DEPARTMENT REPORT
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

    conn = db_connect()

    departments = fetchall(
        conn,
        """
        SELECT DISTINCT department
        FROM workers
        WHERE department IS NOT NULL
        AND TRIM(department) <> ''
        ORDER BY department
        """
    )

    conn.close()

    departments = [
        x["department"]
        for x in departments
    ]

    return render_template(
        "department.html",
        data=data,
        month=month,
        departments=departments,
        selected_department=selected
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

    rows = payroll_rows(month)

    worker = None

    for row in rows:

        if int(row["id"]) == int(wid):
            worker = row
            break

    if not worker:
        abort(404)

    return render_template(
        "payslip.html",
        w=worker,
        worker=worker,
        month=month,
        gross=worker["gross"],
        net=worker["net"]
    )


# =========================================================
# ADVANCE SALARY
# =========================================================

@app.route(
    "/advances",
    methods=["GET", "POST"]
)
@login_required
def advances():

    if request.method == "POST":

        worker_id = int(
            request.form.get(
                "worker_id"
            )
        )

        month_year = request.form.get(
            "month_year",
            datetime.now().strftime("%Y-%m")
        )

        advance_date = request.form.get(
            "advance_date",
            datetime.now().strftime("%Y-%m-%d")
        )

        amount = float(
            request.form.get(
                "amount",
                0
            ) or 0
        )

        note = request.form.get(
            "note",
            ""
        ).strip()

        conn = db_connect()

        execute(conn, """
            INSERT INTO worker_advances
            (
                worker_id,
                month_year,
                advance_date,
                amount,
                note
            )
            VALUES (?, ?, ?, ?, ?)
        """, (
            worker_id,
            month_year,
            advance_date,
            amount,
            note
        ))

        conn.commit()
        conn.close()

        log_activity(
            "Added advance salary"
        )

        flash(
            "Advance salary saved.",
            "success"
        )

        return redirect(
            url_for("advances")
        )

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    conn = db_connect()

    workers_data = fetchall(
        conn,
        """
        SELECT id, name, bangla_name
        FROM workers
        ORDER BY id
        """
    )

    advance_rows = fetchall(
        conn,
        """
        SELECT
            wa.*,
            w.name,
            w.bangla_name
        FROM worker_advances wa
        LEFT JOIN workers w
            ON w.id=wa.worker_id
        WHERE wa.month_year=?
        ORDER BY wa.advance_date DESC,
                 wa.id DESC
        """,
        (month,)
    )

    conn.close()

    return render_template(
        "advances.html",
        workers=workers_data,
        advances=advance_rows,
        month=month
    )


# =========================================================
# DELETE ADVANCE
# =========================================================

@app.route(
    "/advances/delete/<int:aid>",
    methods=["POST", "GET"]
)
@admin_required
def delete_advance(aid):

    conn = db_connect()

    execute(
        conn,
        "DELETE FROM worker_advances WHERE id=?",
        (aid,)
    )

    conn.commit()
    conn.close()

    log_activity(
        "Deleted advance ID: "
        + str(aid)
    )

    flash(
        "Advance deleted.",
        "success"
    )

    return redirect(
        url_for("advances")
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

    current = get_company_settings()

    if request.method == "POST":

        values = {
            "company_name": request.form.get(
                "company_name",
                DEFAULT_COMPANY_NAME
            ).strip(),

            "company_address": request.form.get(
                "company_address",
                ""
            ).strip(),

            "company_phone": request.form.get(
                "company_phone",
                ""
            ).strip(),

            "company_email": request.form.get(
                "company_email",
                ""
            ).strip(),

            "company_logo": request.form.get(
                "company_logo",
                ""
            ).strip()
        }

        conn = db_connect()

        for key, value in values.items():

            existing = fetchone(
                conn,
                "SELECT key FROM company_settings WHERE key=?",
                (key,)
            )

            if existing:

                execute(conn, """
                    UPDATE company_settings
                    SET value=?
                    WHERE key=?
                """, (
                    value,
                    key
                ))

            else:

                execute(conn, """
                    INSERT INTO company_settings
                    (key, value)
                    VALUES (?, ?)
                """, (
                    key,
                    value
                ))

        conn.commit()
        conn.close()

        log_activity(
            "Updated company settings"
        )

        flash(
            "Settings saved successfully.",
            "success"
        )

        current = get_company_settings()

    return render_template(
        "settings.html",
        settings=current
    )


# =========================================================
# USERS
# =========================================================

@app.route(
    "/users",
    methods=["GET", "POST"]
)
@admin_required
def users():

    if request.method == "POST":

        username = request.form.get(
            "username",
            ""
        ).strip()

        password = request.form.get(
            "password",
            ""
        )

        full_name = request.form.get(
            "full_name",
            ""
        ).strip()

        role = request.form.get(
            "role",
            "User"
        )

        if not username or not password:

            flash(
                "Username and password are required.",
                "danger"
            )

            return redirect(
                url_for("users")
            )

        conn = db_connect()

        try:

            execute(conn, """
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
            """, (
                username,
                hash_password(password),
                full_name,
                role,
                1,
                datetime.now().strftime(
                    "%Y-%m-%d %H:%M:%S"
                )
            ))

            conn.commit()

            flash(
                "User created successfully.",
                "success"
            )

        except Exception:

            conn.rollback()

            flash(
                "Username already exists.",
                "danger"
            )

        conn.close()

        return redirect(
            url_for("users")
        )

    conn = db_connect()

    user_rows = fetchall(
        conn,
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

    conn.close()

    return render_template(
        "users.html",
        users=user_rows
    )


# =========================================================
# TOGGLE USER
# =========================================================

@app.route(
    "/users/toggle/<int:uid>",
    methods=["POST", "GET"]
)
@admin_required
def toggle_user(uid):

    conn = db_connect()

    user = fetchone(
        conn,
        "SELECT active FROM users WHERE id=?",
        (uid,)
    )

    if user:

        new_status = (
            0
            if int(user["active"] or 0) == 1
            else 1
        )

        execute(
            conn,
            "UPDATE users SET active=? WHERE id=?",
            (new_status, uid)
        )

        conn.commit()

    conn.close()

    return redirect(
        url_for("users")
    )


# =========================================================
# CHANGE PASSWORD
# =========================================================

@app.route(
    "/change-password",
    methods=["GET", "POST"]
)
@login_required
def change_password():

    if request.method == "POST":

        old_password = request.form.get(
            "old_password",
            ""
        )

        new_password = request.form.get(
            "new_password",
            ""
        )

        confirm_password = request.form.get(
            "confirm_password",
            ""
        )

        if new_password != confirm_password:

            flash(
                "New passwords do not match.",
                "danger"
            )

            return redirect(
                url_for("change_password")
            )

        conn = db_connect()

        user = fetchone(
            conn,
            """
            SELECT *
            FROM users
            WHERE id=?
            """,
            (session["user_id"],)
        )

        if not user or user.get("password_hash") != hash_password(
            old_password
        ):

            conn.close()

            flash(
                "Current password is incorrect.",
                "danger"
            )

            return redirect(
                url_for("change_password")
            )

        execute(
            conn,
            """
            UPDATE users
            SET password_hash=?
            WHERE id=?
            """,
            (
                hash_password(new_password),
                session["user_id"]
            )
        )

        conn.commit()
        conn.close()

        flash(
            "Password changed successfully.",
            "success"
        )

        return redirect(
            url_for("dashboard")
        )

    return render_template(
        "change_password.html"
    )


# =========================================================
# ACTIVITY LOG
# =========================================================

@app.route("/activity")
@admin_required
def activity():

    conn = db_connect()

    logs = fetchall(
        conn,
        """
        SELECT *
        FROM activity_log
        ORDER BY id DESC
        LIMIT 500
        """
    )

    conn.close()

    return render_template(
        "activity.html",
        logs=logs
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
        "q",
        ""
    ).strip()

    department = request.args.get(
        "department",
        "All Departments"
    )

    data = payroll_rows(
        month,
        search,
        department
    )

    output = io.StringIO()

    writer = csv.writer(output)

    writer.writerow([
        "ID",
        "Name",
        "Department",
        "Basic Salary",
        "Present Days",
        "Absent Days",
        "OT Hours",
        "OT Amount",
        "Refreshment",
        "Advance",
        "Gross",
        "Net Salary"
    ])

    for row in data:

        writer.writerow([
            row["id"],
            row.get("name", ""),
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

    mem = io.BytesIO(
        output.getvalue().encode("utf-8-sig")
    )

    mem.seek(0)

    return send_file(
        mem,
        mimetype="text/csv",
        as_attachment=True,
        download_name=f"payroll_{month}.csv"
    )


# =========================================================
# EXCEL EXPORT
# =========================================================

@app.route("/report/excel")
@login_required
def report_excel():

    if Workbook is None:

        flash(
            "openpyxl is not installed.",
            "danger"
        )

        return redirect(
            url_for("report")
        )

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q",
        ""
    ).strip()

    department = request.args.get(
        "department",
        "All Departments"
    )

    data = payroll_rows(
        month,
        search,
        department
    )

    wb = Workbook()

    ws = wb.active

    ws.title = "Payroll"

    headers = [
        "ID",
        "Name",
        "Department",
        "Basic Salary",
        "Present Days",
        "Absent Days",
        "OT Hours",
        "OT Amount",
        "Refreshment",
        "Advance",
        "Gross",
        "Net Salary"
    ]

    ws.append(headers)

    for row in data:

        ws.append([
            row["id"],
            row.get("name", ""),
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

    for cell in ws[1]:

        cell.font = cell.font.copy(
            bold=True
        )

    output = io.BytesIO()

    wb.save(output)

    output.seek(0)

    return send_file(
        output,
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        as_attachment=True,
        download_name=f"payroll_{month}.xlsx"
    )


# =========================================================
# PDF EXPORT
# =========================================================

@app.route("/report/pdf")
@login_required
def report_pdf():

    if SimpleDocTemplate is None:

        flash(
            "reportlab is not installed.",
            "danger"
        )

        return redirect(
            url_for("report")
        )

    month = (
        request.args.get("month")
        or datetime.now().strftime("%Y-%m")
    )

    search = request.args.get(
        "q",
        ""
    ).strip()

    department = request.args.get(
        "department",
        "All Departments"
    )

    data = payroll_rows(
        month,
        search,
        department
    )

    settings = get_company_settings()

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
            settings.get(
                "company_name",
                DEFAULT_COMPANY_NAME
            ),
            styles["Title"]
        )
    )

    story.append(
        Paragraph(
            f"Payroll Report - {month}",
            styles["Heading2"]
        )
    )

    story.append(Spacer(1, 10))

    table_data = [[
        "ID",
        "Name",
        "Department",
        "Basic",
        "Present",
        "Absent",
        "OT",
        "OT Amount",
        "Nasta",
        "Advance",
        "Gross",
        "Net"
    ]]

    for row in data:

        table_data.append([
            row["id"],
            row.get("name", ""),
            row.get("department", ""),
            f'{float(row.get("basic_salary", 0) or 0):,.2f}',
            row.get("present_days", 0),
            row.get("absent_days", 0),
            row.get("ot_hours", 0),
            f'{float(row.get("ot_amount", 0) or 0):,.2f}',
            f'{float(row.get("refreshment", 0) or 0):,.2f}',
            f'{float(row.get("advance", 0) or 0):,.2f}',
            f'{float(row.get("gross", 0) or 0):,.2f}',
            f'{float(row.get("net_salary", 0) or 0):,.2f}'
        ])

    table = Table(
        table_data,
        repeatRows=1
    )

    table.setStyle(
        TableStyle([
            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),
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
                (-1, 0),
                "Helvetica-Bold"
            ),
            (
                "FONTSIZE",
                (0, 0),
                (-1, -1),
                7
            ),
            (
                "ALIGN",
                (3, 1),
                (-1, -1),
                "RIGHT"
            )
        ])
    )

    story.append(table)

    doc.build(story)

    output.seek(0)

    return send_file(
        output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"payroll_{month}.pdf"
    )


# =========================================================
# ERROR HANDLERS
# =========================================================

@app.errorhandler(404)
def page_not_found(error):

    return (
        render_template(
            "base.html",
            error_message="Page not found."
        ),
        404
    )


# =========================================================
# STARTUP
# =========================================================

init_db()


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
