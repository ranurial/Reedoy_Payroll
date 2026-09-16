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

app = Flask(**name**)

app.secret_key = os.environ.get(
"SECRET_KEY",
"REEDOY-CHANGE-THIS-SECRET"
)

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

# Render sometimes provides postgres://

# psycopg2 expects postgresql://

if DATABASE_URL.startswith("postgres://"):
DATABASE_URL = DATABASE_URL.replace(
"postgres://",
"postgresql://",
1
)

SQLITE_PATH = os.environ.get(
"SQLITE_PATH",
os.path.join(
os.path.dirname(**file**),
"reedoy_payroll.db"
)
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

def is_postgres():
return bool(DATABASE_URL)

def get_db():
"""
Render:
PostgreSQL

```
Local:
    SQLite
"""

if is_postgres():

    try:
        import psycopg2
        from psycopg2.extras import RealDictCursor

        conn = psycopg2.connect(
            DATABASE_URL,
            connect_timeout=10
        )

        return conn

    except ImportError as exc:
        raise RuntimeError(
            "psycopg2 is not installed. "
            "Add psycopg2-binary to requirements.txt"
        ) from exc

conn = sqlite3.connect(
    SQLITE_PATH,
    timeout=30
)

conn.row_factory = sqlite3.Row

return conn
```

def close_db(conn):
try:
conn.close()
except Exception:
pass

# =========================================================

# DATABASE HELPERS

# =========================================================

def fetch_all(sql, params=()):

```
conn = get_db()

try:

    if is_postgres():

        from psycopg2.extras import RealDictCursor

        cur = conn.cursor(
            cursor_factory=RealDictCursor
        )

        cur.execute(sql, params)

        return cur.fetchall()

    cur = conn.cursor()

    cur.execute(
        sql,
        params
    )

    return cur.fetchall()

finally:

    close_db(conn)
```

def fetch_one(sql, params=()):

```
conn = get_db()

try:

    if is_postgres():

        from psycopg2.extras import RealDictCursor

        cur = conn.cursor(
            cursor_factory=RealDictCursor
        )

        cur.execute(
            sql,
            params
        )

        return cur.fetchone()

    cur = conn.cursor()

    cur.execute(
        sql,
        params
    )

    return cur.fetchone()

finally:

    close_db(conn)
```

def execute_sql(sql, params=()):

```
conn = get_db()

try:

    cur = conn.cursor()

    cur.execute(
        sql,
        params
    )

    conn.commit()

    try:
        return cur.lastrowid
    except Exception:
        return None

except Exception:

    conn.rollback()
    raise

finally:

    close_db(conn)
```

# =========================================================

# ROW HELPERS

# =========================================================

def row_to_dict(row):

```
if row is None:
    return {}

if isinstance(row, dict):
    return dict(row)

return dict(row)
```

def row_value(row, key, default=None):

```
if row is None:
    return default

if isinstance(row, dict):
    return row.get(
        key,
        default
    )

try:
    return row[key]
except Exception:
    return default
```

def safe_float(value, default=0):

```
try:

    if value is None:
        return default

    return float(value)

except (ValueError, TypeError):

    return default
```

# =========================================================

# PASSWORD

# =========================================================

def hash_password(password):

```
return hashlib.sha256(
    str(password).encode("utf-8")
).hexdigest()
```

def check_password(
password,
password_hash
):

```
return (
    hash_password(password)
    == str(password_hash)
)
```

# =========================================================

# DATABASE INITIALIZATION

# =========================================================

def init_db():

```
conn = get_db()

try:

    cur = conn.cursor()

    if is_postgres():

        # -------------------------------------------------
        # USERS
        # -------------------------------------------------

        cur.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT DEFAULT 'User'
            )
        """)

        # -------------------------------------------------
        # COMPANY SETTINGS
        # -------------------------------------------------

        cur.execute("""
            CREATE TABLE IF NOT EXISTS company_settings (
                id SERIAL PRIMARY KEY,
                setting_key TEXT UNIQUE NOT NULL,
                setting_value TEXT
            )
        """)

        # -------------------------------------------------
        # WORKERS
        # -------------------------------------------------

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

        # -------------------------------------------------
        # MONTHLY ATTENDANCE
        # -------------------------------------------------

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

        # -------------------------------------------------
        # ADVANCE
        # -------------------------------------------------

        cur.execute("""
            CREATE TABLE IF NOT EXISTS worker_advances (
                id SERIAL PRIMARY KEY,
                worker_id INTEGER,
                advance_date TEXT,
                amount DOUBLE PRECISION DEFAULT 0,
                note TEXT
            )
        """)

        # -------------------------------------------------
        # ACTIVITY
        # -------------------------------------------------

        cur.execute("""
            CREATE TABLE IF NOT EXISTS activity_log (
                id SERIAL PRIMARY KEY,
                username TEXT,
                action TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

    else:

        # -------------------------------------------------
        # SQLITE
        # -------------------------------------------------

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

        cur.execute("""
            SELECT id
            FROM users
            WHERE username=%s
        """, ("admin",))

    else:

        cur.execute("""
            SELECT id
            FROM users
            WHERE username=?
        """, ("admin",))

    admin = cur.fetchone()

    if not admin:

        if is_postgres():

            cur.execute("""
                INSERT INTO users
                (
                    username,
                    password_hash,
                    role
                )
                VALUES (%s,%s,%s)
            """, (
                "admin",
                hash_password("admin123"),
                "Administrator"
            ))

        else:

            cur.execute("""
                INSERT INTO users
                (
                    username,
                    password_hash,
                    role
                )
                VALUES (?,?,?)
            """, (
                "admin",
                hash_password("admin123"),
                "Administrator"
            ))

    # -------------------------------------------------
    # DEFAULT SETTINGS
    # -------------------------------------------------

    defaults = {
        "company_name":
            "Reedoy Textile Dyeing Printing & Finishing",
        "address": "",
        "phone": "",
        "language": "English"
    }

    for key, value in defaults.items():

        if is_postgres():

            cur.execute("""
                SELECT id
                FROM company_settings
                WHERE setting_key=%s
            """, (key,))

        else:

            cur.execute("""
                SELECT id
                FROM company_settings
                WHERE setting_key=?
            """, (key,))

        exists = cur.fetchone()

        if not exists:

            if is_postgres():

                cur.execute("""
                    INSERT INTO company_settings
                    (
                        setting_key,
                        setting_value
                    )
                    VALUES (%s,%s)
                """, (
                    key,
                    value
                ))

            else:

                cur.execute("""
                    INSERT INTO company_settings
                    (
                        setting_key,
                        setting_value
                    )
                    VALUES (?,?)
                """, (
                    key,
                    value
                ))

    conn.commit()

finally:

    close_db(conn)
```

# =========================================================

# LOGIN

# =========================================================

def login_required(f):

```
@wraps(f)
def decorated(*args, **kwargs):

    if not session.get("user_id"):

        return redirect(
            url_for("login")
        )

    return f(
        *args,
        **kwargs
    )

return decorated
```

def admin_required(f):

```
@wraps(f)
def decorated(*args, **kwargs):

    if not session.get("user_id"):

        return redirect(
            url_for("login")
        )

    if session.get("role") != "Administrator":

        flash(
            "Administrator access required.",
            "danger"
        )

        return redirect(
            url_for("dashboard")
        )

    return f(
        *args,
        **kwargs
    )

return decorated
```

# =========================================================

# ACTIVITY LOG

# =========================================================

def log_activity(action):

```
try:

    username = session.get(
        "username",
        "system"
    )

    if is_postgres():

        execute_sql("""
            INSERT INTO activity_log
            (
                username,
                action
            )
            VALUES (%s,%s)
        """, (
            username,
            action
        ))

    else:

        execute_sql("""
            INSERT INTO activity_log
            (
                username,
                action
            )
            VALUES (?,?)
        """, (
            username,
            action
        ))

except Exception as exc:

    print(
        "Activity log error:",
        exc
    )
```

# =========================================================

# SETTINGS

# =========================================================

def get_settings():

```
rows = fetch_all("""
    SELECT setting_key, setting_value
    FROM company_settings
""")

result = {}

for row in rows:

    result[
        row_value(row, "setting_key")
    ] = row_value(
        row,
        "setting_value",
        ""
    )

return result
```

# =========================================================

# TEMPLATE CONTEXT

# =========================================================

@app.context_processor
def inject_globals():

```
settings = {}

try:
    settings = get_settings()
except Exception:
    settings = {}

language = (
    settings.get(
        "language",
        "English"
    )
    if settings
    else "English"
)

return {
    "BN": BN,
    "DEPT_BN": DEPT_BN,
    "APP_LANGUAGE": language,
    "company_settings": settings
}
```

# =========================================================

# SALARY CALCULATION

# =========================================================

def calculate_salary(
worker,
attendance=None,
advance=0
):

```
attendance = attendance or {}

basic = safe_float(
    row_value(
        worker,
        "basic_salary",
        0
    )
)

present = safe_float(
    row_value(
        attendance,
        "present_days",
        0
    )
)

absent = safe_float(
    row_value(
        attendance,
        "absent_days",
        0
    )
)

ot_hours = safe_float(
    row_value(
        attendance,
        "ot_hours",
        0
    )
)

ot_rate = safe_float(
    row_value(
        worker,
        "ot_rate",
        0
    )
)

refreshment = safe_float(
    row_value(
        worker,
        "refreshment_bill",
        0
    )
)

advance = safe_float(
    advance
)

# 30-day salary calculation
daily_basic = basic / 30.0

earned_basic = (
    daily_basic * present
)

ot_amount = (
    ot_hours * ot_rate
)

gross = (
    earned_basic
    + ot_amount
    + refreshment
)

net = (
    gross - advance
)

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
```

# =========================================================

# DEPARTMENTS

# =========================================================

def get_departments():

```
rows = fetch_all("""
    SELECT DISTINCT department
    FROM workers
    WHERE department IS NOT NULL
      AND TRIM(department) <> ''
    ORDER BY department
""")

departments = [
    "All Departments"
]

for row in rows:

    dep = row_value(
        row,
        "department"
    )

    if dep and dep not in departments:

        departments.append(dep)

return departments
```

# =========================================================

# PAYROLL DATA

# =========================================================

def payroll_rows(
month,
search="",
department="All Departments"
):

```
search = (
    search or ""
).strip()

department = (
    department
    or "All Departments"
)

# -----------------------------------------------------
# WORKERS + MONTHLY ATTENDANCE
# -----------------------------------------------------

if is_postgres():

    sql = """
        SELECT
            w.*,
            COALESCE(a.present_days, 0)
                AS present_days,
            COALESCE(a.absent_days, 0)
                AS absent_days,
            COALESCE(a.ot_hours, 0)
                AS ot_hours
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
            COALESCE(a.present_days, 0)
                AS present_days,
            COALESCE(a.absent_days, 0)
                AS absent_days,
            COALESCE(a.ot_hours, 0)
                AS ot_hours
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

data = []

for row in rows:

    worker = row_to_dict(row)

    worker_id = row_value(
        worker,
        "id"
    )

    attendance = {
        "present_days":
            row_value(
                worker,
                "present_days",
                0
            ),

        "absent_days":
            row_value(
                worker,
                "absent_days",
                0
            ),

        "ot_hours":
            row_value(
                worker,
                "ot_hours",
                0
            )
    }

    # -------------------------------------------------
    # ADVANCE FOR MONTH
    # -------------------------------------------------

    advance = 0

    if worker_id:

        if is_postgres():

            adv_row = fetch_one("""
                SELECT
                    COALESCE(
                        SUM(amount),
                        0
                    ) AS total
                FROM worker_advances
                WHERE worker_id=%s
                  AND (
                      advance_date IS NULL
                      OR LEFT(
                          advance_date,
                          7
                      )=%s
                  )
            """, (
                worker_id,
                month
            ))

        else:

            adv_row = fetch_one("""
                SELECT
                    COALESCE(
                        SUM(amount),
                        0
                    ) AS total
                FROM worker_advances
                WHERE worker_id=?
                  AND (
                      advance_date IS NULL
                      OR substr(
                          advance_date,
                          1,
                          7
                      )=?
                  )
            """, (
                worker_id,
                month
            ))

        advance = safe_float(
            row_value(
                adv_row,
                "total",
                0
            )
        )

    salary = calculate_salary(
        worker,
        attendance,
        advance
    )

    worker.update(salary)

    data.append(worker)

return data
```

# =========================================================

# HOME

# =========================================================

@app.route("/")
def home():

```
if session.get("user_id"):

    return redirect(
        url_for("dashboard")
    )

return redirect(
    url_for("login")
)
```

# =========================================================

# LOGIN

# =========================================================

@app.route(
"/login",
methods=["GET", "POST"]
)
def login():

```
if request.method == "POST":

    username = request.form.get(
        "username",
        ""
    ).strip()

    password = request.form.get(
        "password",
        ""
    )

    if not username:

        flash(
            "Username is required.",
            "danger"
        )

        return render_template(
            "login.html"
        )

    if is_postgres():

        user = fetch_one("""
            SELECT *
            FROM users
            WHERE username=%s
        """, (
            username,
        ))

    else:

        user = fetch_one("""
            SELECT *
            FROM users
            WHERE username=?
        """, (
            username,
        ))

    if user:

        password_hash = row_value(
            user,
            "password_hash",
            ""
        )

        if check_password(
            password,
            password_hash
        ):

            session["user_id"] = row_value(
                user,
                "id"
            )

            session["username"] = username

            session["role"] = row_value(
                user,
                "role",
                "User"
            )

            log_activity(
                "Login"
            )

            return redirect(
                url_for("dashboard")
            )

    flash(
        "Invalid username or password.",
        "danger"
    )

return render_template(
    "login.html"
)
```

# =========================================================

# LOGOUT

# =========================================================

@app.route("/logout")
def logout():

```
if session.get("user_id"):

    log_activity(
        "Logout"
    )

session.clear()

return redirect(
    url_for("login")
)
```

# =========================================================

# DASHBOARD

# =========================================================

@app.route("/dashboard")
@login_required
def dashboard():

```
worker_row = fetch_one("""
    SELECT COUNT(*) AS total
    FROM workers
""")

total_workers = int(
    row_value(
        worker_row,
        "total",
        0
    ) or 0
)

return render_template(
    "dashboard.html",
    total_workers=total_workers,
    settings=get_settings()
)
```

# =========================================================

# WORKERS

# =========================================================

@app.route("/workers")
@login_required
def workers():

```
rows = fetch_all("""
    SELECT *
    FROM workers
    ORDER BY id
""")

workers_data = [
    row_to_dict(row)
    for row in rows
]

return render_template(
    "workers.html",
    workers=workers_data
)
```

# =========================================================

# ADD WORKER

# =========================================================

@app.route(
"/workers/add",
methods=["GET", "POST"]
)
@login_required
def add_worker():

```
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
        ""
    ).strip()

    designation = request.form.get(
        "designation",
        ""
    ).strip()

    basic_salary = safe_float(
        request.form.get(
            "basic_salary",
            0
        )
    )

    ot_rate = safe_float(
        request.form.get(
            "ot_rate",
            0
        )
    )

    refreshment_bill = safe_float(
        request.form.get(
            "refreshment_bill",
            0
        )
    )

    phone = request.form.get(
        "phone",
        ""
    ).strip()

    address = request.form.get(
        "address",
        ""
    ).strip()

    joining_date = request.form.get(
        "joining_date",
        ""
    ).strip()

    status = request.form.get(
        "status",
        "Active"
    )

    if not name:

        flash(
            "Worker name is required.",
            "danger"
        )

        return render_template(
            "worker_form.html",
            worker=None
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
                %s,%s,%s,%s,%s,%s,
                %s,%s,%s,%s,%s
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
```

# =========================================================

# EDIT WORKER

# =========================================================

@app.route(
"/workers/edit/[int:worker_id](int:worker_id)",
methods=["GET", "POST"]
)
@login_required
def edit_worker(worker_id):

```
if is_postgres():

    worker = fetch_one("""
        SELECT *
        FROM workers
        WHERE id=%s
    """, (
        worker_id,
    ))

else:

    worker = fetch_one("""
        SELECT *
        FROM workers
        WHERE id=?
    """, (
        worker_id,
    ))

if not worker:

    abort(404)

worker = row_to_dict(worker)

if request.method == "POST":

    values = (
        request.form.get(
            "name",
            ""
        ).strip(),

        request.form.get(
            "bangla_name",
            ""
        ).strip(),

        request.form.get(
            "department",
            ""
        ).strip(),

        request.form.get(
            "designation",
            ""
        ).strip(),

        safe_float(
            request.form.get(
                "basic_salary",
                0
            )
        ),

        safe_float(
            request.form.get(
                "ot_rate",
                0
            )
        ),

        safe_float(
            request.form.get(
                "refreshment_bill",
                0
            )
        ),

        request.form.get(
            "phone",
            ""
        ).strip(),

        request.form.get(
            "address",
            ""
        ).strip(),

        request.form.get(
            "joining_date",
            ""
        ).strip(),

        request.form.get(
            "status",
            "Active"
        ),

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
```

# =========================================================

# DELETE WORKER

# =========================================================

@app.route(
"/workers/delete/[int:worker_id](int:worker_id)",
methods=["POST"]
)
@admin_required
def delete_worker(worker_id):

```
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
```

# =========================================================

# ATTENDANCE

# =========================================================

@app.route("/attendance")
@login_required
def attendance():

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
)

data = payroll_rows(
    month=month
)

return render_template(
    "attendance.html",
    data=data,
    month=month
)
```

# =========================================================

# SAVE ATTENDANCE

# =========================================================

@app.route(
"/attendance/save",
methods=["POST"]
)
@login_required
def save_attendance():

```
month = (
    request.form.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
)

worker_ids = request.form.getlist(
    "worker_id"
)

for worker_id in worker_ids:

    present = safe_float(
        request.form.get(
            f"present_{worker_id}",
            0
        )
    )

    absent = safe_float(
        request.form.get(
            f"absent_{worker_id}",
            0
        )
    )

    ot_hours = safe_float(
        request.form.get(
            f"ot_{worker_id}",
            0
        )
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

    attendance_id = row_value(
        exists,
        "id"
    )

    if attendance_id:

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
```

# =========================================================

# REPORT

# =========================================================

@app.route("/report")
@login_required
def report():

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
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
```

# =========================================================

# DEPARTMENT SALARY

# =========================================================

@app.route("/department")
@login_required
def department():

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
)

selected = (
    request.args.get(
        "department"
    )
    or "All Departments"
)

data = payroll_rows(
    month=month,
    department=selected
)

departments = get_departments()

return render_template(
    "department.html",
    data=data,
    report_rows=data,
    rows=data,
    workers=data,
    month=month,
    selected=selected,
    selected_department=selected,
    department=selected,
    departments=departments
)
```

# =========================================================

# PAYSLIP

# =========================================================

@app.route(
"/payslip/[int:worker_id](int:worker_id)"
)
@login_required
def payslip(worker_id):

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
)

data = payroll_rows(
    month=month
)

worker = None

for row in data:

    if int(
        row_value(
            row,
            "id",
            0
        )
    ) == worker_id:

        worker = row
        break

if not worker:

    abort(404)

return render_template(
    "payslip.html",
    worker=worker,
    data=worker,
    month=month,
    settings=get_settings()
)
```

# =========================================================

# ADVANCE

# =========================================================

@app.route("/advance")
@login_required
def advance():

```
workers_rows = fetch_all("""
    SELECT
        id,
        name,
        bangla_name
    FROM workers
    ORDER BY id
""")

workers_data = [
    row_to_dict(row)
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
    row_to_dict(row)
    for row in advances
]

return render_template(
    "advance.html",
    workers=workers_data,
    advances=advances_data
)
```

# =========================================================

# SAVE ADVANCE

# =========================================================

@app.route(
"/advance/save",
methods=["POST"]
)
@login_required
def save_advance():

```
worker_id = request.form.get(
    "worker_id"
)

amount = safe_float(
    request.form.get(
        "amount",
        0
    )
)

advance_date = (
    request.form.get(
        "advance_date"
    )
    or datetime.now().strftime(
        "%Y-%m-%d"
    )
)

note = request.form.get(
    "note",
    ""
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
```

# =========================================================

# SETTINGS

# =========================================================

@app.route(
"/settings",
methods=["GET", "POST"]
)
@admin_required
def settings():

```
if request.method == "POST":

    values = {
        "company_name":
            request.form.get(
                "company_name",
                ""
            ).strip(),

        "address":
            request.form.get(
                "address",
                ""
            ).strip(),

        "phone":
            request.form.get(
                "phone",
                ""
            ).strip(),

        "language":
            request.form.get(
                "language",
                "English"
            )
    }

    for key, value in values.items():

        if is_postgres():

            execute_sql("""
                INSERT INTO company_settings
                (
                    setting_key,
                    setting_value
                )
                VALUES (%s,%s)
                ON CONFLICT(setting_key)
                DO UPDATE SET
                    setting_value=
                        EXCLUDED.setting_value
            """, (
                key,
                value
            ))

        else:

            exists = fetch_one("""
                SELECT id
                FROM company_settings
                WHERE setting_key=?
            """, (
                key,
            ))

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
                    (
                        setting_key,
                        setting_value
                    )
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
```

# =========================================================

# USERS

# =========================================================

@app.route("/users")
@admin_required
def users():

```
rows = fetch_all("""
    SELECT
        id,
        username,
        role
    FROM users
    ORDER BY id
""")

users_data = [
    row_to_dict(row)
    for row in rows
]

return render_template(
    "users.html",
    users=users_data
)
```

# =========================================================

# ADD USER

# =========================================================

@app.route(
"/users/add",
methods=["POST"]
)
@admin_required
def add_user():

```
username = request.form.get(
    "username",
    ""
).strip()

password = request.form.get(
    "password",
    ""
)

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

except Exception as exc:

    print(
        "Add user error:",
        exc
    )

    flash(
        "Username already exists.",
        "danger"
    )

return redirect(
    url_for("users")
)
```

# =========================================================

# DELETE USER

# =========================================================

@app.route(
"/users/delete/[int:user_id](int:user_id)",
methods=["POST"]
)
@admin_required
def delete_user(user_id):

```
if user_id == session.get(
    "user_id"
):

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
```

# =========================================================

# ACTIVITY

# =========================================================

@app.route("/activity")
@admin_required
def activity():

```
rows = fetch_all("""
    SELECT *
    FROM activity_log
    ORDER BY id DESC
    LIMIT 500
""")

logs = [
    row_to_dict(row)
    for row in rows
]

return render_template(
    "activity.html",
    logs=logs,
    activity=logs
)
```

# =========================================================

# CSV EXPORT

# =========================================================

@app.route("/report/csv")
@login_required
def report_csv():

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
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
        row_value(row, "id", ""),
        row_value(row, "name", ""),
        row_value(row, "bangla_name", ""),
        row_value(row, "department", ""),
        row_value(row, "basic_salary", 0),
        row_value(row, "present_days", 0),
        row_value(row, "absent_days", 0),
        row_value(row, "ot_hours", 0),
        row_value(row, "ot_amount", 0),
        row_value(row, "refreshment", 0),
        row_value(row, "advance", 0),
        row_value(row, "gross", 0),
        row_value(row, "net_salary", 0)
    ])

output.seek(0)

filename = (
    f"Reedoy_Payroll_{month}.csv"
)

return send_file(
    io.BytesIO(
        output.getvalue().encode(
            "utf-8-sig"
        )
    ),
    mimetype="text/csv",
    as_attachment=True,
    download_name=filename
)
```

# =========================================================

# EXCEL EXPORT

# =========================================================

@app.route("/report/excel")
@login_required
def report_excel():

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
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
        row_value(row, "id", ""),
        row_value(row, "name", ""),
        row_value(row, "bangla_name", ""),
        row_value(row, "department", ""),
        row_value(row, "basic_salary", 0),
        row_value(row, "present_days", 0),
        row_value(row, "absent_days", 0),
        row_value(row, "ot_hours", 0),
        row_value(row, "ot_amount", 0),
        row_value(row, "refreshment", 0),
        row_value(row, "advance", 0),
        row_value(row, "gross", 0),
        row_value(row, "net_salary", 0)
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
```

# =========================================================

# PDF EXPORT

# =========================================================

@app.route("/report/pdf")
@login_required
def report_pdf():

```
month = (
    request.args.get(
        "month"
    )
    or datetime.now().strftime(
        "%Y-%m"
    )
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
    month=month,
    search=search,
    department=department
)

try:

    from reportlab.lib import colors
    from reportlab.lib.pagesizes import (
        A4,
        landscape
    )
    from reportlab.lib.styles import (
        getSampleStyleSheet
    )
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
        row_value(row, "id", ""),
        row_value(row, "name", ""),
        row_value(row, "department", ""),

        f"{safe_float(
            row_value(
                row,
                "basic_salary",
                0
            )
        ):.2f}",

        f"{safe_float(
            row_value(
                row,
                "present_days",
                0
            )
        ):.1f}",

        f"{safe_float(
            row_value(
                row,
                "absent_days",
                0
            )
        ):.1f}",

        f"{safe_float(
            row_value(
                row,
                "ot_hours",
                0
            )
        ):.1f}",

        f"{safe_float(
            row_value(
                row,
                "ot_amount",
                0
            )
        ):.2f}",

        f"{safe_float(
            row_value(
                row,
                "refreshment",
                0
            )
        ):.2f}",

        f"{safe_float(
            row_value(
                row,
                "advance",
                0
            )
        ):.2f}",

        f"{safe_float(
            row_value(
                row,
                "net_salary",
                0
            )
        ):.2f}"
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
            colors.grey
        ),
        (
            "TEXTCOLOR",
            (0, 0),
            (-1, 0),
            colors.white
        ),
        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.5,
            colors.black
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

filename = (
    f"Reedoy_Payroll_{month}.pdf"
)

return send_file(
    output,
    mimetype="application/pdf",
    as_attachment=True,
    download_name=filename
)
```

# =========================================================

# ERROR HANDLERS

# =========================================================

@app.errorhandler(404)
def page_not_found(error):

```
return render_template(
    "404.html"
), 404
```

@app.errorhandler(500)
def internal_error(error):

```
return render_template(
    "500.html"
), 500
```

# =========================================================

# STARTUP

# =========================================================

try:

```
init_db()

print(
    "Reedoy Payroll database initialization OK."
)
```

except Exception as exc:

```
print(
    "Database initialization error:",
    repr(exc)
)
```

# =========================================================

# LOCAL RUN

# =========================================================

if **name** == "**main**":

```
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
