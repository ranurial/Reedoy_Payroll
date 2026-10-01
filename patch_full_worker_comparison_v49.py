from pathlib import Path
import sqlite3
import hashlib
import json
import shutil
import datetime
import py_compile

BASE = Path(".")
APP = BASE / "app.py"
DB = BASE / "factory_payroll.db"
SEED = BASE / "workers_compare_seed.json"

FIELDS = [
    "name",
    "bangla_name",
    "department",
    "designation",
    "basic_salary",
    "ot_rate",
    "refreshment_bill",
    "phone",
    "address",
    "joining_date",
    "status",
]

def canonical(value):
    if value is None:
        return ""
    return str(value).strip()

def field_hash(value):
    return hashlib.sha256(
        canonical(value).encode("utf-8")
    ).hexdigest()

# ------------------------------------------------------------
# 1. Create LOCAL comparison seed
# ------------------------------------------------------------
conn = sqlite3.connect(str(DB))

rows = conn.execute("""
    SELECT id,
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
    FROM workers
    ORDER BY id
""").fetchall()

conn.close()

if len(rows) != 169:
    raise SystemExit(
        f"[ERROR] Expected 169 workers, found {len(rows)}"
    )

seed = []

for row in rows:
    worker_id = int(row[0])

    values = {
        field: row[i + 1]
        for i, field in enumerate(FIELDS)
    }

    seed.append({
        "id": worker_id,
        "hashes": {
            field: field_hash(values[field])
            for field in FIELDS
        }
    })

SEED.write_text(
    json.dumps(seed, ensure_ascii=False, indent=2),
    encoding="utf-8"
)

print("[OK] Local comparison seed created")
print("File:", SEED)
print("Workers:", len(seed))
print("Fields compared:", len(FIELDS))

# ------------------------------------------------------------
# 2. Patch app.py
# ------------------------------------------------------------
text = APP.read_text(encoding="utf-8-sig")

route_path = "/sync/compare-online-workers"

if route_path in text:
    print("[INFO] Comparison route already exists.")
    raise SystemExit(0)

stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

BACKUP = BASE / (
    f"app.py.BEFORE_WORKER_COMPARISON_{stamp}.bak"
)

shutil.copy2(APP, BACKUP)

route = r'''
@app.route("/sync/compare-online-workers")
@login_required
def compare_online_workers():
    import json
    import hashlib
    from pathlib import Path

    if not is_postgres():
        return """
        <h2>Full Worker Comparison</h2>
        <h3 style="color:red">ERROR</h3>
        <p>PostgreSQL connection is not active.</p>
        """, 400

    seed_path = Path(__file__).resolve().parent / "workers_compare_seed.json"

    if not seed_path.exists():
        return """
        <h2>Full Worker Comparison</h2>
        <h3 style="color:red">ERROR</h3>
        <p>workers_compare_seed.json was not found.</p>
        """, 500

    fields = [
        "name",
        "bangla_name",
        "department",
        "designation",
        "basic_salary",
        "ot_rate",
        "refreshment_bill",
        "phone",
        "address",
        "joining_date",
        "status",
    ]

    def canonical(value):
        if value is None:
            return ""
        return str(value).strip()

    def field_hash(value):
        return hashlib.sha256(
            canonical(value).encode("utf-8")
        ).hexdigest()

    conn = None

    try:
        with open(seed_path, "r", encoding="utf-8") as f:
            seed = json.load(f)

        if len(seed) != 169:
            raise RuntimeError(
                f"Comparison seed contains {len(seed)} workers; expected 169."
            )

        local_map = {}

        for item in seed:
            worker_id = int(item["id"])

            if worker_id in local_map:
                raise RuntimeError(
                    f"Duplicate worker ID in comparison seed: {worker_id}"
                )

            local_map[worker_id] = item["hashes"]

        conn = db_connect()
        cur = conn.cursor()

        cur.execute("""
            SELECT id,
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
            FROM workers
            ORDER BY id
        """)

        rows = cur.fetchall()

        online_ids = {int(row[0]) for row in rows}
        local_ids = set(local_map.keys())

        missing_online = sorted(local_ids - online_ids)
        extra_online = sorted(online_ids - local_ids)

        differences = []

        for row in rows:
            worker_id = int(row[0])

            if worker_id not in local_map:
                continue

            online_values = {
                field: row[i + 1]
                for i, field in enumerate(fields)
            }

            for field in fields:
                online_hash = field_hash(online_values[field])
                local_hash = local_map[worker_id].get(field, "")

                if online_hash != local_hash:
                    differences.append({
                        "id": worker_id,
                        "field": field
                    })

        db_release(conn)
        conn = None

        difference_count = len(differences)

        if (
            len(rows) == 169
            and not missing_online
            and not extra_online
            and difference_count == 0
        ):
            status_html = """
            <h2 style="color:green">
                ALL WORKER DATA MATCH
            </h2>
            """
        else:
            status_html = """
            <h2 style="color:#b00000">
                WORKER DATA DIFFERENCES FOUND
            </h2>
            """

        return render_template_string("""
        <h2>Reedoy v49 - Full Worker Comparison</h2>

        {{ status_html|safe }}

        <table border="1" cellpadding="8">
            <tr>
                <th>Check</th>
                <th>Result</th>
            </tr>

            <tr>
                <td>Local Workers</td>
                <td>169</td>
            </tr>

            <tr>
                <td>Online Workers</td>
                <td>{{ online_count }}</td>
            </tr>

            <tr>
                <td>Missing Online IDs</td>
                <td>{{ missing_online }}</td>
            </tr>

            <tr>
                <td>Extra Online IDs</td>
                <td>{{ extra_online }}</td>
            </tr>

            <tr>
                <td>Total Field Differences</td>
                <td>{{ difference_count }}</td>
            </tr>
        </table>

        {% if differences %}
        <h3>Differences</h3>

        <table border="1" cellpadding="8">
            <tr>
                <th>Worker ID</th>
                <th>Different Field</th>
            </tr>

            {% for item in differences %}
            <tr>
                <td>{{ item.id }}</td>
                <td>{{ item.field }}</td>
            </tr>
            {% endfor %}
        </table>
        {% endif %}

        <p>
        Compared fields:
        Name, Bangla Name, Department, Designation,
        Basic Salary, OT Rate, Refreshment Bill,
        Phone, Address, Joining Date, Status.
        </p>

        <p>
        <b>READ ONLY — no database changes were made.</b>
        </p>

        """,
        status_html=status_html,
        online_count=len(rows),
        missing_online=missing_online,
        extra_online=extra_online,
        difference_count=difference_count,
        differences=differences
        )

    except Exception as e:

        if conn is not None:
            try:
                db_release(conn)
            except Exception:
                pass

        return f"""
        <h2>Full Worker Comparison Error</h2>
        <p>{type(e).__name__}: {e}</p>
        <p><b>READ ONLY — no database changes were made.</b></p>
        """, 500

'''

marker = 'if __name__ == "__main__":'

if marker not in text:
    raise SystemExit("[ERROR] Application startup marker not found.")

text = text.replace(
    marker,
    route + "\n" + marker,
    1
)

APP.write_text(text, encoding="utf-8")

py_compile.compile(str(APP), doraise=True)

print()
print("=" * 70)
print("FULL WORKER COMPARISON PATCH SUCCESS")
print("=" * 70)
print("app.py syntax: PASS")
print("Backup:", BACKUP)
print("Comparison seed:", SEED)
print()
print("IMPORTANT:")
print("workers_compare_seed.json contains HASHES only.")
print("No worker names, phone numbers or addresses are stored in it.")
print()
print("Commit ONLY:")
print("  app.py")
print("  workers_compare_seed.json")
print("=" * 70)
