from pathlib import Path

p = Path("app.py")
text = p.read_text(encoding="utf-8")

start_marker = '@app.route("/sync/compare-online-workers")'
end_marker = '\n\nif __name__ == "__main__":'

start = text.find(start_marker)
end = text.find(end_marker, start)

if start < 0 or end < 0:
    raise SystemExit("ERROR: comparison route boundaries not found")

new_block = r'''@app.route("/sync/compare-online-workers")
@login_required
def compare_online_workers():
    import json
    import hashlib
    from pathlib import Path
    from collections import Counter

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

        if isinstance(value, float):
            return str(round(value, 10)).strip()

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
                        "field": field,
                        "online": canonical(online_values[field]),
                        "local_hash": local_hash,
                    })

        db_release(conn)
        conn = None

        # Build a read-only field summary.
        field_counts = Counter(
            item["field"] for item in differences
        )

        # Read the local database separately to obtain the actual
        # local values for the differing fields.
        local_conn = None

        try:
            local_conn = sqlite3.connect(DB_PATH)
            local_conn.row_factory = sqlite3.Row

            local_cur = local_conn.cursor()

            local_cur.execute("""
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

            local_rows = local_cur.fetchall()

            local_value_map = {
                int(row["id"]): dict(row)
                for row in local_rows
            }

        finally:
            if local_conn is not None:
                local_conn.close()

        detailed_differences = []

        for item in differences:
            worker_id = item["id"]
            field = item["field"]

            local_row = local_value_map.get(worker_id, {})
            local_value = canonical(local_row.get(field, ""))

            detailed_differences.append({
                "id": worker_id,
                "field": field,
                "local": local_value,
                "online": item["online"],
            })

        difference_count = len(detailed_differences)

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

        {% if field_counts %}
        <h3>Differences by Field</h3>

        <table border="1" cellpadding="8">
            <tr>
                <th>Field</th>
                <th>Difference Count</th>
            </tr>

            {% for field, count in field_counts %}
            <tr>
                <td>{{ field }}</td>
                <td>{{ count }}</td>
            </tr>
            {% endfor %}
        </table>
        {% endif %}

        {% if detailed_differences %}
        <h3>Detailed Differences</h3>

        <table border="1" cellpadding="8">
            <tr>
                <th>Worker ID</th>
                <th>Different Field</th>
                <th>Local Value</th>
                <th>Online Value</th>
            </tr>

            {% for item in detailed_differences %}
            <tr>
                <td>{{ item.id }}</td>
                <td>{{ item.field }}</td>
                <td>{{ item.local }}</td>
                <td>{{ item.online }}</td>
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
        field_counts=sorted(
            field_counts.items(),
            key=lambda x: (-x[1], x[0])
        ),
        detailed_differences=detailed_differences
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

p.write_text(
    text[:start] + new_block + text[end:],
    encoding="utf-8"
)

print("SUCCESS: Read-only detailed worker comparison route patched.")
print("No database operation was performed.")
