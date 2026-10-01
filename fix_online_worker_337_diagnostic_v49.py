from pathlib import Path
import re
import shutil
import datetime
import py_compile

app = Path("app.py")

if not app.exists():
    raise SystemExit("[ERROR] app.py not found")

text = app.read_text(encoding="utf-8-sig")

start_marker = '@app.route("/diagnostic/online-worker-337")'
start = text.find(start_marker)

if start == -1:
    raise SystemExit("[ERROR] Worker 337 diagnostic route not found")

# Find the next route after this one, or the application startup marker.
next_route = text.find("\n@app.route(", start + len(start_marker))
startup = text.find("\nif __name__ == \"__main__\":", start + len(start_marker))

ends = [x for x in (next_route, startup) if x != -1]

if not ends:
    raise SystemExit("[ERROR] Could not find end of diagnostic route")

end = min(ends)

old_route = text[start:end]

stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
backup = app.with_name(
    f"app.py.BEFORE_WORKER_337_DIAGNOSTIC_FIX_{stamp}.bak"
)

shutil.copy2(app, backup)

new_route = r'''
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

'''

text = text[:start] + new_route + text[end:]

app.write_text(text, encoding="utf-8")

py_compile.compile(str(app), doraise=True)

print("=" * 70)
print("WORKER 337 DIAGNOSTIC FIX SUCCESS")
print("=" * 70)
print("app.py syntax: PASS")
print("Backup:", backup)
print()
print("IMPORTANT:")
print("This patch does NOT modify the Online database.")
print("Commit ONLY app.py")
print("=" * 70)
