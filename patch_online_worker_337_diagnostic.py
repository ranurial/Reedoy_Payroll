from pathlib import Path
import shutil
import datetime
import py_compile

app = Path("app.py")

text = app.read_text(encoding="utf-8-sig")

if "/diagnostic/online-worker-337" in text:
    print("Route already exists.")
    raise SystemExit

backup = app.with_name(
    "app.py.BEFORE_ONLINE_WORKER_337_DIAGNOSTIC_" +
    datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
)

shutil.copy2(app, backup)

route = r'''

@app.route("/diagnostic/online-worker-337")
@login_required
def diagnostic_online_worker_337():
    if not is_postgres():
        return "ERROR: PostgreSQL connection is not active.", 400

    try:
        row = fetch_one("""
            SELECT id, name, bangla_name, department,
                   designation, basic_salary, ot_rate,
                   refreshment_bill, status
            FROM workers
            WHERE id=337
        """)

        if not row:
            return "<h2>Online Worker ID 337 NOT FOUND</h2>", 404

        return render_template_string("""
        <h2>Reedoy v49 - Online Worker ID 337</h2>
        <table border="1" cellpadding="8">
            <tr><th>Field</th><th>Value</th></tr>
            <tr><td>ID</td><td>{{ row[0] }}</td></tr>
            <tr><td>English Name</td><td>{{ row[1] }}</td></tr>
            <tr><td>Bangla Name</td><td>{{ row[2] }}</td></tr>
            <tr><td>Department</td><td>{{ row[3] }}</td></tr>
            <tr><td>Designation</td><td>{{ row[4] }}</td></tr>
            <tr><td>Basic Salary</td><td>{{ row[5] }}</td></tr>
            <tr><td>OT Rate</td><td>{{ row[6] }}</td></tr>
            <tr><td>Refreshment Bill</td><td>{{ row[7] }}</td></tr>
            <tr><td>Status</td><td>{{ row[8] }}</td></tr>
        </table>
        <p><b>READ ONLY — no database changes were made.</b></p>
        """, row=row)

    except Exception as e:
        return f"<h2>ERROR</h2><p>{type(e).__name__}: {e}</p>", 500

'''

marker = 'if __name__ == "__main__":'

if marker not in text:
    raise SystemExit("Startup marker not found.")

text = text.replace(marker, route + "\n" + marker, 1)

app.write_text(text, encoding="utf-8")

py_compile.compile(str(app), doraise=True)

print("PATCH SUCCESS")
print("Syntax: PASS")
print("Backup:", backup)
print()
print("Commit ONLY app.py")
