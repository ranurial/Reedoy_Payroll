from pathlib import Path
import shutil
import datetime
import py_compile

app = Path("app.py")

text = app.read_text(encoding="utf-8-sig")

if "/diagnostic/online-lamia" in text:
    print("Route already exists.")
    raise SystemExit(0)

backup = app.with_name(
    "app.py.BEFORE_ONLINE_LAMIA_DIAGNOSTIC_" +
    datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
)

shutil.copy2(app, backup)

route = r'''
@app.route("/diagnostic/online-lamia")
@login_required
def diagnostic_online_lamia():
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
            WHERE name ILIKE '%Lamia%'
               OR bangla_name LIKE '%লামিয়া%'
            ORDER BY id
        """)

        rows = cur.fetchall()

        db_release(conn)
        conn = None

        return render_template_string("""
        <h2>Reedoy v49 - Online Lamia Workers</h2>

        <p><b>Count: {{ rows|length }}</b></p>

        <table border="1" cellpadding="8">
            <tr>
                <th>ID</th>
                <th>English Name</th>
                <th>Bangla Name</th>
                <th>Department</th>
                <th>Designation</th>
                <th>Basic Salary</th>
                <th>OT Rate</th>
                <th>Refreshment</th>
                <th>Status</th>
            </tr>

            {% for row in rows %}
            <tr>
                {% for value in row %}
                <td>{{ value if value is not none else '' }}</td>
                {% endfor %}
            </tr>
            {% endfor %}
        </table>

        <p><b>READ ONLY — no database changes were made.</b></p>
        """, rows=rows)

    except Exception as e:

        if conn is not None:
            try:
                db_release(conn)
            except Exception:
                pass

        return f"""
        <h2>Online Lamia Diagnostic Error</h2>
        <p>{type(e).__name__}: {e}</p>
        <p><b>READ ONLY — no database changes were made.</b></p>
        """, 500

'''

marker = 'if __name__ == "__main__":'

if marker not in text:
    raise SystemExit("Startup marker not found.")

text = text.replace(marker, route + "\n" + marker, 1)

app.write_text(text, encoding="utf-8")

py_compile.compile(str(app), doraise=True)

print("=" * 70)
print("ONLINE LAMIA DIAGNOSTIC PATCH SUCCESS")
print("=" * 70)
print("app.py syntax: PASS")
print("Backup:", backup)
print()
print("Commit ONLY app.py")
