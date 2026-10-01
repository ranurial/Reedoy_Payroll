from pathlib import Path
import shutil
import datetime
import sys

p = Path("app.py")

if not p.exists():
    print("[ERROR] app.py not found")
    raise SystemExit(1)

text = p.read_text(encoding="utf-8-sig")

marker = 'if __name__ == "__main__":'

if '"/diagnostic/online-workers"' in text:
    print("[INFO] Diagnostic route already exists.")
    raise SystemExit(0)

if marker not in text:
    print("[ERROR] Could not find app startup marker.")
    raise SystemExit(1)

backup = p.with_name(
    "app.py.BEFORE_ONLINE_WORKERS_DIAGNOSTIC_"
    + datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    + ".bak"
)

shutil.copy2(p, backup)

route = r'''

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

'''

text = text.replace(marker, route + "\n" + marker, 1)

p.write_text(text, encoding="utf-8")

print("=" * 70)
print("TEMPORARY ONLINE WORKERS DIAGNOSTIC ADDED")
print("=" * 70)
print()
print("Backup created:")
print(backup)
print()
print("Route:")
print("/diagnostic/online-workers")
print()
print("Testing Python syntax...")

import py_compile
py_compile.compile(str(p), doraise=True)

print("py_compile: PASS")
print()
print("NO DATABASE DATA WAS MODIFIED.")
print("=" * 70)
