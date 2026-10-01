from pathlib import Path
import shutil
import datetime
import py_compile

app = Path("app.py")

if not app.exists():
    raise SystemExit("[ERROR] app.py not found")

text = app.read_text(encoding="utf-8-sig")

route_path = "/sync/verify-online-worker-ids"

if route_path in text:
    print("[INFO] Exact ID verification route already exists.")
    raise SystemExit(0)

backup = app.with_name(
    "app.py.BEFORE_EXACT_WORKER_ID_VERIFY_" +
    datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
)

shutil.copy2(app, backup)

route = r'''
@app.route("/sync/verify-online-worker-ids")
@login_required
def verify_online_worker_ids():
    import json
    from collections import Counter
    from pathlib import Path

    if not is_postgres():
        return """
        <h2>Exact Worker ID Verification</h2>
        <h3 style="color:red">ERROR</h3>
        <p>PostgreSQL connection is not active.</p>
        """, 400

    seed_path = Path(__file__).resolve().parent / "workers_sync_seed.json"

    if not seed_path.exists():
        return """
        <h2>Exact Worker ID Verification</h2>
        <h3 style="color:red">ERROR</h3>
        <p>workers_sync_seed.json was not found.</p>
        """, 500

    conn = None

    try:
        with open(seed_path, "r", encoding="utf-8") as f:
            seed = json.load(f)

        # ----------------------------------------------------
        # LOCAL SEED CHECK
        # ----------------------------------------------------
        local_ids = [int(item["id"]) for item in seed]

        local_duplicates = sorted(
            [
                worker_id
                for worker_id, count
                in Counter(local_ids).items()
                if count > 1
            ]
        )

        local_uuid_values = [
            str(item.get("sync_uuid") or "").strip()
            for item in seed
        ]

        uuid_counter = Counter(local_uuid_values)

        duplicate_uuids = sorted(
            [
                value
                for value, count in uuid_counter.items()
                if value and count > 1
            ]
        )

        # ----------------------------------------------------
        # ONLINE CHECK
        # ----------------------------------------------------
        conn = db_connect()
        cur = conn.cursor()

        cur.execute("""
            SELECT id
            FROM workers
            ORDER BY id
        """)

        online_rows = cur.fetchall()

        online_ids = [int(row[0]) for row in online_rows]

        online_duplicates = sorted(
            [
                worker_id
                for worker_id, count
                in Counter(online_ids).items()
                if count > 1
            ]
        )

        local_set = set(local_ids)
        online_set = set(online_ids)

        missing_online = sorted(local_set - online_set)
        extra_online = sorted(online_set - local_set)

        exact_match = (
            len(local_ids) == 169
            and len(online_ids) == 169
            and not local_duplicates
            and not online_duplicates
            and not duplicate_uuids
            and not missing_online
            and not extra_online
        )

        db_release(conn)
        conn = None

        if exact_match:
            status_html = """
            <h2 style="color:green">
                EXACT WORKER ID MATCH — SAFE TO PROCEED
            </h2>
            """
        else:
            status_html = """
            <h2 style="color:red">
                STOP — WORKER ID MATCH FAILED
            </h2>
            """

        return render_template_string("""
        <h2>Reedoy v49 - Exact Worker ID Verification</h2>

        {{ status_html|safe }}

        <table border="1" cellpadding="8">
            <tr>
                <th>Check</th>
                <th>Result</th>
            </tr>

            <tr>
                <td>Local Seed Workers</td>
                <td>{{ local_count }}</td>
            </tr>

            <tr>
                <td>Online Workers</td>
                <td>{{ online_count }}</td>
            </tr>

            <tr>
                <td>Duplicate Local IDs</td>
                <td>{{ local_duplicates }}</td>
            </tr>

            <tr>
                <td>Duplicate Online IDs</td>
                <td>{{ online_duplicates }}</td>
            </tr>

            <tr>
                <td>Duplicate Local UUIDs</td>
                <td>{{ duplicate_uuids }}</td>
            </tr>

            <tr>
                <td>IDs Missing Online</td>
                <td>{{ missing_online }}</td>
            </tr>

            <tr>
                <td>IDs Extra Online</td>
                <td>{{ extra_online }}</td>
            </tr>
        </table>

        <h3>Important</h3>

        <p>
        This verification uses <b>Worker ID only</b>.
        Worker names are NOT used for pairing.
        </p>

        <p>
        <b>READ ONLY — no database changes were made.</b>
        </p>

        """,
        status_html=status_html,
        local_count=len(local_ids),
        online_count=len(online_ids),
        local_duplicates=local_duplicates,
        online_duplicates=online_duplicates,
        duplicate_uuids=duplicate_uuids,
        missing_online=missing_online,
        extra_online=extra_online
        )

    except Exception as e:

        if conn is not None:
            try:
                db_release(conn)
            except Exception:
                pass

        return f"""
        <h2>Exact Worker ID Verification Error</h2>
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

app.write_text(text, encoding="utf-8")

py_compile.compile(str(app), doraise=True)

print("=" * 70)
print("EXACT WORKER ID VERIFICATION PATCH SUCCESS")
print("=" * 70)
print("app.py syntax: PASS")
print("Backup:", backup)
print()
print("IMPORTANT:")
print("This patch is READ ONLY.")
print("No Online database data was changed.")
print()
print("Commit ONLY: app.py")
print("=" * 70)
