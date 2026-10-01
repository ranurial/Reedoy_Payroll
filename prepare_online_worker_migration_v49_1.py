from pathlib import Path
import sqlite3
import hashlib
import json
import shutil
import datetime
import py_compile

BASE = Path(__file__).resolve().parent
APP = BASE / "app.py"
DB = BASE / "factory_payroll.db"
SEED = BASE / "workers_sync_seed.json"

if not APP.exists():
    raise SystemExit("[ERROR] app.py not found")

if not DB.exists():
    raise SystemExit("[ERROR] factory_payroll.db not found")

text = APP.read_text(encoding="utf-8-sig")

route_marker = '@app.route("/diagnostic/online-workers")'

if "/sync/prepare-online-workers" in text:
    print("[INFO] Online worker migration route already exists.")
    raise SystemExit(0)

# ------------------------------------------------------------
# 1. Create safe seed from LOCAL database
# ------------------------------------------------------------
conn = sqlite3.connect(str(DB))

rows = conn.execute("""
    SELECT id, name, sync_uuid
    FROM workers
    ORDER BY id
""").fetchall()

conn.close()

if len(rows) != 169:
    raise SystemExit(
        f"[ERROR] Expected 169 local workers, found {len(rows)}"
    )

seed = []

for worker_id, name, sync_uuid in rows:
    if not sync_uuid:
        raise SystemExit(
            f"[ERROR] Worker {worker_id} has no sync_uuid"
        )

    identity = f"{worker_id}|{name or ''}"
    name_hash = hashlib.sha256(
        identity.encode("utf-8")
    ).hexdigest()

    seed.append({
        "id": worker_id,
        "identity_hash": name_hash,
        "sync_uuid": sync_uuid,
    })

SEED.write_text(
    json.dumps(seed, ensure_ascii=False, indent=2),
    encoding="utf-8"
)

print("[OK] Local sync seed created:")
print(SEED)
print("Seed workers:", len(seed))

# ------------------------------------------------------------
# 2. Backup app.py
# ------------------------------------------------------------
stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

BACKUP = BASE / (
    f"app.py.BEFORE_ONLINE_WORKER_MIGRATION_{stamp}.bak"
)

shutil.copy2(APP, BACKUP)

print("[OK] app.py backup created:")
print(BACKUP)

# ------------------------------------------------------------
# 3. Temporary online migration route
# ------------------------------------------------------------
route = r'''

# ============================================================
# TEMPORARY ONLINE WORKERS UUID MIGRATION
# ============================================================

@app.route("/sync/prepare-online-workers")
@login_required
def prepare_online_workers_sync():
    import json
    import hashlib
    from pathlib import Path

    if not is_postgres():
        return """
        <h2>Online Worker Sync Preparation</h2>
        <p>ERROR: This route must run against PostgreSQL.</p>
        """, 400

    seed_path = Path(__file__).resolve().parent / "workers_sync_seed.json"

    if not seed_path.exists():
        return """
        <h2>Online Worker Sync Preparation</h2>
        <p>ERROR: workers_sync_seed.json not found.</p>
        """, 500

    conn = None

    try:
        with open(seed_path, "r", encoding="utf-8") as f:
            seed = json.load(f)

        conn = db_connect()
        cur = conn.cursor()

        # ----------------------------------------------------
        # Read current Online workers
        # ----------------------------------------------------
        cur.execute("""
            SELECT id, name
            FROM workers
            ORDER BY id
        """)

        online_rows = cur.fetchall()

        if len(online_rows) != len(seed):
            raise RuntimeError(
                f"Worker count mismatch. "
                f"Online={len(online_rows)}, Seed={len(seed)}. "
                f"NO CHANGES WERE MADE."
            )

        seed_map = {
            int(row["id"]): row
            for row in seed
        }

        mismatches = []

        for row in online_rows:
            worker_id = int(row[0])
            online_name = row[1] or ""

            if worker_id not in seed_map:
                mismatches.append(
                    f"ID {worker_id}: not found in local seed"
                )
                continue

            identity = f"{worker_id}|{online_name}"

            online_hash = hashlib.sha256(
                identity.encode("utf-8")
            ).hexdigest()

            if online_hash != seed_map[worker_id]["identity_hash"]:
                mismatches.append(
                    f"ID {worker_id}: name mismatch"
                )

        if mismatches:
            conn.rollback()

            html = """
            <h2>Online Worker Sync Preparation</h2>
            <h3 style="color:red">
                STOP - NO DATABASE CHANGES WERE MADE
            </h3>
            <p>
                The Online workers do not exactly match the Local
                workers.
            </p>
            <p>Mismatches: {{ count }}</p>
            <ul>
            {% for item in mismatches %}
                <li>{{ item }}</li>
            {% endfor %}
            </ul>
            """

            return render_template_string(
                html,
                count=len(mismatches),
                mismatches=mismatches
            ), 409

        # ----------------------------------------------------
        # All workers match
        # ----------------------------------------------------
        cur.execute("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema='public'
              AND table_name='workers'
              AND column_name='sync_uuid'
        """)

        has_sync_uuid = cur.fetchone() is not None

        if not has_sync_uuid:
            cur.execute("""
                ALTER TABLE workers
                ADD COLUMN sync_uuid TEXT
            """)

        # ----------------------------------------------------
        # Make sure existing UUIDs are not conflicting
        # ----------------------------------------------------
        cur.execute("""
            SELECT id, sync_uuid
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
        """)

        existing_uuid_rows = cur.fetchall()

        existing_uuid_map = {
            int(row[0]): str(row[1])
            for row in existing_uuid_rows
        }

        conflicts = []

        for worker_id, local_data in seed_map.items():
            existing = existing_uuid_map.get(worker_id)

            if existing and existing != local_data["sync_uuid"]:
                conflicts.append(
                    f"ID {worker_id}: existing sync_uuid differs"
                )

        if conflicts:
            conn.rollback()

            html = """
            <h2>Online Worker Sync Preparation</h2>
            <h3 style="color:red">
                STOP - UUID CONFLICT
            </h3>
            <p>NO DATABASE CHANGES WERE MADE.</p>
            <ul>
            {% for item in conflicts %}
                <li>{{ item }}</li>
            {% endfor %}
            </ul>
            """

            return render_template_string(
                html,
                conflicts=conflicts
            ), 409

        # ----------------------------------------------------
        # Assign UUIDs
        # ----------------------------------------------------
        updated = 0

        for worker_id, local_data in seed_map.items():

            cur.execute("""
                UPDATE workers
                SET sync_uuid=%s
                WHERE id=%s
                  AND (
                      sync_uuid IS NULL
                      OR TRIM(sync_uuid)=''
                  )
            """, (
                local_data["sync_uuid"],
                worker_id
            ))

            updated += cur.rowcount

        # ----------------------------------------------------
        # Unique index
        # ----------------------------------------------------
        cur.execute("""
            CREATE UNIQUE INDEX IF NOT EXISTS
            idx_workers_sync_uuid
            ON workers(sync_uuid)
        """)

        # ----------------------------------------------------
        # Sync metadata table
        # ----------------------------------------------------
        cur.execute("""
            CREATE TABLE IF NOT EXISTS reedoy_sync_meta (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT
            )
        """)

        # ----------------------------------------------------
        # Workers baseline table
        # ----------------------------------------------------
        cur.execute("""
            CREATE TABLE IF NOT EXISTS reedoy_sync_worker_base (
                sync_uuid TEXT PRIMARY KEY,
                worker_id INTEGER,
                fingerprint TEXT,
                snapshot_json TEXT,
                last_synced_at TEXT
            )
        """)

        conn.commit()

        # ----------------------------------------------------
        # Final verification
        # ----------------------------------------------------
        cur.execute("""
            SELECT COUNT(*)
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
        """)

        uuid_count = int(cur.fetchone()[0])

        cur.execute("SELECT COUNT(*) FROM workers")
        worker_count = int(cur.fetchone()[0])

        db_release(conn)
        conn = None

        html = """
        <h2>Online Worker Sync Preparation</h2>

        <h3 style="color:green">
            SUCCESS
        </h3>

        <table border="1" cellpadding="8">
            <tr>
                <th>Item</th>
                <th>Result</th>
            </tr>
            <tr>
                <td>Online Workers</td>
                <td>{{ worker_count }}</td>
            </tr>
            <tr>
                <td>Workers with sync_uuid</td>
                <td>{{ uuid_count }}</td>
            </tr>
            <tr>
                <td>New UUIDs assigned</td>
                <td>{{ updated }}</td>
            </tr>
        </table>

        <p>
        Existing worker IDs, names, salary, OT, department,
        attendance, advance and payment data were not modified.
        </p>

        <p>
        <b>Workers two-way sync foundation is now ready.</b>
        </p>
        """

        return render_template_string(
            html,
            worker_count=worker_count,
            uuid_count=uuid_count,
            updated=updated
        )

    except Exception as e:

        if conn is not None:
            try:
                conn.rollback()
            except Exception:
                pass

            try:
                db_release(conn)
            except Exception:
                pass

        return f"""
        <h2>Online Worker Sync Preparation</h2>
        <h3 style="color:red">ERROR - NO CHANGES WERE COMMITTED</h3>
        <p>{type(e).__name__}: {e}</p>
        """, 500

'''

# ------------------------------------------------------------
# 4. Insert route before startup
# ------------------------------------------------------------
marker = 'if __name__ == "__main__":'

if marker not in text:
    raise SystemExit(
        "[ERROR] app startup marker not found"
    )

text = text.replace(
    marker,
    route + "\n" + marker,
    1
)

APP.write_text(text, encoding="utf-8")

# ------------------------------------------------------------
# 5. Syntax check
# ------------------------------------------------------------
py_compile.compile(str(APP), doraise=True)

print()
print("=" * 70)
print("ONLINE WORKER MIGRATION PATCH SUCCESSFUL")
print("=" * 70)
print()
print("app.py syntax: PASS")
print("Seed file:", SEED.name)
print("Seed workers:", len(seed))
print()
print("NO ONLINE DATABASE WAS TOUCHED.")
print("NO LOCAL DATABASE WAS MODIFIED.")
print()
print("Next: commit app.py + workers_sync_seed.json")
print("=" * 70)
