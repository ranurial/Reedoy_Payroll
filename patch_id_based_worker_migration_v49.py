from pathlib import Path
import shutil
import datetime
import py_compile

app = Path("app.py")

if not app.exists():
    raise SystemExit("[ERROR] app.py not found")

text = app.read_text(encoding="utf-8-sig")

route_path = "/sync/migrate-online-workers-by-id"

if route_path in text:
    print("[INFO] ID-based migration route already exists.")
    raise SystemExit(0)

backup = app.with_name(
    "app.py.BEFORE_ID_BASED_WORKER_MIGRATION_" +
    datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
)

shutil.copy2(app, backup)

route = r'''
@app.route("/sync/migrate-online-workers-by-id")
@login_required
def migrate_online_workers_by_id():
    import json
    from collections import Counter
    from pathlib import Path
    from datetime import datetime, timezone

    if not is_postgres():
        return """
        <h2>Online Worker UUID Migration</h2>
        <h3 style="color:red">STOP</h3>
        <p>PostgreSQL connection is not active.</p>
        """, 400

    seed_path = Path(__file__).resolve().parent / "workers_sync_seed.json"

    if not seed_path.exists():
        return """
        <h2>Online Worker UUID Migration</h2>
        <h3 style="color:red">STOP</h3>
        <p>workers_sync_seed.json was not found.</p>
        """, 500

    conn = None

    try:
        with open(seed_path, "r", encoding="utf-8") as f:
            seed = json.load(f)

        # ----------------------------------------------------
        # Validate LOCAL seed
        # ----------------------------------------------------
        if len(seed) != 169:
            raise RuntimeError(
                f"Local seed contains {len(seed)} workers, expected 169."
            )

        seed_map = {}

        for item in seed:
            worker_id = int(item["id"])
            sync_uuid = str(item.get("sync_uuid") or "").strip()

            if not sync_uuid:
                raise RuntimeError(
                    f"Local worker ID {worker_id} has no sync_uuid."
                )

            if worker_id in seed_map:
                raise RuntimeError(
                    f"Duplicate worker ID in local seed: {worker_id}"
                )

            seed_map[worker_id] = sync_uuid

        uuid_values = list(seed_map.values())
        duplicate_uuids = [
            value
            for value, count in Counter(uuid_values).items()
            if count > 1
        ]

        if duplicate_uuids:
            raise RuntimeError(
                "Duplicate sync_uuid values found in local seed."
            )

        # ----------------------------------------------------
        # Connect to Online PostgreSQL
        # ----------------------------------------------------
        conn = db_connect()
        cur = conn.cursor()

        # ----------------------------------------------------
        # Verify Online worker IDs
        # ----------------------------------------------------
        cur.execute("""
            SELECT id
            FROM workers
            ORDER BY id
        """)

        online_rows = cur.fetchall()
        online_ids = [int(row[0]) for row in online_rows]

        if len(online_ids) != 169:
            raise RuntimeError(
                f"Online worker count is {len(online_ids)}, expected 169."
            )

        if len(set(online_ids)) != len(online_ids):
            raise RuntimeError(
                "Duplicate worker IDs exist Online. "
                "NO CHANGES WERE MADE."
            )

        local_ids = set(seed_map.keys())
        online_id_set = set(online_ids)

        missing_online = sorted(local_ids - online_id_set)
        extra_online = sorted(online_id_set - local_ids)

        if missing_online or extra_online:
            raise RuntimeError(
                "Worker ID mismatch. "
                f"Missing Online={missing_online}; "
                f"Extra Online={extra_online}. "
                "NO CHANGES WERE MADE."
            )

        # ----------------------------------------------------
        # Check sync_uuid column
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
        # Check existing Online UUIDs
        # ----------------------------------------------------
        cur.execute("""
            SELECT id, sync_uuid
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
        """)

        existing_rows = cur.fetchall()

        existing_map = {
            int(row[0]): str(row[1]).strip()
            for row in existing_rows
        }

        conflicts = []

        for worker_id, local_uuid in seed_map.items():
            existing_uuid = existing_map.get(worker_id)

            if existing_uuid and existing_uuid != local_uuid:
                conflicts.append(
                    f"ID {worker_id}: existing Online UUID differs"
                )

        if conflicts:
            raise RuntimeError(
                "UUID conflicts detected: " +
                "; ".join(conflicts[:20])
            )

        # ----------------------------------------------------
        # Ensure no Online UUID belongs to another worker
        # ----------------------------------------------------
        cur.execute("""
            SELECT sync_uuid, COUNT(*)
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
            GROUP BY sync_uuid
            HAVING COUNT(*) > 1
        """)

        duplicate_online_uuid_rows = cur.fetchall()

        if duplicate_online_uuid_rows:
            raise RuntimeError(
                "Duplicate sync_uuid values already exist Online."
            )

        # ----------------------------------------------------
        # Assign Local UUID to matching Online ID
        # ----------------------------------------------------
        updated = 0

        for worker_id, local_uuid in seed_map.items():

            cur.execute("""
                UPDATE workers
                SET sync_uuid=%s
                WHERE id=%s
                  AND (
                      sync_uuid IS NULL
                      OR TRIM(sync_uuid)=''
                  )
            """, (
                local_uuid,
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
        # Sync metadata
        # ----------------------------------------------------
        cur.execute("""
            CREATE TABLE IF NOT EXISTS reedoy_sync_meta (
                meta_key TEXT PRIMARY KEY,
                meta_value TEXT
            )
        """)

        # ----------------------------------------------------
        # Worker baseline table
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

        # ----------------------------------------------------
        # Store migration state
        # ----------------------------------------------------
        migration_time = datetime.now(timezone.utc).isoformat()

        cur.execute("""
            INSERT INTO reedoy_sync_meta(meta_key, meta_value)
            VALUES (%s, %s)
            ON CONFLICT(meta_key)
            DO UPDATE SET meta_value=EXCLUDED.meta_value
        """, (
            "workers_uuid_migration",
            migration_time
        ))

        cur.execute("""
            INSERT INTO reedoy_sync_meta(meta_key, meta_value)
            VALUES (%s, %s)
            ON CONFLICT(meta_key)
            DO UPDATE SET meta_value=EXCLUDED.meta_value
        """, (
            "workers_uuid_migration_count",
            str(updated)
        ))

        # ----------------------------------------------------
        # FINAL verification BEFORE COMMIT
        # ----------------------------------------------------
        cur.execute("""
            SELECT COUNT(*)
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
        """)

        uuid_count = int(cur.fetchone()[0])

        if uuid_count != 169:
            raise RuntimeError(
                f"Final UUID verification failed: {uuid_count}/169"
            )

        cur.execute("""
            SELECT COUNT(*)
            FROM workers
            WHERE sync_uuid IS NULL
               OR TRIM(sync_uuid)=''
        """)

        missing_uuid_count = int(cur.fetchone()[0])

        if missing_uuid_count != 0:
            raise RuntimeError(
                f"{missing_uuid_count} workers still have no sync_uuid."
            )

        # ----------------------------------------------------
        # COMMIT — only after every safety check passed
        # ----------------------------------------------------
        conn.commit()

        db_release(conn)
        conn = None

        return render_template_string("""
        <h2>Reedoy v49 - Online Worker UUID Migration</h2>

        <h2 style="color:green">
            SUCCESS — WORKER UUID MIGRATION COMPLETE
        </h2>

        <table border="1" cellpadding="8">
            <tr>
                <th>Check</th>
                <th>Result</th>
            </tr>

            <tr>
                <td>Local Worker IDs</td>
                <td>169</td>
            </tr>

            <tr>
                <td>Online Worker IDs</td>
                <td>169</td>
            </tr>

            <tr>
                <td>Workers with sync_uuid</td>
                <td>{{ uuid_count }}</td>
            </tr>

            <tr>
                <td>New UUIDs assigned</td>
                <td>{{ updated }}</td>
            </tr>

            <tr>
                <td>Missing UUIDs</td>
                <td>0</td>
            </tr>
        </table>

        <p>
        Existing Worker IDs, names, salary, OT, department,
        attendance, advance and payment data were not changed.
        </p>

        <p>
        <b>Two-way Worker Sync foundation is now installed.</b>
        </p>

        <p>
        Migration time: {{ migration_time }}
        </p>
        """,
        uuid_count=uuid_count,
        updated=updated,
        migration_time=migration_time
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
        <h2>Reedoy v49 - Online Worker UUID Migration</h2>

        <h2 style="color:red">
            STOP — NO DATABASE CHANGES WERE COMMITTED
        </h2>

        <p>
        <b>{type(e).__name__}:</b> {e}
        </p>

        <p>
        The transaction was rolled back.
        </p>
        """, 409

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
print("ID-BASED ONLINE WORKER MIGRATION PATCH SUCCESS")
print("=" * 70)
print("app.py syntax: PASS")
print("Backup:", backup)
print()
print("This route:")
print("  - uses Worker ID, not worker name")
print("  - checks all 169 IDs")
print("  - checks duplicate UUIDs")
print("  - uses a PostgreSQL transaction")
print("  - rolls back on any error")
print()
print("Commit ONLY: app.py")
print("=" * 70)
