import sqlite3
import uuid
import shutil
from pathlib import Path
from datetime import datetime

BASE = Path(__file__).resolve().parent
DB = BASE / "factory_payroll.db"

print("=" * 70)
print("REEDOY v49.1 - WORKERS SYNC FOUNDATION")
print("=" * 70)

if not DB.exists():
    print("[ERROR] factory_payroll.db not found:")
    print(DB)
    raise SystemExit(1)

# ------------------------------------------------------------
# Safety backup
# ------------------------------------------------------------
stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = BASE / f"factory_payroll_BEFORE_WORKERS_SYNC_{stamp}.db"

shutil.copy2(DB, BACKUP)

print()
print("[OK] Safety backup created:")
print(BACKUP)

conn = sqlite3.connect(str(DB))
conn.row_factory = sqlite3.Row

try:
    # --------------------------------------------------------
    # Check workers table
    # --------------------------------------------------------
    table = conn.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type='table' AND name='workers'
    """).fetchone()

    if not table:
        raise RuntimeError("workers table not found")

    columns = [
        row["name"]
        for row in conn.execute('PRAGMA table_info("workers")').fetchall()
    ]

    # --------------------------------------------------------
    # Add sync_uuid only if missing
    # --------------------------------------------------------
    if "sync_uuid" not in columns:
        print()
        print("[STEP] Adding workers.sync_uuid ...")

        conn.execute("""
            ALTER TABLE workers
            ADD COLUMN sync_uuid TEXT
        """)

        conn.commit()

        print("[OK] workers.sync_uuid added")
    else:
        print()
        print("[OK] workers.sync_uuid already exists")

    # --------------------------------------------------------
    # Create unique index
    # --------------------------------------------------------
    conn.execute("""
        CREATE UNIQUE INDEX IF NOT EXISTS
        idx_workers_sync_uuid
        ON workers(sync_uuid)
    """)

    conn.commit()

    # --------------------------------------------------------
    # Generate UUID for existing workers
    # --------------------------------------------------------
    workers = conn.execute("""
        SELECT id, sync_uuid
        FROM workers
        ORDER BY id
    """).fetchall()

    assigned = 0
    already = 0

    for worker in workers:
        existing_uuid = (worker["sync_uuid"] or "").strip()

        if existing_uuid:
            already += 1
            continue

        new_uuid = str(uuid.uuid4())

        conn.execute(
            "UPDATE workers SET sync_uuid=? WHERE id=?",
            (new_uuid, worker["id"])
        )

        assigned += 1

    conn.commit()

    print()
    print("[STEP] Worker UUID assignment")
    print("Workers checked:", len(workers))
    print("UUID already present:", already)
    print("UUID newly assigned:", assigned)

    # --------------------------------------------------------
    # Sync metadata table
    # --------------------------------------------------------
    conn.execute("""
        CREATE TABLE IF NOT EXISTS reedoy_sync_meta (
            meta_key TEXT PRIMARY KEY,
            meta_value TEXT
        )
    """)

    # --------------------------------------------------------
    # Workers baseline table
    # --------------------------------------------------------
    conn.execute("""
        CREATE TABLE IF NOT EXISTS reedoy_sync_worker_base (
            sync_uuid TEXT PRIMARY KEY,
            worker_id INTEGER,
            fingerprint TEXT,
            snapshot_json TEXT,
            last_synced_at TEXT
        )
    """)

    conn.commit()

    # --------------------------------------------------------
    # Verification
    # --------------------------------------------------------
    total = conn.execute(
        "SELECT COUNT(*) FROM workers"
    ).fetchone()[0]

    uuid_count = conn.execute("""
        SELECT COUNT(*)
        FROM workers
        WHERE sync_uuid IS NOT NULL
          AND TRIM(sync_uuid) <> ''
    """).fetchone()[0]

    duplicate_count = conn.execute("""
        SELECT COUNT(*)
        FROM (
            SELECT sync_uuid
            FROM workers
            WHERE sync_uuid IS NOT NULL
              AND TRIM(sync_uuid) <> ''
            GROUP BY sync_uuid
            HAVING COUNT(*) > 1
        )
    """).fetchone()[0]

    meta_exists = conn.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type='table'
          AND name='reedoy_sync_meta'
    """).fetchone()

    base_exists = conn.execute("""
        SELECT name
        FROM sqlite_master
        WHERE type='table'
          AND name='reedoy_sync_worker_base'
    """).fetchone()

    print()
    print("=" * 70)
    print("VERIFICATION")
    print("=" * 70)
    print("Workers:", total)
    print("Workers with sync_uuid:", uuid_count)
    print("Duplicate sync_uuid groups:", duplicate_count)
    print(
        "reedoy_sync_meta:",
        "OK" if meta_exists else "ERROR"
    )
    print(
        "reedoy_sync_worker_base:",
        "OK" if base_exists else "ERROR"
    )

    if total != uuid_count:
        raise RuntimeError(
            f"Not every worker has sync_uuid: {uuid_count}/{total}"
        )

    if duplicate_count != 0:
        raise RuntimeError("Duplicate sync_uuid detected")

    print()
    print("=" * 70)
    print("WORKERS SYNC FOUNDATION SUCCESSFUL")
    print("=" * 70)
    print()
    print("IMPORTANT:")
    print("- Existing worker IDs were NOT changed.")
    print("- Worker names/salary/department were NOT changed.")
    print("- Attendance was NOT changed.")
    print("- Advance/payment data was NOT changed.")
    print("- No online database was touched.")
    print()
    print("Safety backup:")
    print(BACKUP)

except Exception as e:
    conn.rollback()
    print()
    print("=" * 70)
    print("ERROR - DATABASE WAS ROLLED BACK")
    print("=" * 70)
    print(type(e).__name__ + ":", e)
    raise

finally:
    conn.close()
