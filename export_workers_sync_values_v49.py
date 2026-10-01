import sqlite3, json, os, sys

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "factory_payroll.db")
OUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "workers_sync_values.json")

FIELDS = ["id", "sync_uuid", "name", "bangla_name", "designation", "status", "ot_rate"]

print("=" * 60)
print("REEDOY - WORKER SYNC VALUES EXPORT")
print("=" * 60)

if not os.path.exists(DB_PATH):
    print(f"ERROR: factory_payroll.db not found:\n{DB_PATH}")
    sys.exit(1)

conn = sqlite3.connect(DB_PATH)
conn.row_factory = sqlite3.Row
try:
    cols = {r[1] for r in conn.execute("PRAGMA table_info(workers)").fetchall()}
    missing = [c for c in FIELDS if c not in cols]
    if missing:
        print("ERROR: workers table is missing:", ", ".join(missing))
        sys.exit(1)

    rows = conn.execute("""
        SELECT id, sync_uuid, name, bangla_name, designation, status, ot_rate
        FROM workers
        ORDER BY id
    """).fetchall()

    if len(rows) != 169:
        print(f"ERROR: Expected 169 workers, found {len(rows)}.")
        sys.exit(1)

    seen_ids = set()
    seen_uuids = set()
    data = []

    for r in rows:
        worker_id = int(r["id"])
        sync_uuid = str(r["sync_uuid"] or "").strip()

        if worker_id in seen_ids:
            print(f"ERROR: Duplicate worker ID: {worker_id}")
            sys.exit(1)
        if not sync_uuid:
            print(f"ERROR: Worker {worker_id} has no sync_uuid.")
            sys.exit(1)
        if sync_uuid in seen_uuids:
            print(f"ERROR: Duplicate sync_uuid for worker {worker_id}.")
            sys.exit(1)

        seen_ids.add(worker_id)
        seen_uuids.add(sync_uuid)

        data.append({
            "id": worker_id,
            "sync_uuid": sync_uuid,
            "name": r["name"],
            "bangla_name": r["bangla_name"],
            "designation": r["designation"],
            "status": r["status"],
            "ot_rate": r["ot_rate"],
        })

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"Workers exported: {len(data)}")
    print(f"Output file: {OUT_PATH}")
    print("Fields:", ", ".join(FIELDS))
    print("")
    print("This file is for the controlled Worker Local -> Online sync.")
    print("Do NOT change the JSON manually.")
finally:
    conn.close()
