import sqlite3
from pathlib import Path

BASE = Path(__file__).resolve().parent
DB = BASE / "factory_payroll.db"

print("=" * 70)
print("REEDOY v49 - TWO-WAY SYNC READINESS CHECK")
print("=" * 70)

if not DB.exists():
    print("[ERROR] factory_payroll.db not found:")
    print(DB)
    raise SystemExit

print("[OK] Database:", DB)
print()

conn = sqlite3.connect(str(DB))
conn.row_factory = sqlite3.Row

tables = [
    "workers",
    "attendance",
    "daily_attendance",
    "worker_advances",
    "advance_salary",
    "advances",
    "payroll_payments",
    "payroll_locks",
    "payroll_lock",
    "company_settings",
    "users",
    "activity_log",
]

print("TABLES / COLUMNS")
print("-" * 70)

existing = {}

for table in tables:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table,)
    ).fetchone()

    if not row:
        continue

    cols = conn.execute(
        f'PRAGMA table_info("{table}")'
    ).fetchall()

    names = [r["name"] for r in cols]
    count = conn.execute(
        f'SELECT COUNT(*) FROM "{table}"'
    ).fetchone()[0]

    existing[table] = names

    print(f"{table}: {count} rows")
    print("  columns:", ", ".join(names))

print()
print("WORKERS CHECK")
print("-" * 70)

if "workers" in existing:
    print(
        "Worker count:",
        conn.execute("SELECT COUNT(*) FROM workers").fetchone()[0]
    )

    for field in ["id", "name", "basic_salary", "ot_rate",
                  "department", "bangla_name", "status",
                  "phone", "address", "sync_uuid"]:
        print(
            f"Has {field}:",
            "YES" if field in existing["workers"] else "NO"
        )

print()
print("SYNC METADATA")
print("-" * 70)

for table in ["reedoy_sync_worker_base", "reedoy_sync_meta"]:
    row = conn.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table,)
    ).fetchone()
    print(table + ":", "EXISTS" if row else "NOT FOUND")

print()
print("SAMPLE WORKERS")
print("-" * 70)

if "workers" in existing:
    cols = [
        c for c in [
            "id", "name", "basic_salary", "ot_rate",
            "department", "bangla_name", "status",
            "sync_uuid"
        ]
        if c in existing["workers"]
    ]

    rows = conn.execute(
        f'SELECT {",".join(cols)} FROM workers ORDER BY id LIMIT 5'
    ).fetchall()

    for row in rows:
        print(dict(row))

conn.close()

print()
print("=" * 70)
print("CHECK COMPLETE - NO DATA WAS MODIFIED")
print("=" * 70)
