from pathlib import Path
import sqlite3
import shutil
import datetime

BASE = Path(r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll")
DB = BASE / "factory_payroll.db"

if not DB.exists():
    raise SystemExit(f"ERROR: Database not found: {DB}")

stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = BASE / f"factory_payroll_BEFORE_OT_RATE_UPDATE_{stamp}.db"
shutil.copy2(DB, BACKUP)

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row

rows = conn.execute(
    "SELECT id, name, basic_salary, ot_rate FROM workers ORDER BY id"
).fetchall()

changed = []
for row in rows:
    basic = float(row["basic_salary"] or 0)
    if basic < 0:
        continue

    new_ot = int(basic / 30 / 12 + 0.5)

    old_ot = float(row["ot_rate"] or 0)
    if old_ot != new_ot:
        changed.append((row["id"], row["name"], basic, old_ot, new_ot))

conn.executemany(
    "UPDATE workers SET ot_rate=? WHERE id=?",
    [(new_ot, worker_id) for worker_id, _, _, _, new_ot in changed]
)
conn.commit()

check = conn.execute(
    "SELECT COUNT(*) AS c FROM workers"
).fetchone()["c"]

conn.close()

print("EXISTING WORKER OT RATE UPDATE SUCCESSFUL")
print(f"Workers checked: {check}")
print(f"Workers changed: {len(changed)}")
print(f"Safety backup: {BACKUP}")
print()
print("Formula: Basic Salary / 30 / 12")
print("Rounding: nearest whole BDT")
print()
print("First 20 changed workers:")
for worker_id, name, basic, old_ot, new_ot in changed[:20]:
    print(f"{worker_id}: {name} | Basic {basic:.2f} | {old_ot:.2f} -> {new_ot}")
