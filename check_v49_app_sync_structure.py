from pathlib import Path

p = Path("app.py")

if not p.exists():
    print("[ERROR] app.py not found:", p.resolve())
    raise SystemExit

text = p.read_text(encoding="utf-8-sig")

print("=" * 70)
print("REEDOY v49 - APP.PY SYNC STRUCTURE CHECK")
print("=" * 70)

patterns = [
    "def is_postgres",
    "DATABASE_URL",
    "ThreadedConnectionPool",
    "def fetch_all",
    "def fetch_one",
    "def execute",
    "def data_sync",
    "@app.route(\"/data-sync\"",
    "@app.route(\"/data-sync/create\"",
    "@app.route(\"/data-sync/restore\"",
    "data_sync.html",
    "workers",
    "worker_advances",
    "payroll_payments",
]

lines = text.splitlines()

for pattern in patterns:
    print()
    print("-" * 70)
    print("SEARCH:", pattern)
    found = False

    for i, line in enumerate(lines, 1):
        if pattern.lower() in line.lower():
            found = True
            start = max(1, i - 2)
            end = min(len(lines), i + 5)

            print(f"Found near line {i}:")
            for n in range(start, end + 1):
                print(f"{n}: {lines[n-1]}")

    if not found:
        print("NOT FOUND")

print()
print("=" * 70)
print("CHECK COMPLETE - app.py WAS NOT MODIFIED")
print("=" * 70)
