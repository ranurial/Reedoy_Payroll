import sqlite3

conn = sqlite3.connect("factory_payroll.db")

rows = conn.execute(
    "SELECT id, name FROM workers ORDER BY id"
).fetchall()

print("=" * 60)
print("WORKERS ID CHECK")
print("=" * 60)

print("Total:", len(rows))
print("First 10:")
for row in rows[:10]:
    print(row)

print()
print("Last 10:")
for row in rows[-10:]:
    print(row)

print()
print("IDs 385 and above:")
for row in rows:
    if row[0] >= 385:
        print(row)

conn.close()
