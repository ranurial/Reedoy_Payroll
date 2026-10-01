import sqlite3

conn = sqlite3.connect("factory_payroll.db")

rows = conn.execute("""
SELECT id, name, bangla_name, department, designation, basic_salary
FROM workers
WHERE name LIKE '%Lamia%'
   OR bangla_name LIKE '%লামিয়া%'
ORDER BY id
""").fetchall()

print("Lamia workers:")
print("-" * 80)

for row in rows:
    print(row)

print("-" * 80)
print("Total:", len(rows))

conn.close()
