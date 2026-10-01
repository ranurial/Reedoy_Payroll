import sqlite3

conn = sqlite3.connect("factory_payroll.db")

sql = """
SELECT day, status, COUNT(*)
FROM daily_attendance
WHERE month_year = 'October 2026'
  AND day IN (2, 9, 16, 23, 30)
GROUP BY day, status
ORDER BY day, status
"""

rows = conn.execute(sql).fetchall()

print("October 2026 Friday records:")
for row in rows:
    print(row)

conn.close()