import sqlite3
import hashlib

conn = sqlite3.connect("factory_payroll.db")

rows = conn.execute(
    "SELECT id, name FROM workers ORDER BY id"
).fetchall()

text = ""

for worker_id, name in rows:
    text += str(worker_id) + "|" + str(name or "") + "\n"

fingerprint = hashlib.sha256(
    text.encode("utf-8")
).hexdigest()

print("COUNT:", len(rows))
print("ID RANGE:", rows[0][0], "-", rows[-1][0])
print("IDENTITY FINGERPRINT:", fingerprint)

conn.close()
