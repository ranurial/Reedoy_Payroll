import sqlite3, json, os, datetime

DB_PATH = r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll\factory_payroll.db"
OUT_PATH = r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll\workers_review_values.json"

FIELDS = ["name", "bangla_name", "designation", "status"]

def main():
    if not os.path.exists(DB_PATH):
        print("ERROR: factory_payroll.db পাওয়া যায়নি:")
        print(DB_PATH)
        return

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        rows = conn.execute("""
            SELECT id, name, bangla_name, designation, status
            FROM workers
            ORDER BY id
        """).fetchall()
    finally:
        conn.close()

    data = []
    for r in rows:
        item = {"id": int(r["id"])}
        for f in FIELDS:
            v = r[f]
            item[f] = "" if v is None else str(v)
        data.append(item)

    with open(OUT_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print("=" * 60)
    print("REEDOY - LOCAL WORKER REVIEW VALUES")
    print("=" * 60)
    print(f"Workers exported: {len(data)}")
    print(f"Output file: {OUT_PATH}")
    print()
    print("Fields: id, name, bangla_name, designation, status")
    print("Phone/address/salary/attendance/advance/payment data are NOT exported.")
    print()
    print("এই JSON file GitHub project-এর root folder-এ রাখবেন।")
    print("তারপর GitHub Desktop দিয়ে commit + push করবেন।")

if __name__ == "__main__":
    main()
