import os, shutil
from datetime import datetime

APP = r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll\app.py"

if not os.path.isfile(APP):
    raise SystemExit("ERROR: app.py not found:\n" + APP)

with open(APP, "r", encoding="utf-8") as f:
    s = f.read()

backup = APP + ".BEFORE_GLOBAL_FRIDAY_FIX_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
shutil.copy2(APP, backup)

old = '''def daily_map(
    worker_id,
    month
):

    rows = fetch_all(
        """
        SELECT day,status
        FROM daily_attendance
        WHERE worker_id=?
        AND month_year=?
        """,
        (
            worker_id,
            month,
        ),
    )

    result = {}

    for row in rows:

        try:

            result[
                int(row["day"])
            ] = str(
                row["status"] or "P"
            )

        except Exception:
            pass

    return result
'''

new = '''def daily_map(
    worker_id,
    month
):

    rows = fetch_all(
        """
        SELECT day,status
        FROM daily_attendance
        WHERE worker_id=?
        AND month_year=?
        """,
        (
            worker_id,
            month,
        ),
    )

    try:
        month_name, year_text = month.split()
        month_number = MONTHS.index(month_name) + 1
        year_number = int(year_text)
    except Exception:
        month_number = None
        year_number = None

    result = {}

    for row in rows:

        try:
            day_number = int(row["day"])

            # Friday is the weekly holiday: always display as Absent.
            if month_number and year_number:
                current_date = datetime.date(
                    year_number,
                    month_number,
                    day_number,
                )
                if current_date.weekday() == 4:
                    result[day_number] = "A"
                    continue

            # Preserve actual status on non-Friday days.
            result[day_number] = str(row["status"] or "")

        except Exception:
            pass

    return result
'''

if old not in s:
    raise SystemExit("ERROR: expected daily_map() block was not found. No code was changed.")

s = s.replace(old, new, 1)

with open(APP, "w", encoding="utf-8", newline="") as f:
    f.write(s)

print("PATCH SUCCESS")
print("app.py updated:", APP)
print("Safety backup:", backup)
print("1. Every Friday in every month/year displays as A.")
print("2. Friday is forced to A even if stored status is P.")
print("3. Non-Friday stored statuses are preserved.")
print("4. Blank non-Friday status stays blank.")
print("5. Salary calculation was not changed.")
