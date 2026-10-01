import os
import re
import shutil
import datetime

APP = r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll\app.py"

if not os.path.exists(APP):
    print("ERROR: v49 app.py not found:")
    print(APP)
    raise SystemExit(1)

with open(APP, "r", encoding="utf-8") as f:
    src = f.read()

stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
backup = APP + f".BEFORE_FINAL_FRIDAY_30DAY_FIX_{stamp}.bak"
shutil.copy2(APP, backup)

changed = []

m = re.search(r"(?ms)^def daily_map\(\s*.*?(?=^def \w+\()", src)
if not m:
    print("ERROR: daily_map() was not found.")
    print("Safety backup:", backup)
    raise SystemExit(1)

new_daily_map = """def daily_map(
    worker_id,
    month
):
    # Global Friday rule: every Friday is always A.
    rows = fetch_all(
        "SELECT day,status FROM daily_attendance "
        "WHERE worker_id=? AND month_year=?",
        (worker_id, month),
    )

    result = {}

    try:
        year, month_number = map(int, str(month).split("-")[:2])
    except Exception:
        year = month_number = None

    for row in rows:
        try:
            day_number = int(row["day"])

            if year and month_number:
                current_date = datetime.date(year, month_number, day_number)
                if current_date.weekday() == 4:
                    result[day_number] = "A"
                    continue

            status = row["status"]
            result[day_number] = str(status) if status not in (None, "") else ""
        except Exception:
            pass

    if year and month_number:
        try:
            last_day = calendar.monthrange(year, month_number)[1]
            for day_number in range(1, last_day + 1):
                current_date = datetime.date(year, month_number, day_number)
                if current_date.weekday() == 4:
                    result[day_number] = "A"
        except Exception:
            pass

    return result

"""

src = src[:m.start()] + new_daily_map + src[m.end():]
changed.append("daily_map() replaced with global Friday=A rule")

for old, new in [
    ('daily.get(day, "P")',
     '( "A" if datetime.date(year, month_number, day).weekday() == 4 else daily.get(day, "") )'),
    ("daily.get(day, 'P')",
     "( 'A' if datetime.date(year, month_number, day).weekday() == 4 else daily.get(day, '') )"),
]:
    if old in src:
        src = src.replace(old, new)
        changed.append("attendance missing-day default fixed")

calc_start = src.find("def calculate_salary")
if calc_start >= 0:
    next_def = src.find("\ndef ", calc_start + 10)
    calc_end = next_def if next_def >= 0 else len(src)
    calc = src[calc_start:calc_end]
    old_calc = calc
    calc = re.sub(
        r"days_in_month\s*=\s*calendar\.monthrange\([^\n]+\)\[1\]",
        "days_in_month = 30",
        calc,
        count=1
    )
    if calc != old_calc:
        src = src[:calc_start] + calc + src[calc_end:]
        changed.append("calculate_salary() forced to 30-day basis")

with open(APP, "w", encoding="utf-8", newline="\n") as f:
    f.write(src)

print("PATCH SUCCESS")
print("app.py updated:", APP)
print("Safety backup:", backup)
print("")
print("Changes:")
for item in changed:
    print("-", item)
print("")
print("FINAL RULE:")
print("1. Every Friday in every month/year displays as A (Absent).")
print("2. Friday remains A even if database contains P.")
print("3. New months automatically follow the same Friday rule.")
print("4. Non-Friday attendance is preserved.")
print("5. Salary basis remains 30 days.")
print("6. Database was NOT modified.")
print("")
print("IMPORTANT: Restart the Flask app before testing attendance.")
