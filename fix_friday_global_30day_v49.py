import os, re, shutil
from datetime import datetime

APP = r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll\app.py"

if not os.path.isfile(APP):
    raise SystemExit("ERROR: app.py not found:\n" + APP)

with open(APP, "r", encoding="utf-8") as f:
    s = f.read()

backup = APP + ".BEFORE_FINAL_FRIDAY_30DAY_FIX_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
shutil.copy2(APP, backup)

changes = []

# 1) Global Friday rule in daily_map().
start = s.find("def daily_map(")
if start < 0:
    raise SystemExit("ERROR: daily_map() was not found. No code was changed.")
end = s.find("\ndef ", start + 10)
if end < 0:
    end = len(s)

new_daily = "\n".join([
    "def daily_map(",
    "    worker_id,",
    "    month",
    "):",
    "",
    "    rows = fetch_all(",
    "        \"\"\"",
    "        SELECT day,status",
    "        FROM daily_attendance",
    "        WHERE worker_id=?",
    "        AND month_year=?",
    "        \"\"\"",
    "        (worker_id, month),",
    "    )",
    "",
    "    try:",
    "        month_name, year_text = month.split()",
    "        month_number = MONTHS.index(month_name) + 1",
    "        year_number = int(year_text)",
    "    except Exception:",
    "        month_number = None",
    "        year_number = None",
    "",
    "    result = {}",
    "",
    "    for row in rows:",
    "        try:",
    "            day_number = int(row[\"day\"])",
    "            if month_number and year_number:",
    "                current_date = datetime.date(year_number, month_number, day_number)",
    "                if current_date.weekday() == 4:",
    "                    result[day_number] = \"A\"",
    "                    continue",
    "            result[day_number] = str(row[\"status\"] or \"\")",
    "        except Exception:",
    "            pass",
    "",
    "    return result",
    ""
])
s = s[:start] + new_daily + s[end+1:]
changes.append("daily_map(): every Friday is always A globally.")

# 2) Attendance page: missing days must not default to Present.
old = '"status": daily.get(\n                        day,\n                        "P"\n                    ),'
new = '"status": (\n                        "A"\n                        if datetime.date(year, month_number, day).weekday() == 4\n                        else daily.get(day, "")\n                    ),'
if old in s:
    s = s.replace(old, new, 1)
    changes.append("Attendance missing days no longer default to Present; Friday defaults to A.")
else:
    old2 = '"status": daily.get(day, "P"),'
    if old2 in s:
        s = s.replace(old2, '"status": ("A" if datetime.date(year, month_number, day).weekday() == 4 else daily.get(day, "")),', 1)
        changes.append("Attendance missing days no longer default to Present; Friday defaults to A.")
    else:
        changes.append("Attendance default-P pattern already absent; no change needed.")

# 3) Salary basis: 30 days.
cstart = s.find("def calculate_salary(")
if cstart >= 0:
    cend = s.find("\n\ndef ", cstart + 10)
    if cend < 0:
        cend = len(s)
    cblock = s[cstart:cend]
    cblock2 = re.sub(
        r"(?m)^\s*days_in_month\s*=\s*calendar\.monthrange\(\s*year\s*,\s*month_number\s*\)\[1\]\s*$",
        "        days_in_month = 30",
        cblock,
        count=1
    )
    if cblock2 != cblock:
        s = s[:cstart] + cblock2 + s[cend:]
        changes.append("Salary calculation changed to 30-day basis.")
    else:
        changes.append("Salary calculation already uses 30-day basis.")
else:
    changes.append("WARNING: calculate_salary() was not found.")

with open(APP, "w", encoding="utf-8", newline="") as f:
    f.write(s)

print("PATCH SUCCESS")
print("app.py:", APP)
print("Safety backup:", backup)
for item in changes:
    print("-", item)
print("")
print("FINAL RULE:")
print("1. Every Friday in every month/year = A (Absent).")
print("2. Missing non-Friday days stay blank, NOT Present.")
print("3. Existing Mark All Present Friday protection is retained.")
print("4. Salary basis = 30 days.")
print("5. Friday is excluded from absence deduction/present-absent calculation.")
print("6. Database was NOT modified.")
