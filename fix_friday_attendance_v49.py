import os
import shutil
from datetime import datetime

APP = r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll\app.py"

if not os.path.isfile(APP):
    raise SystemExit("ERROR: app.py not found:\n" + APP)

with open(APP, "r", encoding="utf-8") as f:
    s = f.read()

backup = APP + ".BEFORE_FRIDAY_FIX_" + datetime.now().strftime("%Y%m%d_%H%M%S") + ".bak"
shutil.copy2(APP, backup)

old_js = """function markAllAttendance(status){
  var selects=document.querySelectorAll('select[name^="day_"]');
  selects.forEach(function(sel){
    var row=sel.closest('tr');
    var dayCell=row ? row.querySelector('.day') : null;
    var dayNumber=dayCell ? parseInt(dayCell.textContent.trim(),10) : 0;
    // Friday rows are marked with the 'weekend' class by the server.
    if(status==='P' && row && row.classList.contains('weekend')){
      sel.value='A';
    } else {
      sel.value=status;
    }
  });
}"""

new_js = """function markAllAttendance(status){
  var selects=document.querySelectorAll('select[name^="day_"]');
  selects.forEach(function(sel){
    var row=sel.closest('tr');
    var dayCell=row ? row.querySelector('.day') : null;
    var dayNumber=dayCell ? parseInt(dayCell.textContent.trim(),10) : 0;
    // Friday is the weekly holiday. It must always display as Absent.
    if(row && row.classList.contains('weekend')){
      sel.value='A';
    } else {
      sel.value=status;
    }
  });
}"""

if old_js not in s:
    raise SystemExit("ERROR: expected markAllAttendance() block was not found. No code was changed.")

s = s.replace(old_js, new_js, 1)

old_status = """        status_text = str(status or "").strip().upper()
        if not status_text or status_text in {"-", "NOT SET", "NONE"}:
"""

new_status = """        status_text = str(status or "").strip().upper()

        # Friday is the weekly holiday. Keep it as Absent in the
        # attendance calendar, while salary calculation separately
        # ignores Friday for absent deduction.
        try:
            current_date = datetime.date(
                attendance_year,
                attendance_month,
                day_number,
            )
            if current_date.weekday() == 4:
                status_text = "A"
        except Exception:
            pass

        if not status_text or status_text in {"-", "NOT SET", "NONE"}:
"""

if old_status not in s:
    raise SystemExit("ERROR: expected daily-status block was not found. No code was changed.")

s = s.replace(old_status, new_status, 1)

with open(APP, "w", encoding="utf-8", newline="") as f:
    f.write(s)

print("PATCH SUCCESS")
print("app.py updated:", APP)
print("Safety backup:", backup)
print("Changes:")
print("1. Mark All Present keeps every Friday as A (Absent).")
print("2. Saving attendance forces Friday to A.")
print("3. Existing salary logic is not changed.")
