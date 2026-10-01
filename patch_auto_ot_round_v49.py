from pathlib import Path
import shutil
import re
import subprocess
import sys

BASE = Path(r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll")
APP = BASE / "app.py"
TPL = BASE / "templates" / "worker_form.html"

if not APP.exists() or not TPL.exists():
    raise SystemExit("ERROR: app.py or worker_form.html not found.")

app_bak = BASE / "app.py.BEFORE_AUTO_OT_ROUND_FIX.bak"
tpl_bak = BASE / "templates" / "worker_form.html.BEFORE_AUTO_OT_ROUND_FIX.bak"
shutil.copy2(APP, app_bak)
shutil.copy2(TPL, tpl_bak)

NEW_JS = r"""<script>
(function () {
  const basic = document.getElementById("basic_salary");
  const ot = document.getElementById("ot_rate");
  if (!basic || !ot) return;

  function updateOtRate() {
    const b = parseFloat(basic.value);
    if (!Number.isFinite(b) || b < 0) {
      ot.value = "";
      return;
    }
    ot.value = Math.round(b / 30 / 12);
  }

  basic.addEventListener("input", updateOtRate);
  basic.addEventListener("change", updateOtRate);
})();
</script>"""

def replace_js(text):
    pattern = r'<script>\s*\(function \(\) \{\s*const basic = document\.getElementById\("basic_salary"\);.*?</script>'
    new_text, count = re.subn(pattern, lambda m: NEW_JS, text, count=1, flags=re.S)
    if count != 1:
        raise RuntimeError("Existing Auto OT JavaScript block was not found.")
    return new_text

TPL.write_text(replace_js(TPL.read_text(encoding="utf-8")), encoding="utf-8")

app_text = APP.read_text(encoding="utf-8")
marker = "BUILTIN_WORKER_FORM_TEMPLATE = "
start = app_text.find(marker)
if start < 0:
    raise RuntimeError("BUILTIN_WORKER_FORM_TEMPLATE not found.")

triple = chr(34) * 3
qs = app_text.find(triple, start)
qe = app_text.find(triple, qs + 3)
if qs < 0 or qe < 0:
    raise RuntimeError("Built-in worker template boundaries not found.")

builtin = replace_js(app_text[qs + 3:qe])
app_text = app_text[:qs + 3] + builtin + app_text[qe:]
APP.write_text(app_text, encoding="utf-8")

result = subprocess.run(
    [sys.executable, "-m", "py_compile", str(APP)],
    cwd=str(BASE), capture_output=True, text=True
)
if result.returncode != 0:
    shutil.copy2(app_bak, APP)
    shutil.copy2(tpl_bak, TPL)
    print(result.stderr)
    raise SystemExit("ERROR: py_compile failed. Original files restored.")

print("AUTO OT RATE ROUNDING FIXED SUCCESSFULLY")
print("Formula: Basic Salary / 30 / 12")
print("Rounding: nearest whole BDT")
print("20,000 -> 56 | 30,000 -> 83 | 50,000 -> 139 | 200,000 -> 556")
print("py_compile: PASS")
print(f"Backup: {app_bak}")
print(f"Backup: {tpl_bak}")
