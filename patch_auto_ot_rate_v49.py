from pathlib import Path
import shutil
import re
import subprocess
import sys

BASE = Path(r"D:\Reedoy Textile\Software\Reedoy_Payroll_v49\Reedoy_Payroll")
APP = BASE / "app.py"
TPL = BASE / "templates" / "worker_form.html"

if not APP.exists():
    raise SystemExit(f"ERROR: app.py not found: {APP}")
if not TPL.exists():
    raise SystemExit(f"ERROR: worker_form.html not found: {TPL}")

app_bak = APP.with_name("app.py.BEFORE_AUTO_OT_RATE.bak")
tpl_bak = TPL.with_name("worker_form.html.BEFORE_AUTO_OT_RATE.bak")
shutil.copy2(APP, app_bak)
shutil.copy2(TPL, tpl_bak)

JS = r'''
<script>
(function () {
  const basic = document.getElementById("basic_salary");
  const ot = document.getElementById("ot_rate");
  if (!basic || !ot) return;
  let autoMode = true;

  function setAutoRate() {
    if (!autoMode) return;
    const b = parseFloat(basic.value);
    if (!Number.isFinite(b) || b < 0) {
      ot.value = "";
      return;
    }
    ot.value = (b / 30 / 12).toFixed(2);
  }

  basic.addEventListener("input", function () {
    autoMode = true;
    setAutoRate();
  });

  ot.addEventListener("input", function () {
    autoMode = false;
  });

  const initialOt = parseFloat(ot.value);
  if (!Number.isFinite(initialOt) || initialOt === 0) {
    setAutoRate();
  } else {
    autoMode = false;
  }
})();
</script>
'''

def patch_template(text):
    text = re.sub(
        r'(<input\b(?=[^>]*\bname=["\']basic_salary["\'])(?![^>]*\bid=["\']basic_salary["\'])[^>]*)(>)',
        r'\1 id="basic_salary"\2', text, count=1, flags=re.I
    )
    text = re.sub(
        r'(<input\b(?=[^>]*\bname=["\']ot_rate["\'])(?![^>]*\bid=["\']ot_rate["\'])[^>]*)(>)',
        r'\1 id="ot_rate"\2', text, count=1, flags=re.I
    )
    if 'id="basic_salary"' not in text or 'id="ot_rate"' not in text:
        raise RuntimeError("Basic Salary / OT Rate inputs not found")
    if "setAutoRate()" not in text:
        if "</form>" in text:
            text = text.replace("</form>", JS + "\n</form>", 1)
        elif "</body>" in text:
            text = text.replace("</body>", JS + "\n</body>", 1)
        else:
            text += "\n" + JS + "\n"
    return text

TPL.write_text(patch_template(TPL.read_text(encoding="utf-8")), encoding="utf-8")

app_text = APP.read_text(encoding="utf-8")
marker = "BUILTIN_WORKER_FORM_TEMPLATE = "
start = app_text.find(marker)
if start < 0:
    raise RuntimeError("BUILTIN_WORKER_FORM_TEMPLATE not found")

triple = chr(34) * 3
quote_start = app_text.find(triple, start)
if quote_start < 0:
    raise RuntimeError("Built-in template opening quote not found")
quote_end = app_text.find(triple, quote_start + 3)
if quote_end < 0:
    raise RuntimeError("Built-in template closing quote not found")

builtin = app_text[quote_start + 3:quote_end]
builtin_new = patch_template(builtin)
app_text = app_text[:quote_start + 3] + builtin_new + app_text[quote_end:]
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

print("AUTO OT RATE FIXED SUCCESSFULLY")
print("Formula: Basic Salary / 30 / 12")
print("Example: 200000 / 30 / 12 = 555.56")
print(f"Backup app.py: {app_bak}")
print(f"Backup worker_form.html: {tpl_bak}")
print("Add + Edit Worker: Basic Salary change auto-updates OT Rate.")
print("Manual OT Rate editing remains allowed.")
print("py_compile: PASS")
