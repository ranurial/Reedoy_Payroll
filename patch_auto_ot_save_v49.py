from pathlib import Path

app = Path("app.py")
backup = Path("app.py.BEFORE_AUTO_OT_SAVE_FIX.bak")

text = app.read_text(encoding="utf-8")
backup.write_text(text, encoding="utf-8")

old = """        ot_rate = parse_num(
            form.get("ot_rate"),
            0
        )
"""

if text.count(old) != 2:
    raise SystemExit(
        f"ERROR: Expected 2 OT blocks, found {text.count(old)}"
    )

# ---------------- ADD WORKER ----------------
add_new = """        # Auto OT Rate for new worker
        # Formula = Basic Salary / 30 / 12
        # Rounded to nearest whole BDT
        ot_rate = int((basic_salary / 30 / 12) + 0.5) if basic_salary > 0 else 0
"""

first = text.find(old)
text = text[:first] + add_new + text[first + len(old):]

# ---------------- EDIT WORKER ----------------
old_pos = text.find(old)

if old_pos == -1:
    raise SystemExit("ERROR: Edit Worker OT block not found")

edit_new = """        # OT Rate is recalculated automatically when Basic Salary changes.
        # If Basic Salary is unchanged, keep the manually entered OT Rate.
        submitted_ot_rate = parse_num(
            form.get("ot_rate"),
            0
        )

        old_basic_salary = parse_num(
            worker.get("basic_salary"),
            0
        )

        if abs(basic_salary - old_basic_salary) > 0.000001:
            ot_rate = (
                int((basic_salary / 30 / 12) + 0.5)
                if basic_salary > 0
                else 0
            )
        else:
            ot_rate = submitted_ot_rate
"""

text = text[:old_pos] + edit_new + text[old_pos + len(old):]

app.write_text(text, encoding="utf-8")

# Syntax check
compile(text, "app.py", "exec")

print()
print("==============================================")
print("AUTO OT SAVE LOGIC FIXED SUCCESSFULLY")
print("==============================================")
print("Add Worker:")
print("  OT = Basic Salary / 30 / 12")
print("  Rounded to nearest whole BDT")
print()
print("Edit Worker:")
print("  Basic Salary changed -> OT automatically recalculated")
print("  Basic Salary unchanged -> existing/manual OT preserved")
print()
print("Backup created:")
print(backup.resolve())
print()
print("Syntax check: PASS")