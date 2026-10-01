from pathlib import Path

p = Path("app.py")
s = p.read_text(encoding="utf-8")

for n, line in enumerate(s.splitlines(), 1):
    if "daily_map" in line or "weekday()" in line or "weekend" in line or "days.append" in line:
        print(f"{n}: {line}")

print("\n--- attendance section ---")
lines = s.splitlines()
for i in range(5050, 5215):
    if i <= len(lines):
        print(f"{i}: {lines[i-1]}")