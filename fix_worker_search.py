from pathlib import Path

p = Path("app.py")
text = p.read_text(encoding="utf-8")

old = """        conditions.append(\"\"\"
            (CAST(id AS TEXT) LIKE ?
             OR LOWER(COALESCE(name, '')) LIKE ?
             OR LOWER(COALESCE(bangla_name, '')) LIKE ?
             OR LOWER(COALESCE(designation, '')) LIKE ?
             OR LOWER(COALESCE(designation, '')) LIKE ?
             OR LOWER(COALESCE(designation, '')) LIKE ?)
        \"\"\")
        params.extend([like_value] * 6)"""

new = """        conditions.append(\"\"\"
            (CAST(id AS TEXT) LIKE ?
             OR LOWER(COALESCE(name, '')) LIKE ?
             OR LOWER(COALESCE(bangla_name, '')) LIKE ?
             OR LOWER(COALESCE(department, '')) LIKE ?
             OR LOWER(COALESCE(designation, '')) LIKE ?)
        \"\"\")
        params.extend([like_value] * 5)"""

if old not in text:
    print("ERROR: Expected search block পাওয়া যায়নি। কোনো পরিবর্তন করা হয়নি।")
else:
    text = text.replace(old, new, 1)
    p.write_text(text, encoding="utf-8")
    print("WORKER SEARCH FIXED SUCCESSFULLY")
