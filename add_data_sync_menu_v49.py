from pathlib import Path

p = Path("templates") / "dashboard.html"
backup = Path("templates") / "dashboard.html.BEFORE_DATA_SYNC_MENU_FIX.bak"

text = p.read_text(encoding="utf-8")
backup.write_text(text, encoding="utf-8")

# Do not add twice
if "url_for('data_sync')" in text:
    print("Data Sync menu already exists.")
else:
    old = """    <a href="{{ url_for('department') }}">{{ tr('Department') }}</a>
    <a href="{{ url_for('settings') }}">{{ tr('Settings') }}</a>
"""

    new = """    <a href="{{ url_for('department') }}">{{ tr('Department') }}</a>
    <a href="{{ url_for('data_sync') }}">Data Sync</a>
    <a href="{{ url_for('settings') }}">{{ tr('Settings') }}</a>
"""

    if old not in text:
        raise SystemExit("ERROR: Dashboard menu block not found.")

    text = text.replace(old, new, 1)
    p.write_text(text, encoding="utf-8")
    print("DATA SYNC MENU ADDED SUCCESSFULLY")

print("Backup:", backup)