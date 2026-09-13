# Reedoy Payroll — Render Migration Package

This package contains the current Flask/Render application plus the old payroll migration SQL and a safe PostgreSQL importer.

## Files
- `app.py` — current Flask application (patched to create `daily_attendance`)
- `render_import.py` — safe importer from the old SQLite-format migration SQL into Render PostgreSQL
- `reedoy_payroll_migration.sql` — old verified migration dataset
- `render.yaml` — Render service/database definition
- `requirements.txt` — dependencies

## Render steps
1. Push this project to the Git repository connected to Render.
2. Deploy the web service normally.
3. Open the Render service -> **Shell**.
4. Run:
   `python render_import.py --dry-run`
5. If the counts look correct, run:
   `python render_import.py`
6. Log in with the imported admin account if present. If there is no admin row, the app creates:
   username: `admin`
   password: `admin123`
   Change it immediately.

### Safety
The importer does NOT delete existing Render records. It matches existing workers where possible and inserts the rest. Attendance and daily attendance are upserted. Advances, users, settings and activity logs are duplicate-aware.

The old SQL source is never modified; it is loaded into a temporary SQLite database.
