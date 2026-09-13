# Reedoy Textile Payroll - Web Version

Browser-based payroll/attendance system prepared for Render + PostgreSQL.

## Local test
1. Install Python 3.11+.
2. `pip install -r requirements.txt`
3. `python app.py`
4. Open `http://127.0.0.1:5000`

Default login: **admin / admin123**

## Render
Create a new Web Service from this folder/repository and add a PostgreSQL database. The included `render.yaml` can be used as a Blueprint. The app reads `DATABASE_URL` automatically.

## Important
- The web app stores department values internally in English and only translates them for display.
- Bengali worker names are stored separately in `bangla_name`.
- SQLite is used only as a local fallback. PostgreSQL is recommended for multiple users over the Internet.
- Change the default admin password immediately after first login.
- The PDF export uses a standard Latin font; for guaranteed Bengali PDF output, add a Bengali Unicode TTF font to the project and register it in `report_pdf()`.
