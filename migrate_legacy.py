import os, sqlite3, sys
from pathlib import Path
import psycopg

BASE=Path(__file__).resolve().parent
SQL_FILE=BASE/"reedoy_payroll_migration.sql"
DB=os.environ.get("DATABASE_URL","").strip()
MARK="legacy_sql_migration_2026_09"

if not DB: raise SystemExit("DATABASE_URL missing")
if not SQL_FILE.exists(): raise SystemExit("reedoy_payroll_migration.sql missing")
if DB.startswith("postgres://"): DB="postgresql://"+DB[11:]

# Initialize the PostgreSQL schema using the same schema code as the live app.
# This is needed because the migration script runs instead of gunicorn.
import app  # noqa: F401 - app.init_db() runs on import

def q(x): return '"' + x.replace('"','""') + '"'

src=sqlite3.connect(":memory:")
src.execute("PRAGMA foreign_keys=OFF")
src.executescript(SQL_FILE.read_text(encoding="utf-8-sig"))

def stables():
    return {r[0] for r in src.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'")}

def scols(t):
    return [r[1] for r in src.execute(f"PRAGMA table_info({q(t)})")]

order=["workers","users","admin_users","settings","company_settings","attendance","daily_attendance","worker_advances","advance_salary","advances","activity_log"]
aliases={"settings":["settings","company_settings"],"company_settings":["company_settings","settings"]}

try:
  with psycopg.connect(DB) as con:
    with con.cursor() as cur:
      cur.execute("CREATE TABLE IF NOT EXISTS migration_markers (key TEXT PRIMARY KEY, completed_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP)")
      cur.execute("SELECT 1 FROM migration_markers WHERE key=%s",(MARK,))
      if cur.fetchone():
        print("Migration already completed."); sys.exit(0)

      cur.execute("SELECT table_name FROM information_schema.tables WHERE table_schema='public'")
      targets={r[0] for r in cur.fetchall()}
      sources=stables()
      maps=[]; used=set()
      for s in order:
        if s not in sources: continue
        target=next((x for x in aliases.get(s,[s]) if x in targets),None)
        if target and target not in used:
          maps.append((s,target)); used.add(target)

      if not maps: raise RuntimeError("No compatible target tables found")

      # The live app creates a default admin user and default company settings.
      # This migration is intended for the fresh Reedoy PostgreSQL database, so
      # remove those seed rows before restoring the legacy data.
      # Do not remove the migration marker table itself.
      clear_order=[t for t in ["activity_log","worker_advances","attendance","workers","company_settings","users"] if t in targets]
      for t in clear_order:
        cur.execute(f"DELETE FROM {q(t)}")
      print("Cleared fresh app seed data; restoring legacy data...")

      total=0
      for s,t in maps:
        cur.execute("SELECT column_name FROM information_schema.columns WHERE table_schema='public' AND table_name=%s ORDER BY ordinal_position",(t,))
        tc=[r[0] for r in cur.fetchall()]
        common=[c for c in scols(s) if c in tc]
        if not common: continue
        cols=", ".join(q(c) for c in common)
        rows=src.execute(f"SELECT {cols} FROM {q(s)}").fetchall()
        if rows:
          cur.executemany(f"INSERT INTO {q(t)} ({cols}) VALUES ({', '.join(['%s']*len(common))})",rows)
        print(f"{s} -> {t}: {len(rows)} rows")
        total+=len(rows)

      if "workers" in targets:
        cur.execute("SELECT COUNT(*),COALESCE(MAX(id),0) FROM workers")
        n,m=cur.fetchone()
        print(f"Workers: {n}, MAX ID: {m}")
        if m<355: raise RuntimeError("Worker validation failed")

      for _,t in maps:
        cur.execute("SELECT 1 FROM information_schema.columns WHERE table_schema='public' AND table_name=%s AND column_name='id'",(t,))
        if cur.fetchone():
          cur.execute("SELECT pg_get_serial_sequence(%s,'id')",(f"public.{t}",))
          seq=cur.fetchone()[0]
          if seq:
            cur.execute(f"SELECT MAX(id) FROM {q(t)}"); mx=cur.fetchone()[0]
            if mx is not None: cur.execute("SELECT setval(%s,%s,true)",(seq,mx))

      cur.execute("INSERT INTO migration_markers(key) VALUES(%s)",(MARK,))
      con.commit()
      print("SUCCESS - Total rows copied:",total)
finally:
  src.close()
