"""
Reedoy Payroll -> Render PostgreSQL migration
SAFE IMPORT:
- Source SQL is loaded into a temporary SQLite database (read-only source file).
- Existing Render data is NOT deleted.
- Workers are matched by name/department/designation/basic salary/OT rate where possible.
- Source worker IDs are mapped to target worker IDs.
- Monthly attendance, daily attendance, advances, users, company settings and activity logs are imported.
- Run only once from Render Shell after deploying this package.

Environment:
  DATABASE_URL must be set by Render.
  SOURCE_SQL defaults to reedoy_payroll_migration.sql

Usage:
  python render_import.py
  python render_import.py --dry-run
"""
import os, re, sqlite3, argparse, tempfile
from datetime import datetime

SOURCE_SQL = os.environ.get("SOURCE_SQL", "reedoy_payroll_migration.sql")
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

def qident(name):
    return '"' + name.replace('"','""') + '"'

def load_source():
    if not os.path.exists(SOURCE_SQL):
        raise SystemExit(f"Source SQL not found: {SOURCE_SQL}")
    fd, path = tempfile.mkstemp(suffix=".db")
    os.close(fd)
    con = sqlite3.connect(path)
    with open(SOURCE_SQL, "r", encoding="utf-8") as f:
        sql = f.read()
    # The migration SQL is SQLite-format and includes sqlite_sequence entries.
    # Ignore sqlite_sequence inserts because PostgreSQL sequences are handled separately.
    sql = re.sub(r'(?is)DELETE FROM "sqlite_sequence";', '', sql)
    sql = re.sub(r'(?is)INSERT INTO "sqlite_sequence".*?;', '', sql)
    con.executescript(sql)
    con.commit()
    return con, path

def pg():
    if not DATABASE_URL:
        raise SystemExit("DATABASE_URL is not set. Run this on Render, where the database is attached.")
    import psycopg
    return psycopg.connect(DATABASE_URL.replace("postgres://","postgresql://",1))

def ensure_pg_tables(c):
    stmts = [
        """CREATE TABLE IF NOT EXISTS daily_attendance (
            id SERIAL PRIMARY KEY,
            worker_id INTEGER NOT NULL,
            month_year TEXT NOT NULL,
            day INTEGER NOT NULL,
            status TEXT NOT NULL DEFAULT 'P',
            UNIQUE(worker_id, month_year, day)
        )""",
        """CREATE INDEX IF NOT EXISTS idx_daily_worker_month ON daily_attendance(worker_id, month_year)""",
    ]
    for s in stmts:
        c.execute(s)

def table_exists_sqlite(scon, table):
    return scon.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (table,)).fetchone() is not None

def get_source_rows(scon, table):
    if not table_exists_sqlite(scon, table):
        return []
    return scon.execute(f'SELECT * FROM "{table}"').fetchall()

def columns(scon, table):
    return [x[1] for x in scon.execute(f'PRAGMA table_info("{table}")').fetchall()]

def norm(v):
    return (str(v or "").strip().casefold())

def worker_match(pgcur, r):
    # Prefer exact source ID only when that ID already represents the same worker.
    sid, name, basic, ot_rate, department, designation, refreshment, bangla = r
    row = pgcur.execute(
        """SELECT id FROM workers
           WHERE lower(trim(name))=lower(trim(%s))
             AND lower(trim(department))=lower(trim(%s))
             AND COALESCE(lower(trim(designation)),'')=COALESCE(lower(trim(%s)),'')
             AND COALESCE(basic_salary,0)=%s
             AND COALESCE(ot_rate,0)=%s
           ORDER BY id LIMIT 1""",
        (name, department, designation, float(basic or 0), float(ot_rate or 0))
    ).fetchone()
    return row["id"] if row else None

def import_all(scon, pcon, dry=False):
    pcur=pcon.cursor()
    ensure_pg_tables(pcur)
    pcon.commit()

    src_workers=get_source_rows(scon,"workers")
    src_att=get_source_rows(scon,"attendance")
    src_daily=get_source_rows(scon,"daily_attendance")
    src_adv=get_source_rows(scon,"worker_advances")
    src_users=get_source_rows(scon,"users")
    src_settings=get_source_rows(scon,"company_settings")
    src_logs=get_source_rows(scon,"activity_log")

    print(f"Source counts: workers={len(src_workers)}, attendance={len(src_att)}, daily={len(src_daily)}, advances={len(src_adv)}, users={len(src_users)}, logs={len(src_logs)}")

    worker_map={}
    inserted=matched=0
    for r in src_workers:
        target_id=worker_match(pcur,r)
        if target_id:
            worker_map[int(r[0])]=int(target_id); matched+=1
            # Preserve Bangla name if target is blank.
            pcur.execute("UPDATE workers SET bangla_name=COALESCE(NULLIF(bangla_name,''),%s) WHERE id=%s",(r[7],target_id))
        else:
            pcur.execute("""INSERT INTO workers(name,bangla_name,department,designation,basic_salary,ot_rate,refreshment_bill)
                            VALUES(%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                         (r[1],r[7],r[4],r[5],float(r[2] or 0),float(r[3] or 0),float(r[6] or 0)))
            target_id=pcur.fetchone()["id"]
            worker_map[int(r[0])]=int(target_id); inserted+=1

    print(f"Workers: matched={matched}, inserted={inserted}")

    att_count=0
    for r in src_att:
        sid, wid, month, present, absent, ot, *rest = r
        tw=worker_map.get(int(wid))
        if not tw: continue
        pcur.execute("""INSERT INTO attendance(worker_id,month_year,present_days,absent_days,ot_hours)
                        VALUES(%s,%s,%s,%s,%s)
                        ON CONFLICT(worker_id,month_year)
                        DO UPDATE SET present_days=EXCLUDED.present_days,
                                      absent_days=EXCLUDED.absent_days,
                                      ot_hours=EXCLUDED.ot_hours""",
                     (tw,month,int(present or 0),int(absent or 0),float(ot or 0)))
        att_count+=1

    daily_count=0
    for r in src_daily:
        sid,wid,month,day,status=r[:5]
        tw=worker_map.get(int(wid))
        if not tw: continue
        pcur.execute("""INSERT INTO daily_attendance(worker_id,month_year,day,status)
                        VALUES(%s,%s,%s,%s)
                        ON CONFLICT(worker_id,month_year,day)
                        DO UPDATE SET status=EXCLUDED.status""",
                     (tw,month,int(day),status))
        daily_count+=1

    adv_count=0
    for r in src_adv:
        sid,wid,month,adv_date,amount,note=r[:6]
        tw=worker_map.get(int(wid))
        if not tw: continue
        # Avoid duplicate advances by matching worker/month/date/amount/note.
        exists=pcur.execute("""SELECT 1 FROM worker_advances
                               WHERE worker_id=%s AND month_year=%s AND advance_date=%s
                                 AND amount=%s AND COALESCE(note,'')=COALESCE(%s,'')
                               LIMIT 1""",
                            (tw,month,adv_date,float(amount or 0),note or "")).fetchone()
        if not exists:
            pcur.execute("""INSERT INTO worker_advances(worker_id,month_year,advance_date,amount,note)
                            VALUES(%s,%s,%s,%s,%s)""",
                         (tw,month,adv_date,float(amount or 0),note or ""))
            adv_count+=1

    set_count=0
    for r in src_settings:
        k,v=r[:2]
        pcur.execute("""INSERT INTO company_settings(key,value) VALUES(%s,%s)
                        ON CONFLICT(key) DO UPDATE SET value=EXCLUDED.value""",(k,v))
        set_count+=1

    user_count=0
    for r in src_users:
        # Old and new schemas both use the first seven fields in the migration package.
        vals=list(r)
        if len(vals)<7: continue
        uid,username,pwh,full_name,role,active,created_at=vals[:7]
        last_login=vals[7] if len(vals)>7 else None
        exists=pcur.execute("SELECT id FROM users WHERE username=%s",(username,)).fetchone()
        if exists:
            pcur.execute("""UPDATE users SET full_name=COALESCE(NULLIF(%s,''),full_name),
                             role=COALESCE(NULLIF(%s,''),role), active=%s,
                             created_at=COALESCE(created_at,%s), last_login=COALESCE(%s,last_login)
                             WHERE username=%s""",
                         (full_name or "",role or "User",int(active if active is not None else 1),created_at,last_login,username))
        else:
            pcur.execute("""INSERT INTO users(username,password_hash,full_name,role,active,created_at,last_login)
                            VALUES(%s,%s,%s,%s,%s,%s,%s)""",
                         (username,pwh,full_name,role or "User",int(active if active is not None else 1),created_at,last_login))
        user_count+=1

    log_count=0
    for r in src_logs:
        vals=list(r)
        if len(vals)<4: continue
        _,username,action,log_time=vals[:4]
        pcur.execute("""INSERT INTO activity_log(username,action,log_time)
                        SELECT %s,%s,%s
                        WHERE NOT EXISTS (SELECT 1 FROM activity_log WHERE username=%s AND action=%s AND log_time=%s)""",
                     (username,action,log_time,username,action,log_time))
        log_count+=1

    if dry:
        pcon.rollback()
        print("DRY RUN: all database changes rolled back.")
    else:
        pcon.commit()
        # Keep PostgreSQL sequences above imported explicit IDs if needed.
        for table in ("workers","attendance","daily_attendance","worker_advances","users","activity_log"):
            try:
                pcur.execute(f"""SELECT setval(pg_get_serial_sequence(%s,%s),
                    COALESCE((SELECT MAX(id) FROM {qident(table)}),1), true)""",(table,"id"))
            except Exception:
                pcon.rollback()
                pcon.commit()
        print("IMPORT COMPLETE")

    print(f"Imported/upserted: attendance={att_count}, daily={daily_count}, advances={adv_count}, settings={set_count}, users={user_count}, logs={log_count}")

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument("--dry-run",action="store_true")
    args=ap.parse_args()
    scon,path=load_source()
    pcon=pg()
    try:
        import_all(scon,pcon,args.dry_run)
    finally:
        pcon.close(); scon.close()
        try: os.remove(path)
        except OSError: pass

if __name__=="__main__":
    main()
