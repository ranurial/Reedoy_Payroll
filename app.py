import os, csv, io, sqlite3, hashlib
from datetime import datetime, date
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file, abort

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'CHANGE-ME-IN-RENDER')
DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
DB_PATH = os.environ.get('SQLITE_PATH', os.path.join(os.path.dirname(__file__), 'reedoy_payroll.db'))

BN = {'Dashboard':'ড্যাশবোর্ড','Workers':'কর্মী','Attendance':'উপস্থিতি','Advances':'অগ্রিম বেতন','Payroll Report':'বেতন রিপোর্ট','Department Salary':'বিভাগভিত্তিক বেতন','Users':'ইউজার','Activity Log':'কার্যক্রম লগ','Company Settings':'কোম্পানি সেটিংস','Logout':'লগআউট','Login':'লগইন','Username':'ইউজার আইডি','Password':'পাসওয়ার্ড','Name':'নাম','Bangla Name':'বাংলা নাম','Department':'বিভাগ','Designation':'পদবী','Basic Salary':'মূল বেতন','OT Rate/Hour':'ওটি হার/ঘণ্টা','Refreshment':'নাস্তা','Save':'সংরক্ষণ','Search':'খুঁজুন','Month':'মাস','Present':'উপস্থিত','Absent':'অনুপস্থিত','OT Hours':'ওটি ঘণ্টা','Advance':'অগ্রিম','Net Payable':'নেট প্রদেয়','Gross Salary':'মোট বেতন','Add Worker':'কর্মী যোগ করুন','Edit':'সম্পাদনা','Delete':'মুছুন','Generate':'তৈরি করুন','Export CSV':'CSV ডাউনলোড','Export Excel':'Excel ডাউনলোড','Export PDF':'PDF ডাউনলোড','Add Advance':'অগ্রিম যোগ করুন','Date':'তারিখ','Amount':'পরিমাণ','Note':'নোট','Action':'কার্যক্রম','Total':'মোট','Worker ID':'কর্মী আইডি','Company Name':'কোম্পানির নাম','Address':'ঠিকানা','Phone':'ফোন','Email':'ইমেইল','Save Settings':'সেটিংস সংরক্ষণ','Active':'সক্রিয়','Role':'রোল','No data':'কোনো তথ্য নেই','Current Language':'বর্তমান ভাষা','English':'ইংরেজি','Bangla':'বাংলা','Salary Report':'বেতন রিপোর্ট','Payslip':'পে-স্লিপ','New Password':'নতুন পাসওয়ার্ড','Full Name':'পূর্ণ নাম','Create User':'ইউজার তৈরি করুন','User Management':'ইউজার ম্যানেজমেন্ট','Status':'স্ট্যাটাস','Department Filter':'বিভাগ নির্বাচন'}
DEPT_BN={'All Departments':'সব বিভাগ','Cutting':'কাটিং','Sewing':'সেলাই','Finishing':'ফিনিশিং','Quality Control':'কোয়ালিটি কন্ট্রোল','Knitting':'নিটিং','Dyeing':'ডাইং','Printing':'প্রিন্টিং','Washing':'ওয়াশিং','Iron & Packing':'আয়রন ও প্যাকিং','Maintenance':'মেইনটেন্যান্স','Store & Logistics':'স্টোর ও লজিস্টিকস','Human Resources & Admin':'মানবসম্পদ ও প্রশাসন','Accounts & Commercial':'অ্যাকাউন্টস ও কমার্শিয়াল','Utility':'ইউটিলিটি'}

def tr(x): return BN.get(x,x) if session.get('lang','en')=='bn' else x
def dept(x): return DEPT_BN.get(x,x) if session.get('lang','en')=='bn' else x
app.jinja_env.globals.update(tr=tr,dept=dept)

def sha(s): return hashlib.sha256(s.encode('utf-8')).hexdigest()

class DB:
    def __init__(self): self.pg=bool(DATABASE_URL)
    def connect(self):
        if self.pg:
            import psycopg
            from psycopg.rows import dict_row
            return psycopg.connect(DATABASE_URL.replace('postgres://','postgresql://',1), row_factory=dict_row)
        c=sqlite3.connect(DB_PATH); c.row_factory=sqlite3.Row; return c
    def execute(self,sql,params=(),fetch=False,many=False):
        c=self.connect(); cur=c.cursor()
        if self.pg: sql=sql.replace('?', '%s')
        cur.executemany(sql,params) if many else cur.execute(sql,params)
        rows=cur.fetchall() if fetch else None; c.commit(); cur.close(); c.close(); return rows

db=DB()

def rows(sql,p=()): return db.execute(sql,p,True)

def init_db():
    if db.pg:
        stmts=[
        'CREATE TABLE IF NOT EXISTS users (id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, full_name TEXT, role TEXT NOT NULL DEFAULT \'User\', active INTEGER NOT NULL DEFAULT 1, created_at TEXT, last_login TEXT)',
        'CREATE TABLE IF NOT EXISTS company_settings (key TEXT PRIMARY KEY, value TEXT)',
        'CREATE TABLE IF NOT EXISTS workers (id SERIAL PRIMARY KEY, name TEXT NOT NULL, bangla_name TEXT, department TEXT NOT NULL, designation TEXT, basic_salary DOUBLE PRECISION NOT NULL DEFAULT 0, ot_rate DOUBLE PRECISION NOT NULL DEFAULT 0, refreshment_bill DOUBLE PRECISION NOT NULL DEFAULT 0)',
        'CREATE TABLE IF NOT EXISTS attendance (id SERIAL PRIMARY KEY, worker_id INTEGER NOT NULL, month_year TEXT NOT NULL, present_days INTEGER DEFAULT 0, absent_days INTEGER DEFAULT 0, ot_hours DOUBLE PRECISION DEFAULT 0, UNIQUE(worker_id,month_year))',
        'CREATE TABLE IF NOT EXISTS worker_advances (id SERIAL PRIMARY KEY, worker_id INTEGER NOT NULL, month_year TEXT NOT NULL, advance_date TEXT NOT NULL, amount DOUBLE PRECISION NOT NULL, note TEXT)',
        'CREATE TABLE IF NOT EXISTS activity_log (id SERIAL PRIMARY KEY, username TEXT, action TEXT, log_time TEXT NOT NULL)']
    else:
        stmts=[
        'CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password_hash TEXT NOT NULL, full_name TEXT, role TEXT NOT NULL DEFAULT "User", active INTEGER NOT NULL DEFAULT 1, created_at TEXT, last_login TEXT)',
        'CREATE TABLE IF NOT EXISTS company_settings (key TEXT PRIMARY KEY, value TEXT)',
        'CREATE TABLE IF NOT EXISTS workers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, bangla_name TEXT, department TEXT NOT NULL, designation TEXT, basic_salary REAL NOT NULL DEFAULT 0, ot_rate REAL NOT NULL DEFAULT 0, refreshment_bill REAL NOT NULL DEFAULT 0)',
        'CREATE TABLE IF NOT EXISTS attendance (id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id INTEGER NOT NULL, month_year TEXT NOT NULL, present_days INTEGER DEFAULT 0, absent_days INTEGER DEFAULT 0, ot_hours REAL DEFAULT 0, UNIQUE(worker_id,month_year))',
        'CREATE TABLE IF NOT EXISTS worker_advances (id INTEGER PRIMARY KEY AUTOINCREMENT, worker_id INTEGER NOT NULL, month_year TEXT NOT NULL, advance_date TEXT NOT NULL, amount REAL NOT NULL, note TEXT)',
        'CREATE TABLE IF NOT EXISTS activity_log (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT, action TEXT, log_time TEXT NOT NULL)']
    for s in stmts: db.execute(s)
    # Migration for older databases
    for s in ['ALTER TABLE workers ADD COLUMN bangla_name TEXT','ALTER TABLE users ADD COLUMN password_hash TEXT','ALTER TABLE users ADD COLUMN full_name TEXT','ALTER TABLE users ADD COLUMN role TEXT','ALTER TABLE users ADD COLUMN active INTEGER','ALTER TABLE users ADD COLUMN created_at TEXT','ALTER TABLE users ADD COLUMN last_login TEXT']:
        try: db.execute(s)
        except Exception: pass
    if not rows('SELECT id FROM users WHERE username=?',('admin',)):
        db.execute('INSERT INTO users(username,password_hash,full_name,role,active,created_at) VALUES(?,?,?,?,?,?)',('admin',sha('admin123'),'System Administrator','Administrator',1,datetime.now().isoformat()))
    defaults={'company_name':'Reedoy Textile Dyeing Printing & Finishing','company_address':'','company_phone':'','company_email':''}
    for k,v in defaults.items():
        if not rows('SELECT key FROM company_settings WHERE key=?',(k,)): db.execute('INSERT INTO company_settings(key,value) VALUES(?,?)',(k,v))

def setting(k):
    r=rows('SELECT value FROM company_settings WHERE key=?',(k,)); return r[0]['value'] if r else ''
def log(action):
    if session.get('user'): db.execute('INSERT INTO activity_log(username,action,log_time) VALUES(?,?,?)',(session['user'],action,datetime.now().isoformat()))
def login_required(f):
    @wraps(f)
    def w(*a,**kw):
        if not session.get('user'): return redirect(url_for('login'))
        return f(*a,**kw)
    return w
def admin_required(f):
    @wraps(f)
    def w(*a,**kw):
        if not session.get('user'): return redirect(url_for('login'))
        if session.get('role')!='Administrator': abort(403)
        return f(*a,**kw)
    return w

def calc(w,a,adv):
    present=float(a.get('present_days',0) or 0); ot=float(a.get('ot_hours',0) or 0)
    gross=(float(w['basic_salary'])/30.0)*present + ot*float(w['ot_rate']) + float(w.get('refreshment_bill',0) or 0)
    return gross,float(adv or 0)

def payroll_rows(month,search='',department='All Departments'):
    sql='''SELECT w.*, COALESCE(a.present_days,0) present_days, COALESCE(a.absent_days,0) absent_days, COALESCE(a.ot_hours,0) ot_hours, COALESCE((SELECT SUM(amount) FROM worker_advances x WHERE x.worker_id=w.id AND x.month_year=?),0) advance FROM workers w LEFT JOIN attendance a ON a.worker_id=w.id AND a.month_year=? WHERE (w.name LIKE ? OR COALESCE(w.bangla_name,'') LIKE ?) AND (?='All Departments' OR w.department=?) ORDER BY w.department,w.id'''
    data=[]
    for r in rows(sql,(month,month,'%'+search+'%','%'+search+'%',department,department)):
        d=dict(r); g,adv=calc(d,d,d['advance']); d.update(gross=g,net=g-adv); data.append(d)
    return data

@app.context_processor
def common(): return {'company':setting('company_name') or 'Reedoy Textile','today':date.today().isoformat(),'departments':list(DEPT_BN.keys())[1:]}

@app.route('/set-language/<lang>')
def set_language(lang): session['lang']='bn' if lang=='bn' else 'en'; return redirect(request.referrer or url_for('dashboard'))

@app.route('/login',methods=['GET','POST'])
def login():
    if request.method=='POST':
        u=request.form.get('username','').strip(); p=request.form.get('password','')
        r=rows('SELECT * FROM users WHERE username=? AND active=1',(u,))
        if r and r[0]['password_hash']==sha(p):
            session.update(user=u,role=r[0]['role'],lang=session.get('lang','en')); db.execute('UPDATE users SET last_login=? WHERE username=?',(datetime.now().isoformat(),u)); log('Login'); return redirect(url_for('dashboard'))
        flash('Invalid username or password','danger')
    return render_template('login.html')
@app.route('/logout')
def logout(): session.clear(); return redirect(url_for('login'))

@app.route('/')
@login_required
def dashboard():
    return render_template('dashboard.html',workers=rows('SELECT COUNT(*) c FROM workers')[0]['c'],advances=rows('SELECT COALESCE(SUM(amount),0) s FROM worker_advances')[0]['s'],departments=rows('SELECT COUNT(DISTINCT department) c FROM workers')[0]['c'],month=datetime.now().strftime('%Y-%m'))

@app.route('/workers',methods=['GET','POST'])
@login_required
def workers():
    if request.method=='POST':
        f=request.form; wid=f.get('id')
        vals=(f['name'].strip(),f.get('bangla_name','').strip(),f['department'],f.get('designation','').strip(),float(f.get('basic_salary') or 0),float(f.get('ot_rate') or 0),float(f.get('refreshment_bill') or 0))
        if wid:
            db.execute('UPDATE workers SET name=?,bangla_name=?,department=?,designation=?,basic_salary=?,ot_rate=?,refreshment_bill=? WHERE id=?',vals+(int(wid),)); log('Updated worker: '+f['name']); flash('Worker updated','success')
        else:
            db.execute('INSERT INTO workers(name,bangla_name,department,designation,basic_salary,ot_rate,refreshment_bill) VALUES(?,?,?,?,?,?,?)',vals); log('Added worker: '+f['name']); flash('Worker saved','success')
        return redirect(url_for('workers'))
    q=request.args.get('q','').strip(); edit_id=request.args.get('edit'); edit=None
    if edit_id:
        rr=rows('SELECT * FROM workers WHERE id=?',(edit_id,)); edit=rr[0] if rr else None
    data=rows('SELECT * FROM workers WHERE name LIKE ? OR COALESCE(bangla_name,\'\') LIKE ? ORDER BY id DESC',('%'+q+'%','%'+q+'%'))
    return render_template('workers.html',workers=data,q=q,edit=edit)
@app.route('/workers/delete/<int:wid>')
@login_required
def delete_worker(wid):
    db.execute('DELETE FROM attendance WHERE worker_id=?',(wid,)); db.execute('DELETE FROM worker_advances WHERE worker_id=?',(wid,)); db.execute('DELETE FROM workers WHERE id=?',(wid,)); log(f'Deleted worker {wid}'); flash('Worker deleted','success'); return redirect(url_for('workers'))

@app.route('/attendance',methods=['GET','POST'])
@login_required
def attendance():
    month=request.values.get('month') or datetime.now().strftime('%Y-%m')
    if request.method=='POST':
        for w in rows('SELECT id FROM workers'):
            pid=str(w['id']); p=int(request.form.get('present_'+pid,0) or 0); a=int(request.form.get('absent_'+pid,0) or 0); ot=float(request.form.get('ot_'+pid,0) or 0)
            if db.pg: db.execute('INSERT INTO attendance(worker_id,month_year,present_days,absent_days,ot_hours) VALUES(?,?,?,?,?) ON CONFLICT(worker_id,month_year) DO UPDATE SET present_days=EXCLUDED.present_days,absent_days=EXCLUDED.absent_days,ot_hours=EXCLUDED.ot_hours',(w['id'],month,p,a,ot))
            else: db.execute('INSERT INTO attendance(worker_id,month_year,present_days,absent_days,ot_hours) VALUES(?,?,?,?,?) ON CONFLICT(worker_id,month_year) DO UPDATE SET present_days=excluded.present_days,absent_days=excluded.absent_days,ot_hours=excluded.ot_hours',(w['id'],month,p,a,ot))
        log('Updated attendance '+month); flash('Attendance saved','success')
    data=rows('SELECT w.*,COALESCE(a.present_days,0) present_days,COALESCE(a.absent_days,0) absent_days,COALESCE(a.ot_hours,0) ot_hours FROM workers w LEFT JOIN attendance a ON a.worker_id=w.id AND a.month_year=? ORDER BY w.id',(month,))
    return render_template('attendance.html',data=data,month=month)

@app.route('/advances',methods=['GET','POST'])
@login_required
def advances():
    if request.method=='POST':
        f=request.form; db.execute('INSERT INTO worker_advances(worker_id,month_year,advance_date,amount,note) VALUES(?,?,?,?,?)',(int(f['worker_id']),f['month_year'],f['advance_date'],float(f['amount']),f.get('note',''))); log('Added advance'); flash('Advance saved','success'); return redirect(url_for('advances'))
    q=request.args.get('q','').strip(); data=rows('SELECT a.*,w.name,w.bangla_name FROM worker_advances a LEFT JOIN workers w ON w.id=a.worker_id WHERE w.name LIKE ? OR COALESCE(w.bangla_name,\'\') LIKE ? ORDER BY a.id DESC',('%'+q+'%','%'+q+'%'))
    return render_template('advances.html',advances=data,workers=rows('SELECT * FROM workers ORDER BY name'),q=q)
@app.route('/advances/delete/<int:aid>')
@login_required
def delete_advance(aid): db.execute('DELETE FROM worker_advances WHERE id=?',(aid,)); log(f'Deleted advance {aid}'); return redirect(url_for('advances'))

@app.route('/report')
@login_required
def report():
    month=request.args.get('month') or datetime.now().strftime('%Y-%m'); q=request.args.get('q','').strip(); data=payroll_rows(month,q); return render_template('report.html',data=data,month=month,q=q)
@app.route('/report.csv')
@login_required
def report_csv():
    month=request.args.get('month') or datetime.now().strftime('%Y-%m'); data=payroll_rows(month)
    out=io.StringIO(); wr=csv.writer(out); wr.writerow(['Worker ID','Name','Bangla Name','Department','Designation','Present','Absent','OT Hours','Basic','Gross','Advance','Net Payable'])
    for x in data: wr.writerow([x['id'],x['name'],x.get('bangla_name',''),x['department'],x['designation'],x['present_days'],x['absent_days'],x['ot_hours'],x['basic_salary'],round(x['gross'],2),round(x['advance'],2),round(x['net'],2)])
    b=io.BytesIO(out.getvalue().encode('utf-8-sig')); b.seek(0); return send_file(b,as_attachment=True,download_name=f'salary_{month}.csv',mimetype='text/csv')
@app.route('/report.xlsx')
@login_required
def report_xlsx():
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font,Alignment
    except Exception: abort(500)
    month=request.args.get('month') or datetime.now().strftime('%Y-%m'); data=payroll_rows(month); wb=Workbook(); ws=wb.active; ws.title='Salary Report'
    ws.append([setting('company_name'),None,None,None,None,None,None,None,None,None,None,None]); ws.append([f'Salary Report - {month}']); ws.append(['ID','Name','Bangla Name','Department','Designation','Present','Absent','OT Hours','Basic','Gross','Advance','Net Payable'])
    for c in ws[1]: c.font=Font(bold=True,size=14)
    for x in data: ws.append([x['id'],x['name'],x.get('bangla_name',''),dept(x['department']),x['designation'],x['present_days'],x['absent_days'],x['ot_hours'],x['basic_salary'],round(x['gross'],2),round(x['advance'],2),round(x['net'],2)])
    for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=min(max(max(len(str(c.value or '')) for c in col)+2,10),28)
    bio=io.BytesIO(); wb.save(bio); bio.seek(0); return send_file(bio,as_attachment=True,download_name=f'salary_{month}.xlsx',mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
@app.route('/report.pdf')
@login_required
def report_pdf():
    try:
        from reportlab.lib.pagesizes import A4,landscape
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.pdfbase import pdfmetrics
    except Exception: abort(500)
    month=request.args.get('month') or datetime.now().strftime('%Y-%m'); data=payroll_rows(month); bio=io.BytesIO(); c=canvas.Canvas(bio,pagesize=landscape(A4)); w,h=landscape(A4)
    # Latin PDF by default; browser report remains Unicode-capable.
    c.setFont('Helvetica-Bold',14); c.drawString(30,h-30,setting('company_name') or 'Reedoy Textile'); c.setFont('Helvetica',10); c.drawString(30,h-48,f'Salary Report - {month}')
    headers=['ID','Name','Department','Present','Absent','OT','Gross','Advance','Net']; xs=[30,60,180,310,360,410,450,520,590]; y=h-75
    c.setFont('Helvetica-Bold',8)
    for i,t in enumerate(headers): c.drawString(xs[i],y,t)
    c.setFont('Helvetica',8); y-=15
    for x in data:
        vals=[x['id'],x['name'],x['department'],x['present_days'],x['absent_days'],x['ot_hours'],round(x['gross'],2),round(x['advance'],2),round(x['net'],2)]
        for i,v in enumerate(vals): c.drawString(xs[i],y,str(v)[:25])
        y-=13
        if y<30: c.showPage(); y=h-40; c.setFont('Helvetica',8)
    c.save(); bio.seek(0); return send_file(bio,as_attachment=True,download_name=f'salary_{month}.pdf',mimetype='application/pdf')

@app.route('/department')
@login_required
def department():
    month=request.args.get('month') or datetime.now().strftime('%Y-%m'); selected=request.args.get('department','All Departments'); data=payroll_rows(month,'',selected); return render_template('department.html',data=data,month=month,selected=selected)

@app.route('/payslip/<int:wid>')
@login_required
def payslip(wid):
    month=request.args.get('month') or datetime.now().strftime('%Y-%m'); data=payroll_rows(month); x=next((z for z in data if z['id']==wid),None)
    if not x: abort(404)
    return render_template('payslip.html',w=x,month=month,gross=x['gross'],net=x['net'])

@app.route('/settings',methods=['GET','POST'])
@admin_required
def settings():
    if request.method=='POST':
        for k in ['company_name','company_address','company_phone','company_email']: db.execute('INSERT INTO company_settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',(k,request.form.get(k,'')))
        log('Updated company settings'); flash('Settings saved','success'); return redirect(url_for('settings'))
    return render_template('settings.html',settings={k:setting(k) for k in ['company_name','company_address','company_phone','company_email']})

@app.route('/users',methods=['GET','POST'])
@admin_required
def users():
    if request.method=='POST':
        f=request.form; username=f['username'].strip(); password=f['password']; role=f.get('role','User')
        if not username or not password: flash('Username and password required','danger')
        else:
            try: db.execute('INSERT INTO users(username,password_hash,full_name,role,active,created_at) VALUES(?,?,?,?,?,?)',(username,sha(password),f.get('full_name',''),role,1,datetime.now().isoformat())); log('Created user '+username); flash('User created','success')
            except Exception: flash('Username already exists','danger')
    return render_template('users.html',users=rows('SELECT id,username,full_name,role,active,created_at,last_login FROM users ORDER BY id'))
@app.route('/users/toggle/<int:uid>')
@admin_required
def toggle_user(uid): db.execute('UPDATE users SET active=CASE WHEN active=1 THEN 0 ELSE 1 END WHERE id=?',(uid,)); log(f'Toggled user {uid}'); return redirect(url_for('users'))
@app.route('/users/password/<int:uid>',methods=['POST'])
@admin_required
def change_password(uid): db.execute('UPDATE users SET password_hash=? WHERE id=?',(sha(request.form.get('password','')),uid)); log(f'Changed password for user {uid}'); return redirect(url_for('users'))

@app.route('/activity')
@admin_required
def activity(): return render_template('activity.html',logs=rows('SELECT * FROM activity_log ORDER BY id DESC LIMIT 500'))

init_db()
if __name__=='__main__': app.run(host='0.0.0.0',port=int(os.environ.get('PORT',5000)),debug=False)
