from flask import (
    Flask, render_template, redirect, url_for,
    request, flash, jsonify, send_file, make_response, abort
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user,
    login_required, current_user
)
from werkzeug.security import generate_password_hash, check_password_hash
from datetime import datetime, date, timedelta, time as dt_time
import os, smtplib, ssl, io, csv
from email.message import EmailMessage
from sqlalchemy import text
import xlsxwriter  # Excel export (in-memory, safe on Render)

# =========================================================
# App & DB config
# =========================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ChangeThisSecret123!")
app.config["TEMPLATES_AUTO_RELOAD"] = True

db_url = os.environ.get("DATABASE_URL", "sqlite:///leave_system.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql+psycopg://", 1)
elif db_url.startswith("postgresql://") and "+psycopg" not in db_url:
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}

db = SQLAlchemy(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"

@app.after_request
def add_no_cache_headers(resp):
    if resp.mimetype == "text/html":
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return resp

# =========================================================
# Email settings (env vars)
# =========================================================
MAIL_HOST = os.environ.get("MAIL_HOST", "")
MAIL_PORT = int(os.environ.get("MAIL_PORT", "587"))
MAIL_USER = os.environ.get("MAIL_USER", "")
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD", "")
MAIL_USE_TLS = os.environ.get("MAIL_USE_TLS", "true").lower() == "true"
MAIL_FROM = os.environ.get("MAIL_FROM", MAIL_USER or "no-reply@example.com")

ADMIN_EMAILS_ENV = [e.strip() for e in os.environ.get("ADMIN_EMAILS", "").split(",") if e.strip()]

def send_email(to_addrs, subject, body):
    try:
        if not to_addrs:
            app.logger.warning("Email skipped: no recipients provided.")
            return False, "No recipients"
        if isinstance(to_addrs, str):
            to_addrs = [to_addrs]
        to_addrs = [a for a in to_addrs if a]

        if not MAIL_HOST or not MAIL_FROM:
            app.logger.warning("Email skipped: MAIL_HOST or MAIL_FROM not set.")
            return False, "SMTP not configured"

        msg = EmailMessage()
        msg["From"] = MAIL_FROM
        msg["To"] = ", ".join(to_addrs)
        msg["Subject"] = subject
        msg.set_content(body)

        if MAIL_USE_TLS:
            context = ssl.create_default_context()
            with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as server:
                server.ehlo(); server.starttls(context=context); server.ehlo()
                if MAIL_USER: server.login(MAIL_USER, MAIL_PASSWORD)
                server.send_message(msg)
        else:
            with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as server:
                if MAIL_USER: server.login(MAIL_USER, MAIL_PASSWORD)
                server.send_message(msg)
        return True, "Sent"
    except Exception as e:
        app.logger.error(f"Email send failed: {e}")
        return False, f"Failed: {e}"

# =========================================================
# Models & constants
# =========================================================
class Role:
    admin = "admin"
    staff = "faculty_staff"

class RequestStatus:
    pending = "Pending"
    approved = "Approved"
    disapproved = "Disapproved"
    cancelled = "Cancelled"

class RequestMode:
    hourly = "hourly"
    daily = "daily"
    halfday = "halfday"  # half-day option

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default=Role.staff, nullable=False)
    hours_balance = db.Column(db.Float, default=160.0, nullable=False)
    email = db.Column(db.String(255))
    staff_name = db.Column(db.String(150))

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    kind = db.Column(db.String(20), default="annual", nullable=False)
    mode = db.Column(db.String(10), default=RequestMode.hourly, nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.String(5))
    end_time   = db.Column(db.String(5))
    hours = db.Column(db.Float, nullable=False)
    reason = db.Column(db.String(500), default="")
    status = db.Column(db.String(20), default=RequestStatus.pending, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    decided_at = db.Column(db.DateTime)

    is_school_related = db.Column(db.Boolean, default=False, nullable=False)
    substitute = db.Column(db.String(120))

    user = db.relationship("User", backref="leave_requests", lazy="joined")

class SubAssignment(db.Model):
    __tablename__ = "sub_assignment"
    id = db.Column(db.Integer, primary_key=True)
    request_id = db.Column(db.Integer, db.ForeignKey("leave_request.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    hours = db.Column(db.Float, nullable=False, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# =========================================================
# Helpers
# =========================================================
WORKDAY_HOURS = float(os.environ.get("WORKDAY_HOURS", "8"))
HOLIDAYS: set[date] = set()

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in HOLIDAYS

def workdays_between(start: date, end: date) -> int:
    if end < start: start, end = end, start
    n, cur = 0, start
    while cur <= end:
        if is_workday(cur): n += 1
        cur = cur + timedelta(days=1)
    return n

def parse_quarter_time(s: str) -> dt_time | None:
    try:
        hh, mm = s.split(":"); hh_i = int(hh); mm_i = int(mm)
        if 0 <= hh_i <= 23 and mm_i in (0, 15, 30, 45): return dt_time(hh_i, mm_i)
    except Exception:
        return None
    return None

def interval_hours(t1: dt_time, t2: dt_time) -> float:
    dt1 = datetime.combine(date.today(), t1)
    dt2 = datetime.combine(date.today(), t2)
    if dt2 < dt1: dt1, dt2 = dt2, dt1
    return (dt2 - dt1).total_seconds() / 3600.0

def _column_exists(table_name: str, column_name: str) -> bool:
    bind = db.engine; dialect = bind.dialect.name
    if dialect == "sqlite":
        res = db.session.execute(text(f"PRAGMA table_info({table_name})")).fetchall()
        return any(row[1] == column_name for row in res)
    q = text("""SELECT 1 FROM information_schema.columns
                WHERE table_name = :t AND column_name = :c LIMIT 1""")
    return db.session.execute(q, {"t": table_name, "c": column_name}).first() is not None

def ensure_db():
    db.create_all()
    try:
        if db.engine.dialect.name == "sqlite":
            if not _column_exists("leave_request", "start_time"):
                db.session.execute(text("ALTER TABLE leave_request ADD COLUMN start_time VARCHAR(5)"))
            if not _column_exists("leave_request", "end_time"):
                db.session.execute(text("ALTER TABLE leave_request ADD COLUMN end_time VARCHAR(5)"))
            if not _column_exists("leave_request", "is_school_related"):
                db.session.execute(text("ALTER TABLE leave_request ADD COLUMN is_school_related BOOLEAN DEFAULT 0 NOT NULL"))
            if not _column_exists("leave_request", "substitute"):
                db.session.execute(text("ALTER TABLE leave_request ADD COLUMN substitute VARCHAR(120)"))
            if not _column_exists("user", "staff_name"):
                db.session.execute(text("ALTER TABLE \"user\" ADD COLUMN staff_name VARCHAR(150)"))
            db.session.commit()
        else:
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS start_time VARCHAR(5)"))
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS end_time VARCHAR(5)"))
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS is_school_related BOOLEAN NOT NULL DEFAULT FALSE"))
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS substitute VARCHAR(120)"))
            db.session.execute(text("ALTER TABLE \"user\" ADD COLUMN IF NOT EXISTS staff_name VARCHAR(150)"))
            db.session.commit()
    except Exception:
        db.session.rollback()

    if User.query.count() == 0:
        bootstrap_username = os.environ.get("BOOTSTRAP_ADMIN_USERNAME", "mc-admin")
        bootstrap_password = os.environ.get("BOOTSTRAP_ADMIN_PASSWORD", "RWAadmin2")
        bootstrap_email = os.environ.get("BOOTSTRAP_ADMIN_EMAIL", (ADMIN_EMAILS_ENV[0] if ADMIN_EMAILS_ENV else ""))
        db.session.add(User(
            username=bootstrap_username,
            password_hash=generate_password_hash(bootstrap_password),
            role=Role.admin, hours_balance=160.0,
            email=bootstrap_email or None, staff_name="Administrator"
        ))
        db.session.commit()

with app.app_context():
    ensure_db()

def admin_emails() -> list[str]:
    env_list = ADMIN_EMAILS_ENV[:]
    user_list = [u.email for u in User.query.filter_by(role=Role.admin).all() if u.email]
    seen = set(); out = []
    for e in env_list + user_list:
        if e and e not in seen: out.append(e); seen.add(e)
    return out

def assemble_hhmm_from_12h(h_str: str, m_str: str, ampm: str) -> str | None:
    try:
        if not h_str or not m_str or not ampm: return None
        h = int(h_str); m = int(m_str)
        if h < 1 or h > 12 or m not in (0, 15, 30, 45): return None
        ampm = ampm.upper()
        if ampm == "PM" and h != 12: h += 12
        if ampm == "AM" and h == 12: h = 0
        return f"{h:02d}:{m:02d}"
    except Exception:
        return None

@app.template_filter("h12")
def h12(time_str: str) -> str:
    try:
        if not time_str: return ""
        hh, mm = time_str.split(":"); h = int(hh)
        suffix = "AM" if h < 12 else "PM"
        h12 = h % 12; h12 = 12 if h12 == 0 else h12
        return f"{h12}:{mm} {suffix}"
    except Exception:
        return time_str

def _filtered_requests_for(current_user_is_admin: bool):
    status = request.args.get("status", "all").strip().lower()
    start_s = request.args.get("start", "").strip()
    end_s = request.args.get("end", "").strip()

    q = LeaveRequest.query
    if not current_user_is_admin:
        q = q.filter_by(user_id=current_user.id)
    if status and status != "all":
        q = q.filter_by(status=status.capitalize())

    def parse_date(s):
        try: return datetime.strptime(s, "%Y-%m-%d").date()
        except Exception: return None

    sd = parse_date(start_s); ed = parse_date(end_s)
    if sd: q = q.filter(LeaveRequest.start_date >= sd)
    if ed: q = q.filter(LeaveRequest.end_date <= ed)
    return q.order_by(LeaveRequest.created_at.desc())

# =========================================================
# Nav + globals
# =========================================================
@app.context_processor
def inject_globals():
    class NAV: pass
    nav = NAV()
    if current_user.is_authenticated:
        nav.dashboard = url_for("dashboard")
        nav.my_requests = url_for("my_requests")
        nav.team_calendar = url_for("calendar")
        nav.new_request = url_for("new_request")
        nav.admin = url_for("admin_hub") if current_user.role == Role.admin else None
        nav.account = url_for("update_password")
        nav.logout = url_for("logout")
    else:
        login_url = url_for("login")
        nav.dashboard = nav.my_requests = nav.team_calendar = nav.new_request = nav.admin = nav.account = nav.logout = login_url
    return {"current_year": datetime.utcnow().year, "NAV": nav}

# =========================================================
# Routes
# =========================================================
@app.get("/health")
def health():
    return "ok", 200

@app.route("/")
def home():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter(User.username.ilike(username)).first()
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            flash("Logged in.", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "danger")
    return render_template("login.html", title="Login")

@app.get("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

@app.get("/dashboard")
@login_required
def dashboard():
    recent = (LeaveRequest.query.filter_by(user_id=current_user.id)
              .order_by(LeaveRequest.created_at.desc()).limit(10).all())
    return render_template("dashboard.html", title="Dashboard",
                           me=current_user, workday=WORKDAY_HOURS, recent=recent)

# ---------- Admin HUB ----------
@app.get("/admin")
@login_required
def admin_hub():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))
    pending = (LeaveRequest.query.filter_by(status=RequestStatus.pending)
               .order_by(LeaveRequest.created_at.desc()).all())
    return render_template("admin.html", title="Admin", pending=pending)

@app.get("/admin/email-test")
@login_required
def admin_email_test():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))
    recipients = []
    if current_user.email: recipients.append(current_user.email)
    recipients += admin_emails()
    seen=set(); recipients=[r for r in recipients if r and not (r in seen or seen.add(r))]
    ok, msg = send_email(recipients, "RWA Leave System – Test Email",
                         f"Test at {datetime.utcnow().isoformat()}Z\nFrom {MAIL_FROM}\n")
    flash("Test email sent." if ok else f"Test email failed: {msg}", "success" if ok else "danger")
    return redirect(url_for("admin_hub"))

# ---------- New Request ----------
@app.route("/request/new", methods=["GET", "POST"])
@login_required
def new_request():
    minutes_opts = ["00", "15", "30", "45"]
    hours12_opts = [str(i) for i in range(1, 13)]

    if request.method == "POST":
        mode = request.form.get("mode", RequestMode.hourly)
        kind = request.form.get("kind", "annual")
        reason = request.form.get("reason", "")
        is_school = bool(request.form.get("school_related"))

        try:
            sd = datetime.strptime(request.form["start_date"], "%Y-%m-%d").date()
            ed = datetime.strptime(request.form["end_date"], "%Y-%m-%d").date()
        except Exception:
            flash("Invalid dates.", "warning")
            return render_template("new_request.html", title="New Request",
                                   workday=WORKDAY_HOURS,
                                   minutes_opts=minutes_opts, hours12_opts=hours12_opts)

        capacity_hours = workdays_between(sd, ed) * WORKDAY_HOURS
        if capacity_hours <= 0:
            flash("No working days in that range.", "warning")
            return render_template("new_request.html", title="New Request",
                                   workday=WORKDAY_HOURS,
                                   minutes_opts=minutes_opts, hours12_opts=hours12_opts)

        hours = 0.0
        if mode == RequestMode.hourly:
            hours_str = (request.form.get("hours") or "").strip()
            if hours_str:
                try:
                    hours = float(hours_str)
                except Exception:
                    hours = 0.0
            else:
                st_s = assemble_hhmm_from_12h(
                    request.form.get("start_h",""), request.form.get("start_m",""), request.form.get("start_ampm","")
                ) or (request.form.get("start_time") or "").strip()
                et_s = assemble_hhmm_from_12h(
                    request.form.get("end_h",""), request.form.get("end_m",""), request.form.get("end_ampm","")
                ) or (request.form.get("end_time") or "").strip()
                st = parse_quarter_time(st_s) if st_s else None
                et = parse_quarter_time(et_s) if et_s else None
                if st and et and sd == ed:
                    hours = interval_hours(st, et)
                else:
                    hours = 0.0
            st_val = assemble_hhmm_from_12h(request.form.get("start_h",""), request.form.get("start_m",""), request.form.get("start_ampm",""))
            et_val = assemble_hhmm_from_12h(request.form.get("end_h",""), request.form.get("end_m",""), request.form.get("end_ampm",""))
            st_s = st_val or (request.form.get("start_time") or "").strip()
            et_s = et_val or (request.form.get("end_time") or "").strip()

        elif mode == RequestMode.daily:
            wd = workdays_between(sd, ed)
            hours = wd * WORKDAY_HOURS
            st_s = et_s = None
        elif mode == RequestMode.halfday:
            wd = workdays_between(sd, ed)
            hours = wd * (WORKDAY_HOURS / 2.0)
            st_s = et_s = None
        else:
            st_s = et_s = None

        if hours <= 0:
            flash("Requested hours must be greater than zero.", "warning")
            return render_template("new_request.html", title="New Request",
                                   workday=WORKDAY_HOURS,
                                   minutes_opts=minutes_opts, hours12_opts=hours12_opts)

        if mode not in (RequestMode.hourly,) and hours > capacity_hours:
            flash(f"Requested {hours:.2f} exceeds capacity {capacity_hours:.2f} for that range.", "warning")
            return render_template("new_request.html", title="New Request",
                                   workday=WORKDAY_HOURS,
                                   minutes_opts=minutes_opts, hours12_opts=hours12_opts)

        req = LeaveRequest(
            user_id=current_user.id,
            kind=kind,
            mode=mode,
            start_date=sd,
            end_date=ed,
            start_time=st_s or None,
            end_time=et_s or None,
            hours=hours,
            reason=reason,
            is_school_related=is_school
        )
        db.session.add(req); db.session.commit()

        subj = "New Leave Request Submitted"
        body = (
            f"User: {current_user.username}\n"
            f"Kind: {kind}\nMode: {mode}\nHours: {hours:.2f}\n"
            f"Dates: {sd} to {ed}\n"
            f"Times: {req.start_time or '-'} to {req.end_time or '-'}\n"
            f"School-related: {'Yes' if is_school else 'No'}\n"
            f"Reason: {reason or '(none)'}\n"
            f"Status: {req.status}\n"
        )
        ok, emsg = send_email(admin_emails(), subj, body)
        if not ok:
            flash(f"Notice: admin email not sent ({emsg}). Check SMTP settings.", "warning")

        flash("Request submitted.", "success")
        return redirect(url_for("my_requests"))

    return render_template("new_request.html", title="New Request",
                           workday=WORKDAY_HOURS,
                           minutes_opts=minutes_opts, hours12_opts=hours12_opts)

# ---------- Requests list ----------
@app.get("/requests")
@login_required
def my_requests():
    q = _filtered_requests_for(current_user.role == Role.admin)
    reqs = q.all()
    return render_template("requests.html", title="Requests", reqs=reqs, me=current_user,
                           is_admin=(current_user.role == Role.admin),
                           status=request.args.get("status", "all"),
                           start=request.args.get("start", ""),
                           end=request.args.get("end", ""))

# ---------- School-related toggles ----------
@app.post("/requests/<int:req_id>/school")
@login_required
def mark_school_related(req_id):
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Only pending requests can be changed.", "warning"); return redirect(url_for("my_requests"))
    if r.user_id != current_user.id and current_user.role != Role.admin:
        flash("Not allowed.", "danger"); return redirect(url_for("my_requests"))
    r.is_school_related = True; db.session.commit()
    flash("Marked as school-related (no balance deduction on approval).", "success")
    return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/unschool")
@login_required
def unmark_school_related(req_id):
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Only pending requests can be changed.", "warning"); return redirect(url_for("my_requests"))
    if r.user_id != current_user.id and current_user.role != Role.admin:
        flash("Not allowed.", "danger"); return redirect(url_for("my_requests"))
    r.is_school_related = False; db.session.commit()
    flash("Removed school-related flag.", "success")
    return redirect(request.referrer or url_for("my_requests"))

# ---------- Approvals / Disapprovals / Cancel ----------
@app.post("/requests/<int:req_id>/approve")
@login_required
def approve(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning"); return redirect(url_for("my_requests"))
    u = User.query.get(r.user_id)
    if not r.is_school_related:
        u.hours_balance = float(u.hours_balance or 0.0) - float(r.hours or 0.0)
    r.status = RequestStatus.approved; r.decided_at = datetime.utcnow()
    db.session.commit()

    subj = "Leave Request Approved"
    body = (
        f"Hello {u.username},\n\nYour leave request has been APPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
        f"Remaining balance: {u.hours_balance:.2f} hours\n"
    )
    ok, emsg = send_email([u.email] + admin_emails(), subj, body)
    if not ok: flash(f"Notice: approval email not sent ({emsg}).", "warning")
    flash("Approved.", "success"); return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/disapprove")
@login_required
def disapprove(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning"); return redirect(url_for("my_requests"))
    r.status = RequestStatus.disapproved; r.decided_at = datetime.utcnow()
    db.session.commit()

    u = User.query.get(r.user_id)
    ok, emsg = send_email([u.email] + admin_emails(), "Leave Request Disapproved",
                          f"Hello {u.username},\n\nYour leave request has been DISAPPROVED.\n")
    if not ok: flash(f"Notice: disapproval email not sent ({emsg}).", "warning")
    flash("Disapproved.", "info"); return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/cancel")
@login_required
def cancel(req_id):
    r = LeaveRequest.query.get_or_404(req_id)
    if r.user_id != current_user.id and current_user.role != Role.admin:
        flash("Not allowed.", "danger"); return redirect(url_for("my_requests"))
    u = User.query.get(r.user_id)
    if r.status == RequestStatus.approved and not r.is_school_related:
        u.hours_balance = float(u.hours_balance or 0.0) + float(r.hours or 0.0)
    r.status = RequestStatus.cancelled; r.decided_at = datetime.utcnow()
    db.session.commit()

    recipients = admin_emails()
    if u.email: recipients = [u.email] + recipients
    ok, emsg = send_email(recipients, "Leave Request Cancelled",
                          f"User {u.username} cancelled a request.")
    if not ok: flash(f"Notice: cancel email not sent ({emsg}).", "warning")
    flash("Cancelled.", "secondary"); return redirect(request.referrer or url_for("my_requests"))

# ---------- Manage Users ----------
@app.route("/admin/users", methods=["GET"])
@login_required
def manage_users():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("dashboard"))
    qtxt = request.args.get("q", "").strip()
    query = User.query
    if qtxt: query = query.filter(User.username.ilike(f"%{qtxt}%"))
    users = query.order_by(User.username.asc()).all()
    return render_template("manage_users.html", title="Manage Users", users=users, q=qtxt)

@app.post("/admin/users/create")
@login_required
def admin_create_user():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("manage_users"))

    username = (request.form.get("username") or "").strip()
    staff_name = (request.form.get("staff_name") or "").strip()
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or Role.staff).strip()
    hours_str = (request.form.get("hours_balance") or "").strip()
    pw = (request.form.get("password") or "").strip()

    if not username or not pw:
        flash("Username and password are required.", "warning")
        return redirect(url_for("manage_users"))
    if User.query.filter(User.username.ilike(username)).first():
        flash("Username already exists.", "danger")
        return redirect(url_for("manage_users"))

    try: hours_balance = float(hours_str) if hours_str else 160.0
    except Exception: hours_balance = 160.0

    user = User(
        username=username,
        password_hash=generate_password_hash(pw),
        role=role if role in (Role.admin, Role.staff) else Role.staff,
        hours_balance=hours_balance,
        email=email or None,
        staff_name=staff_name or None
    )
    db.session.add(user); db.session.commit()
    flash(f"User '{username}' created.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/update")
@login_required
def admin_update_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or "").strip()
    hb = (request.form.get("hours_balance") or "").strip()
    staff_name = (request.form.get("staff_name") or "").strip()

    u.email = email or None
    if role in (Role.admin, Role.staff) and role: u.role = role
    try:
        if hb != "": u.hours_balance = float(hb)
    except Exception:
        flash("Invalid hours_balance value.", "warning")
    if staff_name != "": u.staff_name = staff_name

    db.session.commit()
    flash(f"Updated {u.username}.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/reset")
@login_required
def admin_reset_password(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("manage_users"))
    new_pw = (request.form.get("new_password") or "").strip()
    if not new_pw:
        flash("Password cannot be empty.", "warning")
        return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    u.password_hash = generate_password_hash(new_pw)
    db.session.commit()
    flash(f"Password updated for {u.username}.", "success")
    return redirect(url_for("manage_users"))

# ---------- Self-service password change ----------
@app.route("/account/password", methods=["GET", "POST"])
@login_required
def update_password():
    if request.method == "POST":
        cur = request.form.get("current_password", "")
        new = (request.form.get("new_password") or "").strip()
        if not check_password_hash(current_user.password_hash, cur):
            flash("Current password is incorrect.", "danger")
        elif not new:
            flash("New password cannot be empty.", "warning")
        else:
            current_user.password_hash = generate_password_hash(new)
            db.session.commit()
            flash("Password updated.", "success")
            return redirect(url_for("dashboard"))
    return render_template("update_password.html", title="Update Password")

# ---------- Calendar ----------
def sub_summary_text(subs, legacy_sub_text):
    if subs:
        parts = [f"{s.name}({s.hours:.1f}h)" for s in subs[:2]]
        more = len(subs) - 2
        tail = f" +{more} more" if more > 0 else ""
        return " – Sub: " + ", ".join(parts) + tail
    if (legacy_sub_text or "").strip():
        return " – Sub: " + legacy_sub_text.strip()
    return ""

@app.get("/calendar")
@login_required
def calendar():
    return render_template("calendar.html", title="Calendar", is_admin=(current_user.role == Role.admin), me=current_user)

@app.get("/calendar-data")
@login_required
def calendar_data():
    q = LeaveRequest.query.filter_by(status=RequestStatus.approved)
    if current_user.role != Role.admin:
        q = q.filter_by(user_id=current_user.id)
    events = []
    for r in q.all():
        if current_user.role == Role.admin:
            title = f"{r.user.username} - {r.kind} ({r.hours:.1f}h)"
            title += sub_summary_text(getattr(r, "subs", []), r.substitute)
        else:
            title = f"{r.kind} ({r.hours:.1f}h)"
        if r.is_school_related: title = "[School] " + title
        events.append({
            "title": title,
            "start": r.start_date.isoformat(),
            "end": (r.end_date + timedelta(days=1)).isoformat(),
        })
    return jsonify(events)

# ---------- Exports ----------
@app.get("/admin/export/requests.csv")
@login_required
def export_requests_csv():
    if current_user.role != Role.admin: abort(403)
    rows = _filtered_requests_for(True).all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "ID","Username","Kind","Mode","Hours","Status","Start","End",
        "StartTime","EndTime","SchoolRelated","Substitutes","Created","Decided"
    ])
    for r in rows:
        # build substitutes text (multi + legacy)
        subs_text = ""
        try:
            parts = [f"{s.name}({s.hours:.2f}h)" for s in getattr(r, "subs", [])]
            subs_text = "; ".join(parts)
        except Exception:
            pass
        if not subs_text:
            subs_text = r.substitute or ""

        writer.writerow([
            r.id, r.user.username, r.kind, r.mode, f"{r.hours:.2f}", r.status,
            r.start_date.isoformat(), r.end_date.isoformat(),
            r.start_time or "", r.end_time or "",
            "Yes" if r.is_school_related else "No",
            subs_text,
            r.created_at.isoformat() if r.created_at else "",
            r.decided_at.isoformat() if r.decided_at else ""
        ])
    resp = make_response(output.getvalue())
    resp.headers["Content-Type"] = "text/csv"
    resp.headers["Content-Disposition"] = "attachment; filename=leave_requests.csv"
    return resp

@app.get("/admin/export/requests.xlsx")
@login_required
def export_requests_xlsx():
    if current_user.role != Role.admin:
        abort(403)
    rows = _filtered_requests_for(True).all()

    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})

    # Sheet 1: Requests
    ws = wb.add_worksheet("Requests")
    headers = ["ID","Username","Kind","Mode","Hours","Status","Start","End",
               "StartTime","EndTime","SchoolRelated","Substitutes","Created","Decided"]
    hdr_fmt = wb.add_format({"bold": True, "bg_color": "#F1F5F9", "border": 1})
    cell_fmt = wb.add_format({"border": 1})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    dt_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm", "border": 1})

    for c, h in enumerate(headers):
        ws.write(0, c, h, hdr_fmt)

    rix = 1
    for r in rows:
        try:
            subs_parts = [f"{s.name}({s.hours:.2f}h)" for s in getattr(r, "subs", [])]
            subs_text = "; ".join(subs_parts)
        except Exception:
            subs_text = ""
        if not subs_text:
            subs_text = r.substitute or ""

        ws.write(rix, 0, r.id, cell_fmt)
        ws.write(rix, 1, r.user.username, cell_fmt)
        ws.write(rix, 2, r.kind, cell_fmt)
        ws.write(rix, 3, r.mode, cell_fmt)
        ws.write_number(rix, 4, float(r.hours or 0.0), cell_fmt)
        ws.write(rix, 5, r.status, cell_fmt)

        ws.write_datetime(rix, 6, datetime.combine(r.start_date, datetime.min.time()), date_fmt)
        ws.write_datetime(rix, 7, datetime.combine(r.end_date, datetime.min.time()), date_fmt)

        ws.write(rix, 8, r.start_time or "", cell_fmt)
        ws.write(rix, 9, r.end_time or "", cell_fmt)

        ws.write(rix, 10, "Yes" if r.is_school_related else "No", cell_fmt)
        ws.write(rix, 11, subs_text, cell_fmt)

        if r.created_at:
            ws.write_datetime(rix, 12, r.created_at, dt_fmt)
        else:
            ws.write(rix, 12, "", cell_fmt)
        if r.decided_at:
            ws.write_datetime(rix, 13, r.decided_at, dt_fmt)
        else:
            ws.write(rix, 13, "", cell_fmt)

        rix += 1

    # Sheet 2: Substitutes (one row per sub assignment)
    ws2 = wb.add_worksheet("Substitutes")
    ws2_headers = ["RequestID","Username","Start","End","Substitute","Hours"]
    for c, h in enumerate(ws2_headers): ws2.write(0, c, h, hdr_fmt)
    rix = 1
    for r in rows:
        try:
            subs = getattr(r, "subs", [])
        except Exception:
            subs = []
        for s in subs:
            ws2.write(rix, 0, r.id, cell_fmt)
            ws2.write(rix, 1, r.user.username, cell_fmt)
            ws2.write_datetime(rix, 2, datetime.combine(r.start_date, datetime.min.time()), date_fmt)
            ws2.write_datetime(rix, 3, datetime.combine(r.end_date, datetime.min.time()), date_fmt)
            ws2.write(rix, 4, s.name, cell_fmt)
            ws2.write_number(rix, 5, float(s.hours or 0.0), cell_fmt)
            rix += 1

    wb.close()
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="leave_requests.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- Errors ----------
@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", title="Not Found", message="The page you requested was not found."), 404

@app.errorhandler(500)
def internal_error(e):
    try:
        return render_template("error.html", title="Server Error", message=str(e)), 500
    except Exception:
        return "Internal Server Error", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
