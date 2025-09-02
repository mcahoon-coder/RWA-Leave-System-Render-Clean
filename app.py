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
import xlsxwriter  # Excel export (in-memory)

# =========================================================
# App & DB config
# =========================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ChangeThisSecret123!")

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
    if not to_addrs:
        return
    if isinstance(to_addrs, str):
        to_addrs = [to_addrs]
    to_addrs = [a for a in to_addrs if a]
    if not to_addrs or not MAIL_HOST or not MAIL_FROM:
        return
    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join(to_addrs)
    msg["Subject"] = subject
    msg.set_content(body)
    if MAIL_USE_TLS:
        ctx = ssl.create_default_context()
        with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as s:
            s.ehlo(); s.starttls(context=ctx); s.ehlo()
            if MAIL_USER: s.login(MAIL_USER, MAIL_PASSWORD)
            s.send_message(msg)
    else:
        with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as s:
            if MAIL_USER: s.login(MAIL_USER, MAIL_PASSWORD)
            s.send_message(msg)

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

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default=Role.staff, nullable=False)
    hours_balance = db.Column(db.Float, default=160.0, nullable=False)
    email = db.Column(db.String(255))

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
    substitute = db.Column(db.String(120))  # legacy single text
    user = db.relationship("User", backref="leave_requests", lazy="joined")
    subs = db.relationship("SubAssignment", backref="request", cascade="all, delete-orphan", lazy="joined")

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

WORKDAY_HOURS = float(os.environ.get("WORKDAY_HOURS", "8"))
HOLIDAYS: set[date] = set()

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in HOLIDAYS

def workdays_between(start: date, end: date) -> int:
    if end < start:
        start, end = end, start
    n, cur = 0, start
    while cur <= end:
        if is_workday(cur):
            n += 1
        cur = cur + timedelta(days=1)
    return n

def parse_quarter_time(s: str):
    try:
        hh, mm = s.split(":")
        hh_i = int(hh); mm_i = int(mm)
        if 0 <= hh_i <= 23 and mm_i in (0, 15, 30, 45):
            return dt_time(hh_i, mm_i)
    except Exception:
        return None
    return None

def interval_hours(t1: dt_time, t2: dt_time) -> float:
    dt1 = datetime.combine(date.today(), t1)
    dt2 = datetime.combine(date.today(), t2)
    if dt2 < dt1:
        dt1, dt2 = dt2, dt1
    return (dt2 - dt1).total_seconds() / 3600.0

def _column_exists(table_name: str, column_name: str) -> bool:
    bind = db.engine
    if bind.dialect.name == "sqlite":
        res = db.session.execute(text(f"PRAGMA table_info({table_name})")).fetchall()
        return any(row[1] == column_name for row in res)
    q = text("""
        SELECT 1 FROM information_schema.columns
        WHERE table_name = :t AND column_name = :c
        LIMIT 1
    """)
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
            db.session.commit()
        else:
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS start_time VARCHAR(5)"))
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS end_time VARCHAR(5)"))
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS is_school_related BOOLEAN NOT NULL DEFAULT FALSE"))
            db.session.execute(text("ALTER TABLE leave_request ADD COLUMN IF NOT EXISTS substitute VARCHAR(120)"))
            db.session.commit()
    except Exception:
        db.session.rollback()

    if User.query.count() == 0:
        u = User(
            username=os.environ.get("BOOTSTRAP_ADMIN_USERNAME", "mc-admin"),
            password_hash=generate_password_hash(os.environ.get("BOOTSTRAP_ADMIN_PASSWORD", "RWAadmin2")),
            role=Role.admin,
            hours_balance=160.0,
            email=os.environ.get("BOOTSTRAP_ADMIN_EMAIL", (ADMIN_EMAILS_ENV[0] if ADMIN_EMAILS_ENV else "")) or None
        )
        db.session.add(u); db.session.commit()

with app.app_context():
    ensure_db()

def admin_emails() -> list[str]:
    env_list = ADMIN_EMAILS_ENV[:]
    user_list = [u.email for u in User.query.filter_by(role=Role.admin).all() if u.email]
    seen, result = set(), []
    for e in env_list + user_list:
        if e and e not in seen:
            result.append(e); seen.add(e)
    return result

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
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except Exception:
            return None

    sd = parse_date(start_s); ed = parse_date(end_s)
    if sd: q = q.filter(LeaveRequest.start_date >= sd)
    if ed: q = q.filter(LeaveRequest.end_date <= ed)
    return q.order_by(LeaveRequest.created_at.desc())

# =========================================================
# Nav + globals to templates
# =========================================================
@app.context_processor
def inject_globals():
    class NAV:
        pass
    nav = NAV()
    if current_user.is_authenticated:
        nav.dashboard = url_for("dashboard")
        nav.my_requests = url_for("my_requests")
        nav.team_calendar = url_for("calendar")
        nav.new_request = url_for("new_request")
        nav.admin = url_for("admin_hub") if current_user.role == Role.admin else None
        nav.logout = url_for("logout")
    else:
        login_url = url_for("login")
        nav.dashboard = nav.my_requests = nav.team_calendar = nav.new_request = nav.admin = nav.logout = login_url
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
            login_user(user); flash("Logged in.", "success")
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
    return render_template("dashboard.html", title="Dashboard", me=current_user,
                           workday=WORKDAY_HOURS, recent=recent)

# ---------- Admin HUB ----------
@app.get("/admin")
@login_required
def admin_hub():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("dashboard"))
    pending = (LeaveRequest.query.filter_by(status=RequestStatus.pending)
               .order_by(LeaveRequest.created_at.desc()).all())
    return render_template("admin.html", title="Admin", pending=pending)

# ---------- New Request ----------
@app.route("/request/new", methods=["GET", "POST"])
@login_required
def new_request():
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
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

        capacity_hours = workdays_between(sd, ed) * WORKDAY_HOURS
        if capacity_hours <= 0:
            flash("No working days in that range.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

        hours = 0.0
        if mode == RequestMode.hourly:
            hours_str = (request.form.get("hours") or "").strip()
            if hours_str:
                try: hours = float(hours_str)
                except Exception: hours = 0.0
            else:
                st = parse_quarter_time((request.form.get("start_time") or "").strip())
                et = parse_quarter_time((request.form.get("end_time") or "").strip())
                if st and et and sd == ed: hours = interval_hours(st, et)
        else:
            hours = workdays_between(sd, ed) * WORKDAY_HOURS

        if hours <= 0:
            flash("Requested hours must be greater than zero.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)
        if hours > capacity_hours and mode != RequestMode.hourly:
            flash(f"Requested {hours:.2f} exceeds capacity {capacity_hours:.2f}.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

        req = LeaveRequest(
            user_id=current_user.id, kind=kind, mode=mode,
            start_date=sd, end_date=ed,
            start_time=request.form.get("start_time") or None,
            end_time=request.form.get("end_time") or None,
            hours=hours, reason=reason, is_school_related=is_school
        )
        db.session.add(req); db.session.commit()

        subj = "New Leave Request Submitted"
        body = (
            f"User: {current_user.username}\nKind: {kind}\nMode: {mode}\nHours: {hours:.2f}\n"
            f"Dates: {sd} to {ed}\nTimes: {req.start_time or '-'} to {req.end_time or '-'}\n"
            f"School-related: {'Yes' if is_school else 'No'}\nReason: {reason or '(none)'}\n"
            f"Status: {req.status}\n"
        )
        send_email(admin_emails(), subj, body)
        flash("Request submitted.", "success")
        return redirect(url_for("my_requests"))

    return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

# ---------- Requests list (admin sees all, staff sees own) ----------
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
    return redirect(url_for("my_requests"))

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
    return redirect(url_for("my_requests"))

# ---------- Substitutes (admin only; multiple with hours) ----------
@app.post("/requests/<int:req_id>/subs/add")
@login_required
def add_substitute(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    name = (request.form.get("sub_name") or "").strip()
    hrs_s = (request.form.get("sub_hours") or "").strip()
    if not name:
        flash("Substitute name required.", "warning"); return redirect(url_for("my_requests"))
    try:
        hours = float(hrs_s) if hrs_s else 0.0
    except Exception:
        flash("Invalid hours.", "warning"); return redirect(url_for("my_requests"))
    db.session.add(SubAssignment(request_id=r.id, name=name, hours=hours))
    db.session.commit()
    flash("Substitute added.", "success")
    return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/subs/<int:sub_id>/update")
@login_required
def update_substitute(req_id, sub_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    sub = SubAssignment.query.get_or_404(sub_id)
    if sub.request_id != req_id: abort(404)
    name = (request.form.get("sub_name") or "").strip()
    hrs_s = (request.form.get("sub_hours") or "").strip()
    if name: sub.name = name
    try:
        if hrs_s != "": sub.hours = float(hrs_s)
    except Exception:
        flash("Invalid hours.", "warning"); return redirect(url_for("my_requests"))
    db.session.commit()
    flash("Substitute updated.", "success")
    return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/subs/<int:sub_id>/delete")
@login_required
def delete_substitute(req_id, sub_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    sub = SubAssignment.query.get_or_404(sub_id)
    if sub.request_id != req_id: abort(404)
    db.session.delete(sub); db.session.commit()
    flash("Substitute removed.", "success")
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
    subs_text = "; ".join([f"{s.name} ({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "(none)")
    subj = "Leave Request Approved"
    body = (
        f"Hello {u.username},\n\nYour leave request has been APPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"School-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\nDates: {r.start_date} to {r.end_date}\n\n"
        f"Remaining balance: {u.hours_balance:.2f} hours\n"
    )
    send_email([u.email] + admin_emails(), subj, body)
    flash("Approved.", "success"); return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/disapprove")
@login_required
def disapprove(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning"); return redirect(url_for("my_requests"))
    r.status = RequestStatus.disapproved; r.decided_at = datetime.utcnow(); db.session.commit()
    u = User.query.get(r.user_id)
    subs_text = "; ".join([f"{s.name} ({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "(none)")
    subj = "Leave Request Disapproved"
    body = (
        f"Hello {u.username},\n\nYour leave request has been DISAPPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"School-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\nDates: {r.start_date} to {r.end_date}\n"
    )
    send_email([u.email] + admin_emails(), subj, body)
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
    r.status = RequestStatus.cancelled; r.decided_at = datetime.utcnow(); db.session.commit()
    subs_text = "; ".join([f"{s.name} ({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "(none)")
    subj = "Leave Request Cancelled"
    body = (
        f"User {u.username} cancelled a leave request.\nKind: {r.kind}\nMode: {r.mode}\n"
        f"Hours: {r.hours:.2f}\nSchool-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\nDates: {r.start_date} to {r.end_date}\n"
        f"Balance is now: {u.hours_balance:.2f} hours\n"
    )
    recipients = admin_emails()
    if u.email: recipients = [u.email] + recipients
    send_email(recipients, subj, body)
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
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or Role.staff).strip()
    hours_str = (request.form.get("hours_balance") or "").strip()
    pw = (request.form.get("password") or "").strip()
    if not username or not pw:
        flash("Username and password are required.", "warning"); return redirect(url_for("manage_users"))
    if User.query.filter(User.username.ilike(username)).first():
        flash("Username already exists.", "danger"); return redirect(url_for("manage_users"))
    try: hours_balance = float(hours_str) if hours_str else 160.0
    except Exception: hours_balance = 160.0
    user = User(username=username, password_hash=generate_password_hash(pw),
                role=role if role in (Role.admin, Role.staff) else Role.staff,
                hours_balance=hours_balance, email=email or None)
    db.session.add(user); db.session.commit()
    flash(f"User '{username}' created.", "success"); return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/update")
@login_required
def admin_update_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    new_username = (request.form.get("username") or "").strip()
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or "").strip()
    hb = (request.form.get("hours_balance") or "").strip()
    if new_username and new_username.lower() != u.username.lower():
        if User.query.filter(User.username.ilike(new_username)).first():
            flash("Username already taken.", "danger"); return redirect(url_for("manage_users"))
        u.username = new_username
    u.email = email or None
    if role in (Role.admin, Role.staff): u.role = role
    try:
        if hb != "": u.hours_balance = float(hb)
    except Exception:
        flash("Invalid hours_balance value.", "warning")
    db.session.commit(); flash(f"Updated {u.username}.", "success"); return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/reset")
@login_required
def admin_reset_password(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("manage_users"))
    new_pw = (request.form.get("new_password") or "").strip()
    if not new_pw:
        flash("Password cannot be empty.", "warning"); return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    u.password_hash = generate_password_hash(new_pw); db.session.commit()
    flash(f"Password updated for {u.username}.", "success"); return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/delete")
@login_required
def admin_delete_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    if u.id == current_user.id:
        flash("You cannot delete your own account.", "warning"); return redirect(url_for("manage_users"))
    if u.role == Role.admin and User.query.filter_by(role=Role.admin).count() <= 1:
        flash("At least one admin must remain.", "warning"); return redirect(url_for("manage_users"))
    LeaveRequest.query.filter_by(user_id=u.id).delete()
    db.session.delete(u); db.session.commit()
    flash("User deleted.", "success"); return redirect(url_for("manage_users"))

# ---------- Password update ----------
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
def sub_summary_text(subs, limit=2):
    if not subs: return ""
    parts = [f"{s.name}({s.hours:.1f}h)" for s in subs[:limit]]
    more = len(subs) - limit
    return " – Sub: " + ", ".join(parts) + (f" +{more} more" if more > 0 else "")

@app.get("/calendar")
@login_required
def calendar():
    return render_template("calendar.html", title="Calendar",
                           is_admin=(current_user.role == Role.admin), me=current_user)

@app.get("/calendar-data")
@login_required
def calendar_data():
    q = LeaveRequest.query.filter_by(status=RequestStatus.approved)
    is_admin = (current_user.role == Role.admin)
    if not is_admin:
        q = q.filter_by(user_id=current_user.id)
    events = []
    for r in q.all():
        if is_admin:
            title = f"{r.user.username} - {r.kind} ({r.hours:.1f}h)"
            sub_text = sub_summary_text(r.subs, limit=2)
            if not sub_text and (r.substitute or ""): sub_text = " – Sub: " + r.substitute.strip()
            title += sub_text
        else:
            title = f"{r.kind} ({r.hours:.1f}h)"
        if r.is_school_related:
            title = "[School] " + title
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
    out = io.StringIO(); w = csv.writer(out)
    w.writerow(["ID","Username","Kind","Mode","Hours","Status","Start","End",
                "StartTime","EndTime","SchoolRelated","Substitutes","Created","Decided"])
    for r in rows:
        subs_text = "; ".join([f"{s.name}({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "")
        w.writerow([r.id, r.user.username, r.kind, r.mode, f"{r.hours:.2f}", r.status,
                    r.start_date.isoformat(), r.end_date.isoformat(),
                    r.start_time or "", r.end_time or "",
                    "Yes" if r.is_school_related else "No",
                    subs_text,
                    r.created_at.isoformat() if r.created_at else "",
                    r.decided_at.isoformat() if r.decided_at else ""])
    resp = make_response(out.getvalue())
    resp.headers["Content-Type"] = "text/csv"
    resp.headers["Content-Disposition"] = "attachment; filename=leave_requests.csv"
    return resp

@app.get("/admin/export/requests.xlsx")
@login_required
def export_requests_xlsx():
    if current_user.role != Role.admin: abort(403)
    rows = _filtered_requests_for(True).all()
    buf = io.BytesIO(); wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = wb.add_worksheet("Requests")
    headers = ["ID","Username","Kind","Mode","Hours","Status","Start","End",
               "StartTime","EndTime","SchoolRelated","Substitutes","Created","Decided"]
    hdr = wb.add_format({"bold": True, "bg_color": "#F1F5F9", "border": 1})
    cell = wb.add_format({"border": 1})
    d = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    dt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm", "border": 1})
    for c,h in enumerate(headers): ws.write(0,c,h,hdr)
    rix = 1
    for r in rows:
        subs_text = "; ".join([f"{s.name}({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "")
        ws.write(rix,0,r.id,cell); ws.write(rix,1,r.user.username,cell)
        ws.write(rix,2,r.kind,cell); ws.write(rix,3,r.mode,cell)
        ws.write_number(rix,4,float(r.hours or 0.0),cell)
        ws.write(rix,5,r.status,cell)
        ws.write_datetime(rix,6,datetime.combine(r.start_date, datetime.min.time()),d)
        ws.write_datetime(rix,7,datetime.combine(r.end_date, datetime.min.time()),d)
        ws.write(rix,8,r.start_time or "",cell); ws.write(rix,9,r.end_time or "",cell)
        ws.write(rix,10,"Yes" if r.is_school_related else "No",cell)
        ws.write(rix,11,subs_text,cell)
        ws.write_datetime(rix,12,r.created_at,dt) if r.created_at else ws.write(rix,12,"",cell)
        ws.write_datetime(rix,13,r.decided_at,dt) if r.decided_at else ws.write(rix,13,"",cell)
        rix += 1
    ws2 = wb.add_worksheet("Substitutes")
    ws2h = ["RequestID","Username","Start","End","Substitute","Hours"]
    for c,h in enumerate(ws2h): ws2.write(0,c,h,hdr)
    rix = 1
    for r in rows:
        for s in r.subs:
            ws2.write(rix,0,r.id,cell); ws2.write(rix,1,r.user.username,cell)
            ws2.write_datetime(rix,2,datetime.combine(r.start_date, datetime.min.time()),d)
            ws2.write_datetime(rix,3,datetime.combine(r.end_date, datetime.min.time()),d)
            ws2.write(rix,4,s.name,cell); ws2.write_number(rix,5,float(s.hours or 0.0),cell)
            rix += 1
    wb.close(); buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="leave_requests.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------- Errors ----------
@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", title="Not Found",
                           message="The page you requested was not found."), 404

@app.errorhandler(500)
def internal_error(e):
    try:
        return render_template("error.html", title="Server Error", message=str(e)), 500
    except Exception:
        return "Internal Server Error", 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
