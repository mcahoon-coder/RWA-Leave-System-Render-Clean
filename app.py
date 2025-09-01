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
import os, smtplib, ssl, io, csv, logging, sys
from email.message import EmailMessage
from sqlalchemy import text
import xlsxwriter  # Excel export (in-memory, safe on Render)

# =========================================================
# App & DB config
# =========================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ChangeThisSecret123!")

# Prefer Render DATABASE_URL; default to SQLite
db_url = os.environ.get("DATABASE_URL", "sqlite:///leave_system.db")

# Normalize old Heroku-style scheme and ensure SQLAlchemy uses psycopg (v3)
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
# Context processors (always-available template vars + safe NAV)
# =========================================================
@app.context_processor
def inject_globals_and_nav():
    # Always-available context so templates don't 500 if a page forgets to pass these
    safe_me = current_user if hasattr(current_user, "is_authenticated") and current_user.is_authenticated else None

    def safe(endpoint, fallback):
        try:
            return url_for(endpoint)
        except Exception:
            return fallback

    return {
        "current_year": datetime.utcnow().year,
        "me": safe_me,
        "user": safe_me,
        "workday": float(os.environ.get("WORKDAY_HOURS", "8")),
        "NAV": {
            "dashboard":     safe("dashboard", "/dashboard"),
            "my_requests":   safe("my_requests", "/requests"),
            "team_calendar": safe("calendar", "/calendar"),
            "new_request":   safe("new_request", "/request/new"),
            "admin":         safe("admin_home", "/admin"),
            "logout":        safe("logout", "/logout"),
            "login":         safe("login", "/login"),
            "home":          safe("home", "/"),
        }
    }

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

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
        context = ssl.create_default_context()
        with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as s:
            s.ehlo(); s.starttls(context=context); s.ehlo()
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
    staff = "faculty_staff"  # requested label

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
    user = db.relationship("User", backref="leave_requests", lazy="joined")

# =========================================================
# Helpers
# =========================================================
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
        cur += timedelta(days=1)
    return n

def parse_quarter_time(s: str) -> dt_time | None:
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
    dialect = bind.dialect.name
    if dialect == "sqlite":
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
        if not _column_exists("leave_request", "start_time"):
            if db.engine.dialect.name == "sqlite":
                db.session.execute(text("ALTER TABLE leave_request ADD COLUMN start_time VARCHAR(5)"))
                db.session.commit()
        if not _column_exists("leave_request", "end_time"):
            if db.engine.dialect.name == "sqlite":
                db.session.execute(text("ALTER TABLE leave_request ADD COLUMN end_time VARCHAR(5)"))
                db.session.commit()
    except Exception:
        db.session.rollback()

    if User.query.count() == 0:
        bootstrap_username = os.environ.get("BOOTSTRAP_ADMIN_USERNAME", "mc-admin")
        bootstrap_password = os.environ.get("BOOTSTRAP_ADMIN_PASSWORD", "RWAadmin2")
        bootstrap_email    = os.environ.get("BOOTSTRAP_ADMIN_EMAIL", (ADMIN_EMAILS_ENV[0] if ADMIN_EMAILS_ENV else ""))
        db.session.add(User(
            username=bootstrap_username,
            password_hash=generate_password_hash(bootstrap_password),
            role=Role.admin,
            hours_balance=160.0,
            email=bootstrap_email or None
        ))
        db.session.commit()

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
    status = (request.args.get("status", "all") or "all").strip().lower()
    start_s = (request.args.get("start", "") or "").strip()
    end_s   = (request.args.get("end", "") or "").strip()

    q = LeaveRequest.query
    if not current_user_is_admin:
        q = q.filter_by(user_id=current_user.id)

    if status and status != "all":
        q = q.filter_by(status=status.capitalize())

    def pd(s):
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except Exception:
            return None

    sd = pd(start_s); ed = pd(end_s)
    if sd: q = q.filter(LeaveRequest.start_date >= sd)
    if ed: q = q.filter(LeaveRequest.end_date <= ed)
    return q.order_by(LeaveRequest.created_at.desc())

# =========================================================
# Routes
# =========================================================
@app.get("/health")
def health():
    return "ok", 200

@app.get("/healthz")
def healthz():
    return {"ok": True}, 200

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
    recent = (
        LeaveRequest.query.filter_by(user_id=current_user.id)
        .order_by(LeaveRequest.created_at.desc())
        .limit(10).all()
    )
    return render_template("dashboard.html", title="Dashboard", recent=recent)

# ----- Route aliases for navbar consistency -----
@app.get("/my-requests")
@login_required
def my_requests_alias():
    return redirect(url_for("my_requests"))

@app.get("/new-request")
@login_required
def new_request_alias():
    return redirect(url_for("new_request"))

# ----- New Request -----
@app.route("/request/new", methods=["GET", "POST"])
@login_required
def new_request():
    if request.method == "POST":
        mode = request.form.get("mode", RequestMode.hourly)
        kind = request.form.get("kind", "annual")
        reason = request.form.get("reason", "")
        try:
            sd = datetime.strptime(request.form["start_date"], "%Y-%m-%d").date()
            ed = datetime.strptime(request.form["end_date"], "%Y-%m-%d").date()
        except Exception:
            flash("Invalid dates.", "warning")
            return render_template("new_request.html", title="New Request")

        capacity_hours = workdays_between(sd, ed) * WORKDAY_HOURS
        if capacity_hours <= 0:
            flash("No working days in that range.", "warning")
            return render_template("new_request.html", title="New Request")

        hours = 0.0
        if mode == RequestMode.hourly:
            hours_str = (request.form.get("hours") or "").strip()
            if hours_str:
                try: hours = float(hours_str)
                except Exception: hours = 0.0
            else:
                st = parse_quarter_time((request.form.get("start_time") or "").strip())
                et = parse_quarter_time((request.form.get("end_time") or "").strip())
                if st and et and sd == ed:
                    hours = interval_hours(st, et)
        else:
            hours = workdays_between(sd, ed) * WORKDAY_HOURS

        if hours <= 0:
            flash("Requested hours must be greater than zero.", "warning")
            return render_template("new_request.html", title="New Request")

        if hours > capacity_hours and mode != RequestMode.hourly:
            flash(f"Requested {hours:.2f} exceeds capacity {capacity_hours:.2f} for that range.", "warning")
            return render_template("new_request.html", title="New Request")

        req = LeaveRequest(
            user_id=current_user.id, kind=kind, mode=mode,
            start_date=sd, end_date=ed,
            start_time=request.form.get("start_time") or None,
            end_time=request.form.get("end_time") or None,
            hours=hours, reason=reason
        )
        db.session.add(req); db.session.commit()

        subj = "New Leave Request Submitted"
        body = (
            f"User: {current_user.username}\nKind: {kind}\nMode: {mode}\nHours: {hours:.2f}\n"
            f"Dates: {sd} to {ed}\nTimes: {req.start_time or '-'} to {req.end_time or '-'}\n"
            f"Reason: {reason or '(none)'}\nStatus: {req.status}\n"
        )
        send_email(admin_emails(), subj, body)
        flash("Request submitted.", "success")
        return redirect(url_for("my_requests"))

    return render_template("new_request.html", title="New Request")

# ----- Requests list -----
@app.get("/requests")
@login_required
def my_requests():
    q = _filtered_requests_for(current_user.role == Role.admin)
    items = q.all()
    # Provide both names to satisfy any existing template expectations
    return render_template(
        "requests.html",
        title="Requests",
        reqs=items,
        requests=items,  # alias
        is_admin=(current_user.role == Role.admin),
        status=request.args.get("status", "all"),
        start=request.args.get("start", ""),
        end=request.args.get("end", "")
    )

# ----- Approvals -----
@app.post("/requests/<int:req_id>/approve")
@login_required
def approve(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning")
        return redirect(url_for("my_requests"))
    u = User.query.get(r.user_id)
    u.hours_balance = float(u.hours_balance or 0.0) - float(r.hours or 0.0)
    r.status = RequestStatus.approved
    r.decided_at = datetime.utcnow()
    db.session.commit()
    subj = "Leave Request Approved"
    body = (
        f"Hello {u.username},\n\nYour leave request has been APPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"Dates: {r.start_date} to {r.end_date}\n\nRemaining balance: {u.hours_balance:.2f} hours\n"
    )
    send_email([u.email] + admin_emails(), subj, body)
    flash("Approved.", "success")
    return redirect(url_for("my_requests"))

@app.post("/requests/<int:req_id>/disapprove")
@login_required
def disapprove(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning")
        return redirect(url_for("my_requests"))
    r.status = RequestStatus.disapproved
    r.decided_at = datetime.utcnow()
    db.session.commit()
    u = User.query.get(r.user_id)
    subj = "Leave Request Disapproved"
    body = (
        f"Hello {u.username},\n\nYour leave request has been DISAPPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
    )
    send_email([u.email] + admin_emails(), subj, body)
    flash("Disapproved.", "info")
    return redirect(url_for("my_requests"))

@app.post("/requests/<int:req_id>/cancel")
@login_required
def cancel(req_id):
    r = LeaveRequest.query.get_or_404(req_id)
    if r.user_id != current_user.id and current_user.role != Role.admin:
        flash("Not allowed.", "danger")
        return redirect(url_for("my_requests"))
    u = User.query.get(r.user_id)
    if r.status == RequestStatus.approved:
        u.hours_balance = float(u.hours_balance or 0.0) + float(r.hours or 0.0)
    r.status = RequestStatus.cancelled
    r.decided_at = datetime.utcnow()
    db.session.commit()
    subj = "Leave Request Cancelled"
    body = (
        f"User {u.username} cancelled a leave request.\nKind: {r.kind}\nMode: {r.mode}\n"
        f"Hours: {r.hours:.2f}\nDates: {r.start_date} to {r.end_date}\n"
        f"Balance is now: {u.hours_balance:.2f} hours\n"
    )
    recipients = admin_emails()
    if u.email: recipients = [u.email] + recipients
    send_email(recipients, subj, body)
    flash("Cancelled.", "secondary")
    return redirect(url_for("my_requests"))

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
@app.get("/calendar")
@login_required
def calendar():
    return render_template("calendar.html", title="Calendar")

@app.get("/calendar-data")
@login_required
def calendar_data():
    events = []
    approved = LeaveRequest.query.filter_by(status=RequestStatus.approved).all()
    for r in approved:
        events.append({
            "title": f"{r.user.username} - {r.kind} ({r.hours:.1f}h)",
            "start": r.start_date.isoformat(),
            "end": (r.end_date + timedelta(days=1)).isoformat()  # exclusive end for FullCalendar
        })
    return jsonify(events)

# ---------- Admin Home (User Management, Approvals, Reports) ----------
@app.get("/admin")
@login_required
def admin_home():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))
    return render_template("admin.html", title="Admin")

# ---------- User Management screen ----------
@app.route("/admin/users", methods=["GET"])
@login_required
def manage_users():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))
    qtxt = request.args.get("q", "").strip()
    query = User.query
    if qtxt:
        query = query.filter(User.username.ilike(f"%{qtxt}%"))
    users = query.order_by(User.username.asc()).all()
    return render_template("manage_users.html", title="Manage Users", users=users, q=qtxt)

@app.post("/admin/users/create")
@login_required
def admin_create_user():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))
    username = (request.form.get("username") or "").strip()
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
    try:
        hours_balance = float(hours_str) if hours_str else 160.0
    except Exception:
        hours_balance = 160.0
    user = User(
        username=username,
        password_hash=generate_password_hash(pw),
        role=role if role in (Role.admin, Role.staff) else Role.staff,
        hours_balance=hours_balance,
        email=email or None
    )
    db.session.add(user); db.session.commit()
    flash(f"User '{username}' created.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/update")
@login_required
def admin_update_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    new_username = (request.form.get("username") or "").strip()
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or "").strip()
    hb = (request.form.get("hours_balance") or "").strip()
    if new_username and new_username.lower() != u.username.lower():
        if User.query.filter(User.username.ilike(new_username)).first():
            flash("Username already taken.", "danger")
            return redirect(url_for("manage_users"))
        u.username = new_username
    u.email = email or None
    if role in (Role.admin, Role.staff):
        u.role = role
    try:
        if hb != "":
            u.hours_balance = float(hb)
    except Exception:
        flash("Invalid hours_balance value.", "warning")
    db.session.commit()
    flash(f"Updated {u.username}.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/reset")
@login_required
def admin_reset_password(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))
    new_pw = (request.form.get("new_password") or "").strip()
    if not new_pw:
        flash("Password cannot be empty.", "warning")
        return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    u.password_hash = generate_password_hash(new_pw)
    db.session.commit()
    flash(f"Password updated for {u.username}.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/delete")
@login_required
def admin_delete_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)
    if u.id == current_user.id:
        flash("You cannot delete your own account.", "warning")
        return redirect(url_for("manage_users"))
    if u.role == Role.admin and User.query.filter_by(role=Role.admin).count() <= 1:
        flash("At least one admin must remain.", "warning")
        return redirect(url_for("manage_users"))
    LeaveRequest.query.filter_by(user_id=u.id).delete()
    db.session.delete(u); db.session.commit()
    flash("User deleted.", "success")
    return redirect(url_for("manage_users"))

# ---------- Exports (admin only) ----------
@app.get("/admin/export/requests.csv")
@login_required
def export_requests_csv():
    if current_user.role != Role.admin:
        abort(403)
    rows = _filtered_requests_for(True).all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["ID","Username","Kind","Mode","Hours","Status","Start","End","StartTime","EndTime","Created","Decided"])
    for r in rows:
        writer.writerow([
            r.id, r.user.username, r.kind, r.mode, f"{r.hours:.2f}", r.status,
            r.start_date.isoformat(), r.end_date.isoformat(),
            r.start_time or "", r.end_time or "",
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
    ws = wb.add_worksheet("Requests")

    headers = ["ID","Username","Kind","Mode","Hours","Status","Start","End","StartTime","EndTime","Created","Decided"]
    hdr_fmt = wb.add_format({"bold": True, "bg_color": "#F1F5F9", "border": 1})
    cell_fmt = wb.add_format({"border": 1})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    dt_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm", "border": 1})

    for c, h in enumerate(headers):
        ws.write(0, c, h, hdr_fmt)

    rix = 1
    for r in rows:
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
        if r.created_at:
            ws.write_datetime(rix, 10, r.created_at, dt_fmt)
        else:
            ws.write(rix, 10, "", cell_fmt)
        if r.decided_at:
            ws.write_datetime(rix, 11, r.decided_at, dt_fmt)
        else:
            ws.write(rix, 11, "", cell_fmt)
        rix += 1

    widths = [len(h) for h in headers]
    for r in rows:
        widths[1] = max(widths[1], len(r.user.username or ""))
        widths[2] = max(widths[2], len(r.kind or ""))
        widths[3] = max(widths[3], len(r.mode or ""))
        widths[5] = max(widths[5], len(r.status or ""))
    for c, w in enumerate(widths):
        ws.set_column(c, c, min(max(w + 2, 10), 32))

    wb.close()
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="leave_requests.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- Errors + logging ----------
@app.errorhandler(404)
def not_found(e):
    return render_template("error.html", title="Not Found", message="The page you requested was not found."), 404

@app.errorhandler(500)
def internal_error(e):
    app.logger.exception("Unhandled 500: %s", e)
    try:
        return render_template("error.html", title="Server Error", message=str(e)), 500
    except Exception:
        return "Internal Server Error", 500

# Log to stderr (helps Render logs)
handler = logging.StreamHandler(sys.stderr)
handler.setLevel(logging.INFO)
app.logger.addHandler(handler)
app.logger.setLevel(logging.INFO)
app.config["PROPAGATE_EXCEPTIONS"] = True

# Dev server entry (ignored by gunicorn)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
