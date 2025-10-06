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
from zoneinfo import ZoneInfo
import os, smtplib, ssl, io, csv
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from sqlalchemy import text
import xlsxwriter  # Excel export (in-memory, safe on Render)

# =========================================================
# App & DB config
# =========================================================
app = Flask(__name__, static_folder="static", static_url_path="/static")
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ChangeThisSecret123!")
app.config["TEMPLATES_AUTO_RELOAD"] = True

# Prefer Render DATABASE_URL; default to SQLite
db_url = os.environ.get("DATABASE_URL", "sqlite:///leave_system.db")

# Normalize old Heroku-style scheme and ensure SQLAlchemy uses psycopg (v3)
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql+psycopg://", 1)
elif db_url.startswith("postgresql://") and "+psycopg" not in db_url:
    db_url = db_url.replace("postgresql://", "postgresql+psycopg://", 1)

app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {
    "pool_pre_ping": True,
    "pool_recycle": 300     # recycle connections every 5 minutes
}

db = SQLAlchemy(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"

# Avoid stale template caching by proxies/browsers
@app.after_request
def add_no_cache_headers(resp):
    if resp.mimetype == "text/html":
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    return resp

# =========================================================
# Email settings (env vars) — ✅ FIXED to match Render
# =========================================================
MAIL_SERVER = os.environ.get("MAIL_HOST", "smtp.gmail.com")
MAIL_PORT = int(os.environ.get("MAIL_PORT", 587))
MAIL_USE_TLS = os.environ.get("MAIL_USE_TLS", "TRUE").lower() in ("true", "1", "yes")
MAIL_USE_SSL = False
MAIL_USERNAME = os.environ.get("MAIL_USER")        # Gmail login
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD")    # Gmail app password
MAIL_DEFAULT_SENDER = os.environ.get("MAIL_USER")  # from address

# comma-separated list of admin emails for notifications
ADMIN_EMAILS_ENV = [
    e.strip() for e in os.environ.get("ADMIN_EMAILS", "").split(",") if e.strip()
]

def send_email(to_addrs, subject, body):
    """Send an email via SMTP using app config."""
    try:
        if not MAIL_SERVER or not MAIL_USERNAME:
            app.logger.warning("Email skipped: MAIL_SERVER or MAIL_USERNAME not set.")
            return False, "SMTP not configured"

        msg = MIMEText(body)
        msg["Subject"] = subject
        msg["From"] = MAIL_DEFAULT_SENDER
        msg["To"] = ", ".join(to_addrs)

        with smtplib.SMTP(MAIL_SERVER, MAIL_PORT) as server:
            if MAIL_USE_TLS:
                server.starttls()
            server.login(MAIL_USERNAME, MAIL_PASSWORD)
            server.sendmail(MAIL_DEFAULT_SENDER, to_addrs, msg.as_string())

        app.logger.info(f"Email sent to {to_addrs}")
        return True, "sent"

    except Exception as e:
        app.logger.error(f"Email send failed: {e}")
        return False, str(e)
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
    halfday = "halfday"  # 4.00 hr option

class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default=Role.staff, nullable=False)
    hours_balance = db.Column(db.Float, default=160.0, nullable=False)
    email = db.Column(db.String(255))  # for notifications
    # Optional display name for staff
    staff_name = db.Column(db.String(150))

    @property
    def is_admin(self):
        return (self.role or "").lower() == "admin"

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    kind = db.Column(db.String(20), default="annual", nullable=False)    # annual/sick
    mode = db.Column(db.String(10), default=RequestMode.hourly, nullable=False)  # hourly/daily/halfday
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)

    # Optional quarter-hour times (when mode == hourly and hours not provided)
    start_time = db.Column(db.String(5))  # "HH:MM"
    end_time   = db.Column(db.String(5))  # "HH:MM"

    hours = db.Column(db.Float, nullable=False)
    reason = db.Column(db.String(500), default="")
    status = db.Column(db.String(20), default=RequestStatus.pending, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    decided_at = db.Column(db.DateTime)

    # Flags/extra
    is_school_related = db.Column(db.Boolean, default=False, nullable=False)
    substitute = db.Column(db.String(120))  # legacy single substitute text (optional)

    # eager-load user to avoid DetachedInstanceError in templates
    user = db.relationship("User", backref="leave_requests", lazy="joined")

    # Multiple substitutes
    subs = db.relationship(
        "SubAssignment",
        backref="request",
        cascade="all, delete-orphan",
        lazy="joined"
    )

class SubAssignment(db.Model):
    __tablename__ = "sub_assignment"
    id = db.Column(db.Integer, primary_key=True)
    request_id = db.Column(db.Integer, db.ForeignKey("leave_request.id"), nullable=False)
    name = db.Column(db.String(120), nullable=False)
    hours = db.Column(db.Float, nullable=False, default=0.0)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

class ManualAdjustment(db.Model):
    __tablename__ = "manual_adjustments"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    admin_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    hours = db.Column(db.Float, nullable=False)
    note = db.Column(db.String(255), nullable=False)
    timestamp = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship("User", foreign_keys=[user_id], backref="adjustments_received")
    admin = db.relationship("User", foreign_keys=[admin_id])


@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# =========================================================
# Helpers
# =========================================================
WORKDAY_HOURS = float(os.environ.get("WORKDAY_HOURS", "8"))
HOLIDAYS: set[date] = set()  # add date(...) objects here if you want static holidays

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in HOLIDAYS  # Mon–Fri & not holiday

def workdays_between(start: date, end: date) -> int:
    """Inclusive range, counts only Mon–Fri not in HOLIDAYS."""
    if end < start:
        start, end = end, start
    n, cur = 0, start
    while cur <= end:
        if is_workday(cur):
            n += 1
        cur = cur + timedelta(days=1)
    return n

def parse_quarter_time(s: str) -> dt_time | None:
    """Parse 'HH:MM' 24h where MM in {00,15,30,45}."""
    try:
        hh, mm = s.split(":")
        hh_i = int(hh); mm_i = int(mm)
        if 0 <= hh_i <= 23 and mm_i in (0, 15, 30, 45):
            return dt_time(hh_i, mm_i)
    except Exception:
        return None
    return None

def interval_hours(t1: dt_time, t2: dt_time) -> float:
    """Compute hours between two times on same day; if t2 < t1, swap."""
    dt1 = datetime.combine(date.today(), t1)
    dt2 = datetime.combine(date.today(), t2)
    if dt2 < dt1:
        dt1, dt2 = dt2, dt1
    delta = dt2 - dt1
    return delta.total_seconds() / 3600.0

def round_quarter(h: float) -> float:
    """Round to the nearest 0.25 hour."""
    return round(h * 4) / 4.0

def _column_exists(table_name: str, column_name: str) -> bool:
    """Check column existence (SQLite + Postgres)."""
    bind = db.engine
    dialect = bind.dialect.name
    if dialect == "sqlite":
        res = db.session.execute(text(f"PRAGMA table_info({table_name})")).fetchall()
        return any(row[1] == column_name for row in res)
    else:
        q = text("""
            SELECT 1 FROM information_schema.columns
            WHERE table_name = :t AND column_name = :c
            LIMIT 1
        """)
        return db.session.execute(q, {"t": table_name, "c": column_name}).first() is not None

def ensure_db():
    # Create tables (including sub_assignment)
    db.create_all()

    # Add newly introduced columns if missing
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
                db.session.execute(text("ALTER TABLE user ADD COLUMN staff_name VARCHAR(150)"))
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

    # Seed admin ONLY if there are no users at all (first boot)
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
    """All admin notification recipients from env + admin users' emails."""
    env_list = ADMIN_EMAILS_ENV[:]
    user_list = [u.email for u in User.query.filter_by(role=Role.admin).all() if u.email]
    combined = env_list + user_list
    # de-dupe while preserving order
    seen = set()
    result = []
    for e in combined:
        if e and e not in seen:
            result.append(e)
            seen.add(e)
    return result

# Jinja filter: 24h "HH:MM" -> "H:MM AM/PM"
@app.template_filter("h12")
def h12_filter(s):
    try:
        hh, mm = (s or "").split(":")
        hh = int(hh); mm = int(mm)
        ampm = "AM" if hh < 12 else "PM"
        h = hh % 12
        if h == 0: h = 12
        return f"{h}:{mm:02d} {ampm}"
    except Exception:
        return s or ""

# Shared filter logic for list + exports
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

    sd = parse_date(start_s)
    ed = parse_date(end_s)

    if sd:
        q = q.filter(LeaveRequest.start_date >= sd)
    if ed:
        q = q.filter(LeaveRequest.end_date <= ed)

    return q.order_by(LeaveRequest.created_at.desc())

# =========================================================
# Nav + globals in templates
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
        .limit(10)
        .all()
    )
    return render_template(
        "dashboard.html",
        title="Dashboard",
        me=current_user,
        workday=WORKDAY_HOURS,
        recent=recent,
    )

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

# Admin email test endpoint (needed by template button)
@app.get("/admin/email-test")
@login_required
def admin_email_test():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    recipients = []
    if current_user.email:
        recipients.append(current_user.email)
    recipients += admin_emails()
    # de-duplicate
    seen = set(); recipients = [r for r in recipients if r and not (r in seen or seen.add(r))]

    subject = "RWA Leave System – Test Email"
    body = (
        "This is a test email from the RWA Leave System.\n\n"
        f"Time: {datetime.utcnow().isoformat()}Z\n"
        f"From: {MAIL_DEFAULT_SENDER}\nHost: {MAIL_SERVER}:{MAIL_PORT} TLS={MAIL_USE_TLS}\n"
        f"Recipients: {', '.join(recipients) if recipients else '(none)'}\n"
    )

    ok, msg = send_email(recipients, subject, body)
    if ok:
        flash(f"Test email sent to: {', '.join(recipients)}", "success")
    else:
        flash(f"Test email failed: {msg}", "danger")
    return redirect(url_for("admin_hub"))

# ---------- New Request ----------
@app.route("/request/new", methods=["GET", "POST"])
@login_required
def new_request():
    if request.method == "POST":
        mode = request.form.get("mode", RequestMode.hourly)
        kind = request.form.get("kind", "annual")
        reason = request.form.get("reason", "")
        is_school = bool(request.form.get("school_related"))

        # dates
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

        # ---------- compute hours ----------
        hours = 0.0
        start_time_str = (request.form.get("start_time") or "").strip() or None
        end_time_str = (request.form.get("end_time") or "").strip() or None

        if mode == RequestMode.hourly:
            # Require times and same-day range
            if not start_time_str or not end_time_str:
                flash("Please select Start and End times for an hourly request.", "warning")
                return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

            if sd != ed:
                flash("Hourly requests must start and end on the same day.", "warning")
                return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

            st = parse_quarter_time(start_time_str)  # accepts :00/:15/:30/:45
            et = parse_quarter_time(end_time_str)
            if not st or not et:
                flash("Times must be in 15-minute increments.", "warning")
                return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

            computed = interval_hours(st, et)
            if computed <= 0:
                flash("End time must be after start time.", "warning")
                return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

            # Round to nearest quarter hour and store
            hours = round(computed * 4) / 4.0

        elif mode == RequestMode.halfday:
            hours = 4.0

        else:  # daily
            wd = workdays_between(sd, ed)
            hours = wd * WORKDAY_HOURS
            if hours > capacity_hours:
                flash(
                    f"Requested {hours:.2f} exceeds capacity {capacity_hours:.2f} for that range.",
                    "warning",
                )
                return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

        # Guard: require a positive hours value after all logic above
        if hours <= 0:
            flash("Requested hours must be greater than zero.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

        # Normalize stored time strings to "HH:MM" (or None when not hourly)
        def _norm(t: str | None) -> str | None:
            if not t:
                return None
            t = t.strip()
            return t[:5] if len(t) >= 5 else t

        req = LeaveRequest(
            user_id=current_user.id,
            kind=kind,
            mode=mode,
            start_date=sd,
            end_date=ed,
            start_time=_norm(start_time_str) if mode == RequestMode.hourly else None,
            end_time=_norm(end_time_str) if mode == RequestMode.hourly else None,
            hours=hours,
            reason=reason,
            is_school_related=is_school,
        )
        db.session.add(req)
        db.session.commit()

        # Notify admins
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

    return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

# ---------- Requests list (admin sees all, staff sees own) ----------
@app.get("/requests")
@login_required
def my_requests():
    is_admin = (current_user.role == Role.admin)

    q = _filtered_requests_for(is_admin)

    # Admin: show a "daily" slice in America/New_York timezone
    selected_day_str = None
    if is_admin:
        tz = ZoneInfo("America/New_York")
        # If ?day=all → no filter; otherwise default to "today" in local tz
        selected_day_str = request.args.get("day") or datetime.now(tz).date().isoformat()

        if selected_day_str != "all":
            try:
                d = datetime.strptime(selected_day_str, "%Y-%m-%d").date()
            except Exception:
                d = datetime.now(tz).date()
                selected_day_str = d.isoformat()

            # Use local midnight boundaries
            start_dt = datetime.combine(d, datetime.min.time(), tzinfo=tz)
            end_dt = start_dt + timedelta(days=1)

            # Convert to UTC before filtering DB (assuming DB stores naive/UTC timestamps)
            start_utc = start_dt.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
            end_utc = end_dt.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)

            q = q.filter(LeaveRequest.created_at >= start_utc,
                         LeaveRequest.created_at < end_utc)

    reqs = q.all()

    staff_overview = None
    if is_admin:
        users = User.query.order_by(User.username.asc()).all()
        pending_user_ids = {
            r.user_id for r in LeaveRequest.query.with_entities(LeaveRequest.user_id)\
            .filter(LeaveRequest.status == RequestStatus.pending).all()
        }
        staff_overview = [
            {
                "id": u.id,
                "username": u.username,
                "hours_balance": float(u.hours_balance or 0.0),
                "has_pending": (u.id in pending_user_ids),
            }
            for u in users
        ]

    return render_template(
        "requests.html",
        title="Requests",
        reqs=reqs,
        me=current_user,
        is_admin=is_admin,
        status=request.args.get("status", "all"),
        start=request.args.get("start", ""),
        end=request.args.get("end", ""),
        staff_overview=staff_overview,
        selected_day=selected_day_str,
    )

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
        flash("Substitute name required.", "warning"); return redirect(request.referrer or url_for("my_requests"))
    try:
        hours = float(hrs_s) if hrs_s else 0.0
    except Exception:
        flash("Invalid hours.", "warning"); return redirect(request.referrer or url_for("my_requests"))
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
    if sub.request_id != req_id:
        abort(404)
    name = (request.form.get("sub_name") or "").strip()
    hrs_s = (request.form.get("sub_hours") or "").strip()
    if name:
        sub.name = name
    try:
        if hrs_s != "":
            sub.hours = float(hrs_s)
    except Exception:
        flash("Invalid hours.", "warning"); return redirect(request.referrer or url_for("my_requests"))
    db.session.commit()
    flash("Substitute updated.", "success")
    return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/subs/<int:sub_id>/delete")
@login_required
def delete_substitute(req_id, sub_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    sub = SubAssignment.query.get_or_404(sub_id)
    if sub.request_id != req_id:
        abort(404)
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
    r.status = RequestStatus.approved
    r.decided_at = datetime.utcnow()
    db.session.commit()

    subs_text = "; ".join([f"{s.name} ({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "(none)")
    subj = "Leave Request Approved"
    body = (
        f"Hello {u.username},\n\n"
        f"Your leave request has been APPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"School-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\n"
        f"Dates: {r.start_date} to {r.end_date}\n\n"
        f"Remaining balance: {u.hours_balance:.2f} hours\n"
    )

    ok, emsg = send_email([u.email] + admin_emails(), subj, body)
    if not ok:
        flash(f"Notice: approval email not sent ({emsg}).", "warning")

    flash("Approved.", "success")
    return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/disapprove")
@login_required
def disapprove(req_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning"); return redirect(url_for("my_requests"))
    r = LeaveRequest.query.get_or_404(req_id)
    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning"); return redirect(url_for("my_requests"))
    r.status = RequestStatus.disapproved
    r.decided_at = datetime.utcnow()
    db.session.commit()

    u = User.query.get(r.user_id)
    subs_text = "; ".join([f"{s.name} ({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "(none)")
    subj = "Leave Request Disapproved"
    body = (
        f"Hello {u.username},\n\n"
        f"Your leave request has been DISAPPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"School-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
    )

    ok, emsg = send_email([u.email] + admin_emails(), subj, body)
    if not ok:
        flash(f"Notice: disapproval email not sent ({emsg}).", "warning")

    flash("Disapproved.", "info")
    return redirect(request.referrer or url_for("my_requests"))

@app.post("/requests/<int:req_id>/cancel")
@login_required
def cancel(req_id):
    r = LeaveRequest.query.get_or_404(req_id)
    if r.user_id != current_user.id and current_user.role != Role.admin:
        flash("Not allowed.", "danger")
        return redirect(url_for("my_requests"))
    u = User.query.get(r.user_id)
    if r.status == RequestStatus.approved and not r.is_school_related:
        u.hours_balance = float(u.hours_balance or 0.0) + float(r.hours or 0.0)
    r.status = RequestStatus.cancelled
    r.decided_at = datetime.utcnow()
    db.session.commit()

    subs_text = "; ".join([f"{s.name} ({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "(none)")
    subj = "Leave Request Cancelled"
    body = (
        f"User {u.username} cancelled a leave request.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours:.2f}\n"
        f"School-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
        f"Balance is now: {u.hours_balance:.2f} hours\n"
    )

    recipients = admin_emails()
    if u.email:
        recipients = [u.email] + recipients
    ok, emsg = send_email(recipients, subj, body)
    if not ok:
        flash(f"Notice: cancel email not sent ({emsg}).", "warning")

    flash("Cancelled.", "secondary")
    return redirect(request.referrer or url_for("my_requests"))

@app.route("/requests/<int:req_id>/edit", methods=["GET", "POST"])
@login_required
def edit_request(req_id):
    if not current_user.is_admin:  # Make sure only admins can edit
        abort(403)

    r = LeaveRequest.query.get_or_404(req_id)

    if request.method == "POST":
        # Update times/dates/hours
        r.start_date = request.form.get("start_date") or r.start_date
        r.end_date = request.form.get("end_date") or r.end_date
        r.start_time = request.form.get("start_time") or r.start_time
        r.end_time = request.form.get("end_time") or r.end_time
        r.hours = float(request.form.get("hours") or r.hours)

        db.session.commit()
        flash("Request updated successfully.", "success")
        return redirect(url_for("my_requests"))

    return render_template("edit_request.html", r=r)

@app.route("/user/<int:user_id>/requests")
@login_required
def user_requests(user_id):
    # Always get the user first — this ensures the variable exists
    user = User.query.get_or_404(user_id)
    is_admin = getattr(current_user, "role", "") == "admin"

    # Retrieve this user’s leave requests
    reqs = LeaveRequest.query.filter_by(user_id=user.id).order_by(LeaveRequest.start_date.desc()).all()

    # Try to fetch manual adjustments (if table exists)
    adjustments = []
    try:
        if "manual_adjustment" in db.metadata.tables:
            adjustments = ManualAdjustment.query.filter_by(user_id=user.id).order_by(ManualAdjustment.timestamp.desc()).all()
    except Exception as e:
        app.logger.warning(f"ManualAdjustment query failed: {e}")

    return render_template(
        "user_requests.html",
        user=user,
        requests=reqs,
        adjustments=adjustments,
        is_admin=is_admin,
        me=current_user
    )

@app.route("/user/<int:user_id>/add_time", methods=["POST"])
@login_required
def add_manual_time(user_id):
    if not getattr(current_user, "role", "") == "admin":
        flash("Admins only.", "danger")
        return redirect(url_for("my_requests"))

    user = User.query.get_or_404(user_id)
    hours = request.form.get("adjust_hours", type=float)
    note = request.form.get("note", "").strip()

    if hours is None or not note:
        flash("Please provide both hours and a note.", "warning")
        return redirect(url_for("user_requests", user_id=user_id))

    adj = ManualAdjustment(
        user_id=user.id,
        admin_id=current_user.id,
        hours=hours,
        note=note
    )
    db.session.add(adj)
    db.session.commit()

    flash(f"Adjustment of {hours:+.2f}h added for {user.username}.", "success")
    return redirect(url_for("user_requests", user_id=user_id))

@app.route("/user/<int:user_id>/requests/export")
@login_required
def export_user_requests(user_id):
    if not current_user.is_admin:
        abort(403)

    from io import StringIO
    import csv

    user = User.query.get_or_404(user_id)
    reqs = LeaveRequest.query.filter_by(user_id=user.id).order_by(LeaveRequest.start_date).all()

    # Create CSV in memory
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["ID", "Kind", "Mode", "Start Date", "End Date", "Start Time", "End Time", "Hours", "Status", "School Related", "Substitutes"])

    for r in reqs:
        subs = ", ".join([f"{s.name} ({s.hours}h)" for s in r.subs]) if r.subs else (r.substitute or "")
        writer.writerow([
            r.id, r.kind, r.mode, r.start_date, r.end_date,
            r.start_time.strftime("%I:%M %p") if r.start_time else "",
            r.end_time.strftime("%I:%M %p") if r.end_time else "",
            f"{r.hours:.2f}", r.status, "Yes" if r.is_school_related else "No", subs
        ])

    output.seek(0)
    filename = f"{user.username}_leave_history.csv"
    return Response(
        output.getvalue(),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )
# ---------- Manage Users (admin) ----------
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

    try:
        hours_balance = float(hours_str) if hours_str else 160.0
    except Exception:
        hours_balance = 160.0

    user = User(
        username=username,
        staff_name=staff_name or None,
        password_hash=generate_password_hash(pw),
        role=role if role in (Role.admin, Role.staff) else Role.staff,
        hours_balance=hours_balance,
        email=email or None
    )
    db.session.add(user)
    db.session.commit()
    flash(f"User '{username}' created.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/update")
@login_required
def admin_update_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))
    u = User.query.get_or_404(user_id)

    # Username change intentionally omitted
    staff_name = (request.form.get("staff_name") or "").strip()
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or "").strip()
    hb = (request.form.get("hours_balance") or "").strip()

    u.staff_name = staff_name or None
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
    db.session.delete(u)
    db.session.commit()
    flash("User deleted.", "success")
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
def sub_summary_text(subs, limit=2):
    """Return a compact summary like 'Sub: A(4h), B(3h) +1 more'."""
    if not subs: return ""
    parts = [f"{s.name}({s.hours:.1f}h)" for s in subs[:limit]]
    more = len(subs) - limit
    tail = f" +{more} more" if more > 0 else ""
    return " – Sub: " + ", ".join(parts) + tail

@app.get("/calendar")
@login_required
def calendar():
    return render_template("calendar.html", title="Calendar", is_admin=(current_user.role == Role.admin), me=current_user)

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
            if not sub_text and (r.substitute or "").strip():
                sub_text = " – Sub: " + r.substitute.strip()
            title += sub_text
        else:
            title = f"{r.kind} ({r.hours:.1f}h)"
        if r.is_school_related:
            title = "[School] " + title

        events.append({
            "title": title,
            "start": r.start_date.isoformat(),
            "end": (r.end_date + timedelta(days=1)).isoformat(),  # exclusive end
            "allDay": True,
        })
    return jsonify(events)

# ---------- Exports (admin only) ----------
@app.get("/admin/export/requests.csv")
@login_required
def export_requests_csv():
    if current_user.role != Role.admin:
        abort(403)
    rows = _filtered_requests_for(True).all()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "ID","Username","StaffName","Kind","Mode","Hours","Status","Start","End",
        "StartTime","EndTime","SchoolRelated","Substitutes","Created","Decided"
    ])
    for r in rows:
        subs_text = "; ".join([f"{s.name}({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "")
        writer.writerow([
            r.id, r.user.username, (r.user.staff_name or ""),
            r.kind, r.mode, f"{r.hours:.2f}", r.status,
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
    headers = ["ID","Username","StaffName","Kind","Mode","Hours","Status","Start","End",
               "StartTime","EndTime","SchoolRelated","Substitutes","Created","Decided"]
    hdr_fmt = wb.add_format({"bold": True, "bg_color": "#F1F5F9", "border": 1})
    cell_fmt = wb.add_format({"border": 1})
    date_fmt = wb.add_format({"num_format": "yyyy-mm-dd", "border": 1})
    dt_fmt = wb.add_format({"num_format": "yyyy-mm-dd hh:mm", "border": 1})

    for c, h in enumerate(headers):
        ws.write(0, c, h, hdr_fmt)

    rix = 1
    for r in rows:
        subs_text = "; ".join([f"{s.name}({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "")
        ws.write(rix, 0, r.id, cell_fmt)
        ws.write(rix, 1, r.user.username, cell_fmt)
        ws.write(rix, 2, (r.user.staff_name or ""), cell_fmt)
        ws.write(rix, 3, r.kind, cell_fmt)
        ws.write(rix, 4, r.mode, cell_fmt)
        ws.write_number(rix, 5, float(r.hours or 0.0), cell_fmt)
        ws.write(rix, 6, r.status, cell_fmt)

        ws.write_datetime(rix, 7, datetime.combine(r.start_date, datetime.min.time()), date_fmt)
        ws.write_datetime(rix, 8, datetime.combine(r.end_date, datetime.min.time()), date_fmt)

        ws.write(rix, 9, r.start_time or "", cell_fmt)
        ws.write(rix, 10, r.end_time or "", cell_fmt)

        ws.write(rix, 11, "Yes" if r.is_school_related else "No", cell_fmt)
        ws.write(rix, 12, subs_text, cell_fmt)

        if r.created_at:
            ws.write_datetime(rix, 13, r.created_at, dt_fmt)
        else:
            ws.write(rix, 13, "", cell_fmt)
        if r.decided_at:
            ws.write_datetime(rix, 14, r.decided_at, dt_fmt)
        else:
            ws.write(rix, 14, "", cell_fmt)

        rix += 1

    # autosize some columns
    widths = [len(h) for h in headers]
    for r in rows:
        widths[1] = max(widths[1], len(r.user.username or ""))
        widths[2] = max(widths[2], len(r.user.staff_name or ""))
        widths[3] = max(widths[3], len(r.kind or ""))
        widths[4] = max(widths[4], len(r.mode or ""))
        widths[6] = max(widths[6], len(r.status or ""))

    for c, w in enumerate(widths):
        ws.set_column(c, c, min(max(w + 2, 10), 32))

    # Sheet 2: Substitutes
    ws2 = wb.add_worksheet("Substitutes")
    ws2_headers = ["RequestID","Username","StaffName","Start","End","Substitute","Hours"]
    for c, h in enumerate(ws2_headers):
        ws2.write(0, c, h, hdr_fmt)
    rix = 1
    for r in rows:
        for s in r.subs:
            ws2.write(rix, 0, r.id, cell_fmt)
            ws2.write(rix, 1, r.user.username, cell_fmt)
            ws2.write(rix, 2, (r.user.staff_name or ""), cell_fmt)
            ws2.write_datetime(rix, 3, datetime.combine(r.start_date, datetime.min.time()), date_fmt)
            ws2.write_datetime(rix, 4, datetime.combine(r.end_date, datetime.min.time()), date_fmt)
            ws2.write(rix, 5, s.name, cell_fmt)
            ws2.write_number(rix, 6, float(s.hours or 0.0), cell_fmt)
            rix += 1

    wb.close()
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="leave_requests.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ---------- Monthly Report (Option 2) ----------
@app.get("/admin/export/monthly")
@login_required
def export_monthly():
    if current_user.role != Role.admin:
        abort(403)

    # Current month range
    today = date.today()
    start = date(today.year, today.month, 1)
    if today.month == 12:
        end = date(today.year + 1, 1, 1) - timedelta(days=1)
    else:
        end = date(today.year, today.month + 1, 1) - timedelta(days=1)

    # Optional override via ?start=YYYY-MM-DD&end=YYYY-MM-DD
    def parse_date_q(s):
        try:
            return datetime.strptime(s, "%Y-%m-%d").date()
        except Exception:
            return None
    qstart = parse_date_q(request.args.get("start", ""))
    qend = parse_date_q(request.args.get("end", ""))
    if qstart: start = qstart
    if qend: end = qend

    rows = (LeaveRequest.query
            .filter(LeaveRequest.start_date >= start, LeaveRequest.end_date <= end)
            .order_by(LeaveRequest.user_id, LeaveRequest.start_date)
            .all())

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Username","StaffName","Kind","Mode","Hours","Status","Start","End","School Related","Substitutes"])
    for r in rows:
        subs = "; ".join([f"{s.name}({s.hours:.2f}h)" for s in r.subs]) or (r.substitute or "")
        writer.writerow([
            r.user.username,
            (r.user.staff_name or ""),
            r.kind,
            r.mode,
            f"{r.hours:.2f}",
            r.status,
            r.start_date.isoformat(),
            r.end_date.isoformat(),
            "Yes" if r.is_school_related else "No",
            subs
        ])

    resp = make_response(output.getvalue())
    resp.headers["Content-Type"] = "text/csv"
    resp.headers["Content-Disposition"] = f"attachment; filename=leave_report_{start.strftime('%Y_%m')}.csv"
    return resp

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

# Dev server entry (ignored by gunicorn)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
