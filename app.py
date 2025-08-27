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
from datetime import datetime, date, timedelta, time as dtime
import os, smtplib, ssl, io, csv
from email.message import EmailMessage
from sqlalchemy import text
import xlsxwriter  # for Excel export (uses memory, safe on Render)

# ------------------------------
# App & DB config
# ------------------------------
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ChangeThisSecret123!")

# Prefer Render DATABASE_URL; default to SQLite
db_url = os.environ.get("DATABASE_URL", "sqlite:///leave_system.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)

login_manager = LoginManager(app)
login_manager.login_view = "login"

# ------------------------------
# Email settings (env vars)
# ------------------------------
MAIL_HOST = os.environ.get("MAIL_HOST", "")
MAIL_PORT = int(os.environ.get("MAIL_PORT", "587"))
MAIL_USER = os.environ.get("MAIL_USER", "")
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD", "")
MAIL_USE_TLS = os.environ.get("MAIL_USE_TLS", "true").lower() == "true"
MAIL_FROM = os.environ.get("MAIL_FROM", MAIL_USER or "no-reply@example.com")
ADMIN_FALLBACK = os.environ.get("ADMIN_EMAIL", "")

def send_email(to_addrs, subject, body):
    if not to_addrs:
        return
    if isinstance(to_addrs, str):
        to_addrs = [to_addrs]
    if not MAIL_HOST or not MAIL_FROM:
        return  # SMTP not configured

    msg = EmailMessage()
    msg["From"] = MAIL_FROM
    msg["To"] = ", ".join([a for a in to_addrs if a])
    msg["Subject"] = subject
    msg.set_content(body)
    if not msg["To"]:
        return

    if MAIL_USE_TLS:
        context = ssl.create_default_context()
        with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            if MAIL_USER:
                server.login(MAIL_USER, MAIL_PASSWORD)
            server.send_message(msg)
    else:
        with smtplib.SMTP(MAIL_HOST, MAIL_PORT) as server:
            if MAIL_USER:
                server.login(MAIL_USER, MAIL_PASSWORD)
            server.send_message(msg)

# ------------------------------
# Models & constants
# ------------------------------
class Role:
    admin = "admin"
    faculty_staff = "faculty_staff"  # standard user

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
    role = db.Column(db.String(20), default=Role.faculty_staff, nullable=False)
    hours_balance = db.Column(db.Float, default=160.0, nullable=False)
    email = db.Column(db.String(255))

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    kind = db.Column(db.String(20), default="annual", nullable=False)   # 'annual' or 'sick'
    mode = db.Column(db.String(10), default=RequestMode.hourly, nullable=False)  # hourly/daily
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    # NEW: optional timed leave in 15-min increments (for hourly mode)
    start_time = db.Column(db.Time)   # nullable
    end_time = db.Column(db.Time)     # nullable
    hours = db.Column(db.Float, nullable=False)
    reason = db.Column(db.String(500), default="")
    status = db.Column(db.String(20), default=RequestStatus.pending, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    decided_at = db.Column(db.DateTime)

    user = db.relationship("User", backref="leave_requests", lazy="joined")

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

# ------------------------------
# Helpers
# ------------------------------
WORKDAY_HOURS = float(os.environ.get("WORKDAY_HOURS", "8"))
HOLIDAYS: set[date] = set()  # add date(...) here if you want static holidays

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in HOLIDAYS  # Mon–Fri & not holiday

def workdays_between(start: date, end: date) -> int:
    if end < start:
        start, end = end, start
    n, cur = 0, start
    while cur <= end:
        if is_workday(cur):
            n += 1
        cur += timedelta(days=1)
    return n

def _column_exists(table_name: str, column_name: str) -> bool:
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
     db.create_all()

     # Add email column if missing (SQLite quick helper only)
     try:
         if db.engine.dialect.name == "sqlite" and not _column_exists("user", "email"):
             db.session.execute(text("ALTER TABLE user ADD COLUMN email VARCHAR(255)"))
            db.session.commit()
    except Exception:
        db.session.rollback()

    # Seed ONCE only if DB is empty
    if User.query.count() == 0:
        db.session.add(User(
            username="mc-admin",
            password_hash=generate_password_hash("RWAadmin2"),
            role=Role.admin,
            hours_balance=160.0,
            email=ADMIN_FALLBACK or None
        ))
        db.session.commit()

with app.app_context():
    ensure_db()

def admin_emails():
    emails = [u.email for u in User.query.filter_by(role=Role.admin).all() if u.email]
    if not emails and ADMIN_FALLBACK:
        emails = [ADMIN_FALLBACK]
    return emails

# Build list of 15-minute time options as "HH:MM"
TIME_SLOTS = [f"{h:02d}:{m:02d}" for h in range(0, 24) for m in (0, 15, 30, 45)]

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

# ------------------------------
# Routes
# ------------------------------
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

@app.route("/request/new", methods=["GET", "POST"])
@login_required
def new_request():
    if request.method == "POST":
        mode = request.form.get("mode", RequestMode.hourly)
        kind = request.form.get("kind", "annual")
        reason = request.form.get("reason", "")

        # Parse dates
        try:
            sd = datetime.strptime(request.form["start_date"], "%Y-%m-%d").date()
            ed = datetime.strptime(request.form["end_date"], "%Y-%m-%d").date()
        except Exception:
            flash("Invalid dates.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)

        capacity_hours = workdays_between(sd, ed) * WORKDAY_HOURS
        if capacity_hours <= 0:
            flash("No working days in that range.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)

        # Determine hours
        hours = 0.0
        start_t_obj = None
        end_t_obj = None

        if mode == RequestMode.hourly:
            use_times = request.form.get("use_times", "no") == "yes"
            if use_times:
                st_s = request.form.get("start_time", "").strip()
                et_s = request.form.get("end_time", "").strip()
                # Require same day for time-window requests
                if sd != ed:
                    flash("For start/end times, start and end date must be the same day.", "warning")
                    return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)
                try:
                    start_t_obj = datetime.strptime(st_s, "%H:%M").time()
                    end_t_obj = datetime.strptime(et_s, "%H:%M").time()
                except Exception:
                    flash("Invalid time selection.", "warning")
                    return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)
                start_dt = datetime.combine(sd, start_t_obj)
                end_dt = datetime.combine(ed, end_t_obj)
                delta_h = (end_dt - start_dt).total_seconds() / 3600.0
                if delta_h <= 0:
                    flash("End time must be after start time.", "warning")
                    return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)
                # 15-min increments already enforced by dropdowns; still clamp to 2 decimals
                hours = round(delta_h, 2)
            else:
                # Manual hourly entry
                try:
                    hours = float(request.form.get("hours", "0"))
                except Exception:
                    hours = 0.0
        else:
            wd = workdays_between(sd, ed)
            hours = wd * WORKDAY_HOURS

        if hours <= 0:
            flash("Requested hours must be greater than zero.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)

        if hours > capacity_hours:
            flash(f"Requested {hours:.2f} exceeds capacity {capacity_hours:.2f} for that range.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)

        req = LeaveRequest(
            user_id=current_user.id,
            kind=kind,
            mode=mode,
            start_date=sd,
            end_date=ed,
            start_time=start_t_obj,
            end_time=end_t_obj,
            hours=hours,
            reason=reason
        )
        db.session.add(req)
        db.session.commit()

        # Notify admins
        subj = "New Leave Request Submitted"
        timed = (f"{req.start_time.strftime('%H:%M')}–{req.end_time.strftime('%H:%M')} on {sd}" if req.start_time and req.end_time else f"{sd} to {ed}")
        body = (
            f"User: {current_user.username}\n"
            f"Kind: {kind}\nMode: {mode}\nHours: {hours}\n"
            f"When: {timed}\nReason: {reason or '(none)'}\n"
            f"Status: {req.status}\n"
        )
        send_email(admin_emails(), subj, body)

        flash("Request submitted.", "success")
        return redirect(url_for("my_requests"))

    return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS, time_slots=TIME_SLOTS)

@app.get("/requests")
@login_required
def my_requests():
    q = _filtered_requests_for(current_user.role == Role.admin)
    reqs = q.all()
    return render_template(
        "requests.html",
        title="Requests",
        reqs=reqs,
        is_admin=(current_user.role == Role.admin),
        status=request.args.get("status", "all"),
        start=request.args.get("start", ""),
        end=request.args.get("end", "")
    )

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
    # Allow negative balances
    u.hours_balance = float(u.hours_balance or 0.0) - float(r.hours or 0.0)
    r.status = RequestStatus.approved
    r.decided_at = datetime.utcnow()
    db.session.commit()

    subj = "Leave Request Approved"
    body = (
        f"Hello {u.username},\n\n"
        f"Your leave request has been APPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
        + (f"Times: {r.start_time.strftime('%H:%M')}–{r.end_time.strftime('%H:%M')}\n" if r.start_time and r.end_time else "")
        + f"\nRemaining balance: {u.hours_balance:.1f} hours\n"
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
        f"Hello {u.username},\n\n"
        f"Your leave request has been DISAPPROVED.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
        + (f"Times: {r.start_time.strftime('%H:%M')}–{r.end_time.strftime('%H:%M')}\n" if r.start_time and r.end_time else "")
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
        f"User {u.username} cancelled a leave request.\n"
        f"Kind: {r.kind}\nMode: {r.mode}\nHours: {r.hours}\n"
        f"Dates: {r.start_date} to {r.end_date}\n"
        + (f"Times: {r.start_time.strftime('%H:%M')}–{r.end_time.strftime('%H:%M')}\n" if r.start_time and r.end_time else "")
        + f"Balance is now: {u.hours_balance:.1f} hours\n"
    )
    recipients = admin_emails()
    if u.email:
        recipients = [u.email] + recipients
    send_email(recipients, subj, body)

    flash("Cancelled.", "secondary")
    return redirect(url_for("my_requests"))

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

    username = request.form.get("username", "").strip()
    password = request.form.get("password", "").strip()
    email = request.form.get("email", "").strip() or None
    role = request.form.get("role", Role.faculty_staff)
    hours_balance = request.form.get("hours_balance", "160.0").strip()

    if not username or not password:
        flash("Username and password are required.", "warning")
        return redirect(url_for("manage_users"))

    if User.query.filter_by(username=username).first():
        flash("Username already exists.", "danger")
        return redirect(url_for("manage_users"))

    try:
        hb = float(hours_balance)
    except Exception:
        hb = 160.0

    u = User(
        username=username,
        password_hash=generate_password_hash(password),
        role=role if role in (Role.admin, Role.faculty_staff) else Role.faculty_staff,
        hours_balance=hb,
        email=email
    )
    db.session.add(u)
    db.session.commit()
    flash(f"User '{username}' created.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/update-all")
@login_required
def admin_update_user_all(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))

    u = User.query.get_or_404(user_id)
    new_username = request.form.get("username", "").strip()
    email = request.form.get("email", "").strip() or None
    role = request.form.get("role", Role.faculty_staff)
    hours_balance_s = request.form.get("hours_balance", "").strip()

    if new_username:
        if new_username != u.username and User.query.filter_by(username=new_username).first():
            flash("That username is already taken.", "danger")
            return redirect(url_for("manage_users"))
        u.username = new_username

    u.email = email
    if role in (Role.admin, Role.faculty_staff):
        u.role = role

    try:
        if hours_balance_s != "":
            u.hours_balance = float(hours_balance_s)
    except Exception:
        flash("Invalid hours balance value.", "warning")

    db.session.commit()
    flash(f"Updated user '{u.username}'.", "success")
    return redirect(url_for("manage_users"))

@app.post("/admin/users/<int:user_id>/reset")
@login_required
def admin_reset_password(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))
    new_pw = request.form.get("new_password", "").strip()
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
        new = request.form.get("new_password", "").strip()
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
        if r.start_time and r.end_time and r.start_date == r.end_date:
            start_dt = datetime.combine(r.start_date, r.start_time)
            end_dt = datetime.combine(r.end_date, r.end_time)
            events.append({
                "title": f"{r.user.username} - {r.kind} ({r.hours:.1f}h)",
                "start": start_dt.isoformat(),
                "end": end_dt.isoformat()
            })
        else:
            events.append({
                "title": f"{r.user.username} - {r.kind} ({r.hours:.1f}h)",
                "start": r.start_date.isoformat(),
                "end": (r.end_date + timedelta(days=1)).isoformat()
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
    writer.writerow(["ID","Username","Kind","Mode","Hours","Status","Start","End","Start Time","End Time","Created","Decided"])
    for r in rows:
        writer.writerow([
            r.id, r.user.username, r.kind, r.mode, f"{r.hours:.2f}", r.status,
            r.start_date.isoformat(), r.end_date.isoformat(),
            r.start_time.strftime("%H:%M") if r.start_time else "",
            r.end_time.strftime("%H:%M") if r.end_time else "",
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

    headers = ["ID","Username","Kind","Mode","Hours","Status","Start","End","Start Time","End Time","Created","Decided"]
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
        ws.write(rix, 8, r.start_time.strftime("%H:%M") if r.start_time else "", cell_fmt)
        ws.write(rix, 9, r.end_time.strftime("%H:%M") if r.end_time else "", cell_fmt)
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
    for row in rows:
        widths[1] = max(widths[1], len(row.user.username or ""))
        widths[2] = max(widths[2], len(row.kind or ""))
        widths[3] = max(widths[3], len(row.mode or ""))
        widths[5] = max(widths[5], len(row.status or ""))

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
    app.run(host="0.0.0.0", port=5000)
