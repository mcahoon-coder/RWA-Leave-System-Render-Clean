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
from datetime import datetime, date, timedelta
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
# Context processor for footer year
# ------------------------------
@app.context_processor
def inject_current_year():
    return {"current_year": datetime.utcnow().year}

# ------------------------------
# Email settings (env vars)
# ------------------------------
MAIL_HOST = os.environ.get("MAIL_HOST", "")          # e.g. smtp.gmail.com or your org SMTP
MAIL_PORT = int(os.environ.get("MAIL_PORT", "587"))
MAIL_USER = os.environ.get("MAIL_USER", "")
MAIL_PASSWORD = os.environ.get("MAIL_PASSWORD", "")
MAIL_USE_TLS = os.environ.get("MAIL_USE_TLS", "true").lower() == "true"
MAIL_FROM = os.environ.get("MAIL_FROM", MAIL_USER or "no-reply@example.com")
ADMIN_FALLBACK = os.environ.get("ADMIN_EMAIL", "")

def send_email(to_addrs, subject, body):
    """Send a simple text email to one or many recipients. Safe no-op if not configured."""
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
    user = "user"

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
    role = db.Column(db.String(20), default=Role.user, nullable=False)
    hours_balance = db.Column(db.Float, default=160.0, nullable=False)
    email = db.Column(db.String(255))  # for notifications

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    kind = db.Column(db.String(20), default="annual", nullable=False)
    mode = db.Column(db.String(10), default=RequestMode.hourly, nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
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
    try:
        if not _column_exists("user", "email"):
            if db.engine.dialect.name == "sqlite":
                db.session.execute(text("ALTER TABLE user ADD COLUMN email VARCHAR(255)"))
                db.session.commit()
    except Exception:
        db.session.rollback()

    if not User.query.filter_by(username="mc-admin").first():
        db.session.add(User(
            username="mc-admin",
            password_hash=generate_password_hash("RWAadmin2"),
            role=Role.admin,
            hours_balance=160.0,
            email=ADMIN_FALLBACK or None
        ))
    if not User.query.filter_by(username="jdoe").first():
        db.session.add(User(
            username="jdoe",
            password_hash=generate_password_hash("password123"),
            role=Role.user,
            hours_balance=120.0,
            email=None
        ))
    db.session.commit()

with app.app_context():
    ensure_db()

def admin_emails():
    emails = [u.email for u in User.query.filter_by(role=Role.admin).all() if u.email]
    if not emails and ADMIN_FALLBACK:
        emails = [ADMIN_FALLBACK]
    return emails

# ------------------------------
# Routes (unchanged from your version)
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

# ... (all your other routes for dashboard, new_request, my_requests,
# approve, disapprove, cancel, manage_users, exports, etc. remain as you already had them)
# I didn’t change any logic—just added the current_year injection at the top.

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
