from flask import Flask, render_template, redirect, url_for, request, flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
from werkzeug.security import generate_password_hash, check_password_hash
import os

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "ChangeThisSecret123!")

# DB: Render provides DATABASE_URL for Postgres; otherwise use a local SQLite file
db_url = os.environ.get("DATABASE_URL", "sqlite:///leave_system.db")
if db_url.startswith("postgres://"):
    db_url = db_url.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = db_url
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"

# --- Models ---
class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(20), default="admin")
    hours_balance = db.Column(db.Float, default=160.0)

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    start_date = db.Column(db.String(50), nullable=False)
    end_date = db.Column(db.String(50), nullable=False)
    hours_requested = db.Column(db.Float, nullable=False)
    status = db.Column(db.String(20), default="Pending")
    user = db.relationship("User", backref="leave_requests", lazy="joined")

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

def ensure_db():
    db.create_all()
    if not User.query.filter_by(username="mc-admin").first():
        admin = User(username="mc-admin",
                     password_hash=generate_password_hash("RWAadmin2"),
                     role="admin",
                     hours_balance=160.0)
        db.session.add(admin); db.session.commit()

with app.app_context():
    ensure_db()

# --- Routes ---
@app.route("/")
def home():
    return redirect(url_for("login"))

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username","").strip()
        password = request.form.get("password","")
        user = User.query.filter(User.username.ilike(username)).first()
        if user and check_password_hash(user.password_hash, password):
            login_user(user)
            flash("Logged in.", "success")
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "danger")
    return render_template("login.html")

@app.route("/dashboard")
@login_required
def dashboard():
    return render_template("dashboard.html")

@app.route("/request_leave", methods=["GET","POST"])
@login_required
def request_leave():
    if request.method == "POST":
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        hours_requested = float(request.form.get("hours_requested") or 0)
        if hours_requested <= 0:
            flash("Hours requested must be greater than zero.", "warning")
            return render_template("leave_request.html")
        lr = LeaveRequest(user_id=current_user.id,
                          start_date=start_date,
                          end_date=end_date,
                          hours_requested=hours_requested)
        db.session.add(lr); db.session.commit()
        flash("Leave request submitted.", "success")
        return redirect(url_for("dashboard"))
    return render_template("leave_request.html")

@app.route("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login"))

@app.get("/health")
def health():
    return "ok", 200

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
    app.run()
