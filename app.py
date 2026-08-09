import json
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
from sqlalchemy import text, func
import xlsxwriter  # Excel export (in-memory, safe on Render)
from google.oauth2 import service_account
from googleapiclient.discovery import build

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
# Google Sheets settings
# =========================================================
GOOGLE_SHEETS_ENABLED = os.environ.get(
    "GOOGLE_SHEETS_ENABLED", "FALSE"
).lower() in ("true", "1", "yes")

GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get(
    "GOOGLE_SERVICE_ACCOUNT_JSON", ""
).strip()

GOOGLE_SERVICE_ACCOUNT_FILE = os.environ.get(
    "GOOGLE_SERVICE_ACCOUNT_FILE",
    "/etc/secrets/google-service-account.json",
).strip()

GOOGLE_REPORT_FOLDER_ID = os.environ.get(
    "GOOGLE_REPORT_FOLDER_ID", ""
).strip()

GOOGLE_REPORT_SHARE_EMAIL = os.environ.get(
    "GOOGLE_REPORT_SHARE_EMAIL", ""
).strip()

GOOGLE_API_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _load_google_service_account_info():
    """
    Load Google service-account credentials from Render Secret Files first.
    Falls back to GOOGLE_SERVICE_ACCOUNT_JSON if needed.
    """
    if (
        GOOGLE_SERVICE_ACCOUNT_FILE
        and os.path.exists(GOOGLE_SERVICE_ACCOUNT_FILE)
    ):
        try:
            with open(
                GOOGLE_SERVICE_ACCOUNT_FILE,
                "r",
                encoding="utf-8",
            ) as credential_file:
                return json.load(credential_file), "secret_file"
        except Exception as exc:
            app.logger.exception(
                "Could not read Google service-account secret file"
            )
            return None, f"secret_file_error: {exc}"

    if GOOGLE_SERVICE_ACCOUNT_JSON:
        try:
            return (
                json.loads(GOOGLE_SERVICE_ACCOUNT_JSON),
                "environment",
            )
        except Exception as exc:
            app.logger.exception(
                "Could not parse GOOGLE_SERVICE_ACCOUNT_JSON"
            )
            return None, f"environment_error: {exc}"

    return None, "missing"


def google_sheets_configuration_status():
    missing = []

    credentials_info, credential_source = (
        _load_google_service_account_info()
    )

    credentials_valid = False
    service_account_email = None

    if credentials_info:
        service_account_email = credentials_info.get("client_email")
        credentials_valid = bool(
            credentials_info.get("type") == "service_account"
            and credentials_info.get("client_email")
            and credentials_info.get("private_key")
            and credentials_info.get("token_uri")
        )

    if not credentials_valid:
        if credential_source == "missing":
            missing.append("Google service-account secret file")
        else:
            missing.append("valid service-account JSON")

    return {
        "enabled": GOOGLE_SHEETS_ENABLED,
        "ready": GOOGLE_SHEETS_ENABLED and credentials_valid,
        "missing": missing,
        "folder_id": GOOGLE_REPORT_FOLDER_ID,
        "share_email": GOOGLE_REPORT_SHARE_EMAIL,
        "service_account_email": service_account_email,
        "credentials_valid": credentials_valid,
        "credential_source": credential_source,
        "credential_file": GOOGLE_SERVICE_ACCOUNT_FILE,
    }


def get_google_credentials():
    status = google_sheets_configuration_status()

    if not status["ready"]:
        raise RuntimeError(
            "Google Sheets is disabled or its Google credentials "
            "are incomplete."
        )

    credentials_info, _ = _load_google_service_account_info()

    if not credentials_info:
        raise RuntimeError(
            "Google service-account credentials could not be loaded."
        )

    return service_account.Credentials.from_service_account_info(
        credentials_info,
        scopes=GOOGLE_API_SCOPES,
    )

def google_services():
    credentials = get_google_credentials()
    sheets_service = build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )
    drive_service = build(
        "drive",
        "v3",
        credentials=credentials,
        cache_discovery=False,
    )
    return sheets_service, drive_service


def month_date_range(year, month):
    start = date(year, month, 1)
    if month == 12:
        next_month = date(year + 1, 1, 1)
    else:
        next_month = date(year, month + 1, 1)
    return start, next_month - timedelta(days=1)


def format_report_date(value):
    return value.strftime("%Y-%m-%d") if value else ""


def format_report_datetime(value):
    return value.strftime("%Y-%m-%d %I:%M %p") if value else ""


def build_google_report_data(report_year, report_month):
    start_date, end_date = month_date_range(report_year, report_month)

    requests_rows = (
        LeaveRequest.query
        .filter(
            LeaveRequest.start_date <= end_date,
            LeaveRequest.end_date >= start_date,
        )
        .order_by(LeaveRequest.start_date, LeaveRequest.user_id)
        .all()
    )

    approved_requests = [
        item for item in requests_rows
        if item.status == RequestStatus.approved
    ]

    users = User.query.order_by(
        func.coalesce(User.staff_name, User.username)
    ).all()

    beginning_by_user = {}
    active_year = get_active_school_year()
    if active_year:
        beginning_by_user = {
            row.user_id: float(row.beginning_balance or 0.0)
            for row in SchoolYearBalance.query.filter_by(
                school_year_id=active_year.id
            ).all()
        }

    monthly_used_by_user = {}
    monthly_school_by_user = {}

    for leave_request in approved_requests:
        hours = float(leave_request.hours or 0.0)
        target = monthly_school_by_user if leave_request.is_school_related else monthly_used_by_user
        target[leave_request.user_id] = target.get(leave_request.user_id, 0.0) + hours

    manual_rows = (
        ManualAdjustment.query
        .filter(
            ManualAdjustment.timestamp >= datetime.combine(start_date, datetime.min.time()),
            ManualAdjustment.timestamp < datetime.combine(
                end_date + timedelta(days=1), datetime.min.time()
            ),
        )
        .order_by(ManualAdjustment.timestamp)
        .all()
    )

    manual_by_user = {}
    for adjustment in manual_rows:
        manual_by_user[adjustment.user_id] = (
            manual_by_user.get(adjustment.user_id, 0.0)
            + float(adjustment.hours or 0.0)
        )

    summary_values = [[
        "Employee", "Beginning Balance", "Manual Adjustments This Month",
        "Leave Used This Month", "School-Related Hours", "Current Balance",
    ]]
    for employee in users:
        summary_values.append([
            employee.staff_name or employee.username,
            round(beginning_by_user.get(employee.id, 0.0), 2),
            round(manual_by_user.get(employee.id, 0.0), 2),
            round(monthly_used_by_user.get(employee.id, 0.0), 2),
            round(monthly_school_by_user.get(employee.id, 0.0), 2),
            round(float(employee.hours_balance or 0.0), 2),
        ])

    leave_detail_values = [[
        "Request ID", "Date Submitted", "Employee", "Kind", "Leave Type",
        "Start Date", "End Date", "Start Time", "End Time", "Hours",
        "Status", "School Related", "Reason",
    ]]
    for leave_request in requests_rows:
        leave_type = (
            "Full Day(s)" if leave_request.mode == RequestMode.daily
            else "Hourly" if leave_request.mode == RequestMode.hourly
            else leave_request.mode
        )
        leave_detail_values.append([
            leave_request.id,
            format_report_datetime(leave_request.created_at),
            leave_request.user.staff_name or leave_request.user.username,
            leave_request.kind,
            leave_type,
            format_report_date(leave_request.start_date),
            format_report_date(leave_request.end_date),
            leave_request.start_time or "",
            leave_request.end_time or "",
            round(float(leave_request.hours or 0.0), 2),
            leave_request.status,
            "Yes" if leave_request.is_school_related else "No",
            leave_request.reason or "",
        ])

    substitute_values = [[
        "Leave Request ID", "Coverage Date", "Employee Out",
        "Substitute/Coverage Person", "Hours", "Leave Status", "School Related",
    ]]
    for leave_request in requests_rows:
        for substitute in leave_request.subs:
            substitute_values.append([
                leave_request.id,
                format_report_date(leave_request.start_date),
                leave_request.user.staff_name or leave_request.user.username,
                substitute.name,
                round(float(substitute.hours or 0.0), 2),
                leave_request.status,
                "Yes" if leave_request.is_school_related else "No",
            ])

    school_related_values = [[
        "Request ID", "Employee", "Start Date", "End Date",
        "Hours", "Status", "Reason",
    ]]
    for leave_request in requests_rows:
        if leave_request.is_school_related:
            school_related_values.append([
                leave_request.id,
                leave_request.user.staff_name or leave_request.user.username,
                format_report_date(leave_request.start_date),
                format_report_date(leave_request.end_date),
                round(float(leave_request.hours or 0.0), 2),
                leave_request.status,
                leave_request.reason or "",
            ])

    adjustment_values = [["Date", "Employee", "Hours", "Note", "Entered By"]]
    for adjustment in manual_rows:
        adjustment_values.append([
            format_report_datetime(adjustment.timestamp),
            adjustment.user.staff_name or adjustment.user.username,
            round(float(adjustment.hours or 0.0), 2),
            adjustment.note,
            (adjustment.admin.staff_name or adjustment.admin.username)
            if adjustment.admin else "",
        ])

    ledger_values = [[
        "Date", "Employee", "Entry Type", "Hours", "Description", "Entered By",
    ]]
    if active_year:
        ledger_rows = (
            LeaveLedger.query
            .filter(
                LeaveLedger.school_year_id == active_year.id,
                LeaveLedger.created_at >= datetime.combine(start_date, datetime.min.time()),
                LeaveLedger.created_at < datetime.combine(
                    end_date + timedelta(days=1), datetime.min.time()
                ),
            )
            .order_by(LeaveLedger.created_at)
            .all()
        )
        for entry in ledger_rows:
            ledger_values.append([
                format_report_datetime(entry.created_at),
                entry.user.staff_name or entry.user.username,
                ledger_entry_label(entry.entry_type),
                round(float(entry.hours or 0.0), 2),
                entry.description or "",
                (entry.created_by.staff_name or entry.created_by.username)
                if entry.created_by else "System",
            ])

    month_name = start_date.strftime("%B %Y")
    return {
        "title": f"RWA Leave Report - {month_name}",
        "start_date": start_date,
        "end_date": end_date,
        "sheets": [
            ("Employee Summary", summary_values),
            ("Leave Detail", leave_detail_values),
            ("Substitute Hours", substitute_values),
            ("School Related", school_related_values),
            ("Manual Adjustments", adjustment_values),
            ("Leave Ledger", ledger_values),
        ],
    }


def create_google_sheet_report(report_year, report_month):
    report_data = build_google_report_data(report_year, report_month)
    sheets_service, drive_service = google_services()

    spreadsheet_body = {
        "properties": {"title": report_data["title"]},
        "sheets": [
            {"properties": {"title": sheet_name}}
            for sheet_name, _ in report_data["sheets"]
        ],
    }

    spreadsheet = (
        sheets_service.spreadsheets()
        .create(body=spreadsheet_body, fields="spreadsheetId,spreadsheetUrl")
        .execute()
    )

    spreadsheet_id = spreadsheet["spreadsheetId"]
    spreadsheet_url = spreadsheet["spreadsheetUrl"]

    metadata = (
        sheets_service.spreadsheets()
        .get(spreadsheetId=spreadsheet_id, fields="sheets(properties(sheetId,title))")
        .execute()
    )
    sheet_ids = {
        sheet["properties"]["title"]: sheet["properties"]["sheetId"]
        for sheet in metadata.get("sheets", [])
    }

    value_data = []
    formatting_requests = []
    for sheet_name, values in report_data["sheets"]:
        value_data.append({
            "range": f"'{sheet_name}'!A1",
            "majorDimension": "ROWS",
            "values": values,
        })
        sheet_id = sheet_ids[sheet_name]
        column_count = max((len(row) for row in values), default=1)
        formatting_requests.extend([
            {
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1},
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": {"red": 0.10, "green": 0.25, "blue": 0.45},
                        "textFormat": {
                            "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                            "bold": True,
                        },
                    }},
                    "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
                }
            },
            {
                "updateSheetProperties": {
                    "properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 1}},
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            {
                "setBasicFilter": {
                    "filter": {"range": {
                        "sheetId": sheet_id,
                        "startRowIndex": 0,
                        "endRowIndex": max(len(values), 1),
                        "startColumnIndex": 0,
                        "endColumnIndex": column_count,
                    }}
                }
            },
            {
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": sheet_id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": column_count,
                    }
                }
            },
        ])

    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "USER_ENTERED", "data": value_data},
    ).execute()

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": formatting_requests},
    ).execute()

    if GOOGLE_REPORT_FOLDER_ID:
        metadata = drive_service.files().get(
            fileId=spreadsheet_id,
            fields="parents",
            supportsAllDrives=True,
        ).execute()
        previous_parents = ",".join(metadata.get("parents", []))
        update_kwargs = {
            "fileId": spreadsheet_id,
            "addParents": GOOGLE_REPORT_FOLDER_ID,
            "fields": "id,parents",
            "supportsAllDrives": True,
        }
        if previous_parents:
            update_kwargs["removeParents"] = previous_parents
        drive_service.files().update(**update_kwargs).execute()

    if GOOGLE_REPORT_SHARE_EMAIL:
        drive_service.permissions().create(
            fileId=spreadsheet_id,
            body={
                "type": "user",
                "role": "writer",
                "emailAddress": GOOGLE_REPORT_SHARE_EMAIL,
            },
            sendNotificationEmail=False,
            supportsAllDrives=True,
        ).execute()

    return {
        "spreadsheet_id": spreadsheet_id,
        "spreadsheet_url": spreadsheet_url,
        "title": report_data["title"],
    }


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
    starting_balance = db.Column(db.Float, default=0.0, nullable=False)
    email = db.Column(db.String(255))  # for notifications
    # Optional display name for staff
    staff_name = db.Column(db.String(150))

    # Employment status
    is_active = db.Column(db.Boolean, default=True, nullable=False)
    hire_date = db.Column(db.Date)
    inactive_at = db.Column(db.DateTime)
    inactive_reason = db.Column(db.String(255))

    @property
    def is_admin(self):
        return (self.role or "").lower() == "admin"

class LeaveRequest(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    kind = db.Column(db.String(20), default="annual", nullable=False)    # annual/sick
    mode = db.Column(db.String(10), default=RequestMode.hourly, nullable=False)  # hourly/daily
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



class AbsenceNotice(db.Model):
    __tablename__ = "absence_notice"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    start_time = db.Column(db.String(5))
    end_time = db.Column(db.String(5))
    reason = db.Column(db.String(500))
    office_note = db.Column(db.String(500))
    status = db.Column(db.String(20), default="Open", nullable=False)
    leave_request_id = db.Column(
        db.Integer,
        db.ForeignKey("leave_request.id"),
        nullable=True,
    )
    created_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    reminder_sent_at = db.Column(db.DateTime)
    resolved_at = db.Column(db.DateTime)

    user = db.relationship(
        "User",
        foreign_keys=[user_id],
        backref="absence_notices",
        lazy="joined",
    )
    created_by = db.relationship("User", foreign_keys=[created_by_id])
    leave_request = db.relationship("LeaveRequest", foreign_keys=[leave_request_id])


class SchoolYear(db.Model):
    __tablename__ = "school_year"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(20), unique=True, nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    is_active = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class SchoolYearBalance(db.Model):
    __tablename__ = "school_year_balance"
    id = db.Column(db.Integer, primary_key=True)
    school_year_id = db.Column(db.Integer, db.ForeignKey("school_year.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    beginning_balance = db.Column(db.Float, nullable=False, default=0.0)
    note = db.Column(db.String(500))
    updated_at = db.Column(
        db.DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
    )

    school_year = db.relationship("SchoolYear", backref="employee_balances")
    user = db.relationship("User", backref="school_year_balances")

    __table_args__ = (
        db.UniqueConstraint(
            "school_year_id",
            "user_id",
            name="uq_school_year_balance_year_user",
        ),
    )


class LeaveLedger(db.Model):
    __tablename__ = "leave_ledger"
    id = db.Column(db.Integer, primary_key=True)
    school_year_id = db.Column(db.Integer, db.ForeignKey("school_year.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    entry_type = db.Column(db.String(50), nullable=False)
    hours = db.Column(db.Float, nullable=False, default=0.0)
    description = db.Column(db.String(500))
    created_by_id = db.Column(db.Integer, db.ForeignKey("user.id"))
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    school_year = db.relationship("SchoolYear", backref="ledger_entries")
    user = db.relationship("User", foreign_keys=[user_id], backref="leave_ledger_entries")
    created_by = db.relationship("User", foreign_keys=[created_by_id])

    __table_args__ = (
        db.UniqueConstraint(
            "school_year_id",
            "user_id",
            "entry_type",
            name="uq_leave_ledger_year_user_type",
        ),
    )

class GoogleSheetReport(db.Model):
    __tablename__ = "google_sheet_report"

    id = db.Column(db.Integer, primary_key=True)
    report_year = db.Column(db.Integer, nullable=False)
    report_month = db.Column(db.Integer, nullable=False)
    spreadsheet_id = db.Column(db.String(255), nullable=False)
    spreadsheet_url = db.Column(db.String(500), nullable=False)
    title = db.Column(db.String(255), nullable=False)
    generated_by_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    generated_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)

    generated_by = db.relationship("User", foreign_keys=[generated_by_id])


app.jinja_env.globals["User"] = User

@login_manager.user_loader
def load_user(user_id):
    return User.query.get(int(user_id))

@app.context_processor
def inject_models():
    # Makes the User model available inside Jinja templates, e.g. {{ User.query.get(id) }}
    return dict(User=User)

# =========================================================
# Helpers
# =========================================================
def normalize_hours(value):
    """Normalize hours to 2 decimal places to avoid float drift."""
    return round(float(value or 0.0), 2)

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
    """Create tables and add compatibility columns without deleting data."""
    db.create_all()

    try:
        if db.engine.dialect.name == "sqlite":
            sqlite_columns = [
                ("leave_request", "start_time", "VARCHAR(5)"),
                ("leave_request", "end_time", "VARCHAR(5)"),
                (
                    "leave_request",
                    "is_school_related",
                    "BOOLEAN DEFAULT 0 NOT NULL",
                ),
                ("leave_request", "substitute", "VARCHAR(120)"),
                ("user", "staff_name", "VARCHAR(150)"),
                (
                    "user",
                    "starting_balance",
                    "FLOAT DEFAULT 0 NOT NULL",
                ),
                (
                    "user",
                    "is_active",
                    "BOOLEAN DEFAULT 1 NOT NULL",
                ),
                ("user", "hire_date", "DATE"),
                ("user", "inactive_at", "DATETIME"),
                ("user", "inactive_reason", "VARCHAR(255)"),
            ]

            for table_name, column_name, definition in sqlite_columns:
                if not _column_exists(table_name, column_name):
                    db.session.execute(
                        text(
                            f"ALTER TABLE {table_name} "
                            f"ADD COLUMN {column_name} {definition}"
                        )
                    )
        else:
            postgres_statements = [
                (
                    "ALTER TABLE leave_request "
                    "ADD COLUMN IF NOT EXISTS start_time VARCHAR(5)"
                ),
                (
                    "ALTER TABLE leave_request "
                    "ADD COLUMN IF NOT EXISTS end_time VARCHAR(5)"
                ),
                (
                    "ALTER TABLE leave_request "
                    "ADD COLUMN IF NOT EXISTS is_school_related "
                    "BOOLEAN NOT NULL DEFAULT FALSE"
                ),
                (
                    "ALTER TABLE leave_request "
                    "ADD COLUMN IF NOT EXISTS substitute VARCHAR(120)"
                ),
                (
                    'ALTER TABLE "user" '
                    "ADD COLUMN IF NOT EXISTS staff_name VARCHAR(150)"
                ),
                (
                    'ALTER TABLE "user" '
                    "ADD COLUMN IF NOT EXISTS starting_balance "
                    "DOUBLE PRECISION NOT NULL DEFAULT 0"
                ),
                (
                    'ALTER TABLE "user" '
                    "ADD COLUMN IF NOT EXISTS is_active "
                    "BOOLEAN NOT NULL DEFAULT TRUE"
                ),
                (
                    'ALTER TABLE "user" '
                    "ADD COLUMN IF NOT EXISTS hire_date DATE"
                ),
                (
                    'ALTER TABLE "user" '
                    "ADD COLUMN IF NOT EXISTS inactive_at TIMESTAMP"
                ),
                (
                    'ALTER TABLE "user" '
                    "ADD COLUMN IF NOT EXISTS inactive_reason VARCHAR(255)"
                ),
            ]

            for statement in postgres_statements:
                db.session.execute(text(statement))

        db.session.commit()
    except Exception:
        db.session.rollback()
        app.logger.exception("Database compatibility update failed")

    if User.query.count() == 0:
        bootstrap_username = os.environ.get(
            "BOOTSTRAP_ADMIN_USERNAME",
            "mc-admin",
        )
        bootstrap_password = os.environ.get(
            "BOOTSTRAP_ADMIN_PASSWORD",
            "RWAadmin2",
        )
        bootstrap_email = os.environ.get(
            "BOOTSTRAP_ADMIN_EMAIL",
            ADMIN_EMAILS_ENV[0] if ADMIN_EMAILS_ENV else "",
        )

        db.session.add(
            User(
                username=bootstrap_username,
                password_hash=generate_password_hash(
                    bootstrap_password
                ),
                role=Role.admin,
                hours_balance=160.0,
                email=bootstrap_email or None,
                is_active=True,
            )
        )
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

def _manual_adjust_totals_for(user_ids):
    """Return {user_id: total_hours_adjusted} for the given list of ids."""
    if not user_ids:
        return {}
    rows = (
        db.session.query(
            ManualAdjustment.user_id,
            func.coalesce(func.sum(ManualAdjustment.hours), 0.0),
        )
        .filter(ManualAdjustment.user_id.in_(user_ids))
        .group_by(ManualAdjustment.user_id)
        .all()
    )
    return {uid: float(total or 0.0) for uid, total in rows}

def normalize_hours(h: float) -> float:
    """
    Normalize hours to quarter-hour precision and avoid -0.00.
    """
    try:
        h = float(h or 0.0)
    except Exception:
        h = 0.0

    h = round(h * 4) / 4.0

    if abs(h) < 0.0001:
        h = 0.0
    return h


def manual_adjust_sum(user_id: int) -> float:
    total = (
        db.session.query(func.coalesce(func.sum(ManualAdjustment.hours), 0.0))
        .filter(ManualAdjustment.user_id == user_id)
        .scalar()
    )
    return float(total or 0.0)


def approved_leave_sum(user_id: int) -> float:
    total = (
        db.session.query(func.coalesce(func.sum(LeaveRequest.hours), 0.0))
        .filter(
            LeaveRequest.user_id == user_id,
            LeaveRequest.status == RequestStatus.approved,
            LeaveRequest.is_school_related == False,  # noqa
        )
        .scalar()
    )
    return float(total or 0.0)


def expected_balance_for_user(u: User) -> float:
    start = float(getattr(u, "starting_balance", 0.0) or 0.0)
    manual = manual_adjust_sum(u.id)
    taken = approved_leave_sum(u.id)

    expected = start + manual - taken
    return normalize_hours(expected)

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

        if user and not getattr(user, "is_active", True):
            flash(
                "Your account is inactive. Please contact an administrator.",
                "warning",
            )
            return render_template("login.html", title="Login")

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

@app.route("/dashboard")
@login_required
def dashboard():
    me = current_user
    today = date.today()
    active_year = get_active_school_year()

    request_query = LeaveRequest.query.filter(
        LeaveRequest.user_id == me.id
    )
    adjustment_query = ManualAdjustment.query.filter(
        ManualAdjustment.user_id == me.id
    )
    absence_query = AbsenceNotice.query.filter(
        AbsenceNotice.user_id == me.id
    )

    if active_year:
        request_query = request_query.filter(
            LeaveRequest.start_date <= active_year.end_date,
            LeaveRequest.end_date >= active_year.start_date,
        )
        adjustment_query = adjustment_query.filter(
            ManualAdjustment.timestamp
            >= datetime.combine(
                active_year.start_date,
                datetime.min.time(),
            ),
            ManualAdjustment.timestamp
            < datetime.combine(
                active_year.end_date + timedelta(days=1),
                datetime.min.time(),
            ),
        )
        absence_query = absence_query.filter(
            AbsenceNotice.start_date <= active_year.end_date,
            AbsenceNotice.end_date >= active_year.start_date,
        )

    recent = (
        request_query
        .order_by(LeaveRequest.created_at.desc())
        .limit(8)
        .all()
    )

    my_adjustments = (
        adjustment_query
        .order_by(ManualAdjustment.timestamp.desc())
        .limit(6)
        .all()
    )

    pending_count = (
        request_query
        .filter(LeaveRequest.status == RequestStatus.pending)
        .count()
    )

    upcoming_query = request_query.filter(
        LeaveRequest.status == RequestStatus.approved,
        LeaveRequest.end_date >= today,
    )

    upcoming = (
        upcoming_query
        .order_by(LeaveRequest.start_date.asc())
        .limit(5)
        .all()
    )

    school_related_query = db.session.query(
        func.coalesce(func.sum(LeaveRequest.hours), 0.0)
    ).filter(
        LeaveRequest.user_id == me.id,
        LeaveRequest.status == RequestStatus.approved,
        LeaveRequest.is_school_related.is_(True),
    )

    if active_year:
        school_related_query = school_related_query.filter(
            LeaveRequest.start_date <= active_year.end_date,
            LeaveRequest.end_date >= active_year.start_date,
        )

    school_related_total = school_related_query.scalar()

    open_absence_notices = (
        absence_query
        .filter(AbsenceNotice.status == "Open")
        .order_by(AbsenceNotice.start_date.asc())
        .all()
    )

    beginning_balance = 0.0
    recent_ledger = []

    if active_year:
        year_balance = SchoolYearBalance.query.filter_by(
            school_year_id=active_year.id,
            user_id=me.id,
        ).first()

        if year_balance:
            beginning_balance = normalize_hours(
                year_balance.beginning_balance or 0.0
            )

        recent_ledger = (
            LeaveLedger.query
            .filter_by(
                school_year_id=active_year.id,
                user_id=me.id,
            )
            .order_by(
                LeaveLedger.created_at.desc(),
                LeaveLedger.id.desc(),
            )
            .limit(8)
            .all()
        )

    previous_years = (
        SchoolYear.query
        .filter(SchoolYear.is_active.is_(False))
        .order_by(SchoolYear.start_date.desc())
        .all()
    )

    return render_template(
        "dashboard.html",
        me=me,
        recent=recent,
        my_adjustments=my_adjustments,
        pending_count=pending_count,
        upcoming=upcoming,
        school_related_total=normalize_hours(
            school_related_total or 0.0
        ),
        open_absence_notices=open_absence_notices,
        active_year=active_year,
        beginning_balance=beginning_balance,
        recent_ledger=recent_ledger,
        previous_years=previous_years,
    )

# ---------- Admin HUB ----------
@app.get("/admin")
@login_required
def admin_hub():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    today = date.today()
    month_start = date(today.year, today.month, 1)

    if today.month == 12:
        next_month = date(today.year + 1, 1, 1)
    else:
        next_month = date(today.year, today.month + 1, 1)

    pending = (
        LeaveRequest.query
        .filter_by(status=RequestStatus.pending)
        .order_by(LeaveRequest.created_at.asc())
        .all()
    )

    open_absence_count = (
        AbsenceNotice.query
        .filter_by(status="Open")
        .count()
    )

    approved_active = (
        LeaveRequest.query
        .filter(
            LeaveRequest.status == RequestStatus.approved,
            LeaveRequest.end_date >= today,
        )
        .all()
    )

    coverage_needed_count = 0
    for leave_request in approved_active:
        assigned_hours = sum(
            float(sub.hours or 0.0)
            for sub in leave_request.subs
        )
        if assigned_hours < float(leave_request.hours or 0.0):
            coverage_needed_count += 1

    employees_out_today = (
        LeaveRequest.query
        .filter(
            LeaveRequest.status == RequestStatus.approved,
            LeaveRequest.start_date <= today,
            LeaveRequest.end_date >= today,
        )
        .order_by(LeaveRequest.start_date.asc())
        .all()
    )

    substitute_hours_month = (
        db.session.query(func.coalesce(func.sum(SubAssignment.hours), 0.0))
        .join(LeaveRequest, SubAssignment.request_id == LeaveRequest.id)
        .filter(
            LeaveRequest.start_date >= month_start,
            LeaveRequest.start_date < next_month,
        )
        .scalar()
    )

    school_related_hours_month = (
        db.session.query(func.coalesce(func.sum(LeaveRequest.hours), 0.0))
        .filter(
            LeaveRequest.status == RequestStatus.approved,
            LeaveRequest.is_school_related.is_(True),
            LeaveRequest.start_date >= month_start,
            LeaveRequest.start_date < next_month,
        )
        .scalar()
    )

    negative_balance_count = (
        User.query
        .filter(
            User.role == Role.staff,
            User.hours_balance < 0,
        )
        .count()
    )

    active_year = get_active_school_year()

    recent_activity = (
        LeaveLedger.query
        .order_by(
            LeaveLedger.created_at.desc(),
            LeaveLedger.id.desc(),
        )
        .limit(8)
        .all()
    )

    return render_template(
        "admin.html",
        title="Admin",
        pending=pending,
        open_absence_count=open_absence_count,
        coverage_needed_count=coverage_needed_count,
        employees_out_today=employees_out_today,
        substitute_hours_month=normalize_hours(
            substitute_hours_month or 0.0
        ),
        school_related_hours_month=normalize_hours(
            school_related_hours_month or 0.0
        ),
        negative_balance_count=negative_balance_count,
        active_year=active_year,
        recent_activity=recent_activity,
        today=today,
    )





# ---------- Absence Notices ----------
@app.route("/admin/absence-notices", methods=["GET", "POST"])
@login_required
def absence_notices():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        user_id = request.form.get("user_id", type=int)
        start_text = (request.form.get("start_date") or "").strip()
        end_text = (request.form.get("end_date") or "").strip()
        start_time = (request.form.get("start_time") or "").strip() or None
        end_time = (request.form.get("end_time") or "").strip() or None
        reason = (request.form.get("reason") or "").strip()
        office_note = (request.form.get("office_note") or "").strip()
        send_now = request.form.get("send_email") == "yes"

        employee = db.session.get(User, user_id) if user_id else None
        if employee is None:
            flash("Please select an employee.", "warning")
            return redirect(url_for("absence_notices"))

        try:
            start_date = datetime.strptime(start_text, "%Y-%m-%d").date()
            end_date = datetime.strptime(end_text, "%Y-%m-%d").date()
        except ValueError:
            flash("Please enter valid absence dates.", "warning")
            return redirect(url_for("absence_notices"))

        if end_date < start_date:
            flash("The end date cannot be before the start date.", "warning")
            return redirect(url_for("absence_notices"))

        if bool(start_time) != bool(end_time):
            flash("Enter both a start time and an end time, or leave both blank.", "warning")
            return redirect(url_for("absence_notices"))

        notice = AbsenceNotice(
            user_id=employee.id,
            start_date=start_date,
            end_date=end_date,
            start_time=start_time,
            end_time=end_time,
            reason=reason or None,
            office_note=office_note or None,
            status="Open",
            created_by_id=current_user.id,
        )
        db.session.add(notice)
        db.session.commit()

        if send_now:
            sent, message = send_absence_notice_reminder(notice)
            if sent:
                notice.reminder_sent_at = datetime.utcnow()
                db.session.commit()
                flash("The absence notice was saved and the employee was emailed.", "success")
            else:
                flash(
                    f"The absence notice was saved, but the email was not sent: {message}",
                    "warning",
                )
        else:
            flash("The absence notice was saved.", "success")

        return redirect(url_for("absence_notices"))

    status_filter = (request.args.get("status") or "open").strip().lower()
    query = AbsenceNotice.query

    if status_filter == "resolved":
        query = query.filter(AbsenceNotice.status == "Resolved")
    elif status_filter == "all":
        pass
    else:
        status_filter = "open"
        query = query.filter(AbsenceNotice.status == "Open")

    notices = query.order_by(
        AbsenceNotice.start_date.asc(),
        AbsenceNotice.created_at.desc(),
    ).all()

    users = (
        User.query
        .filter(User.is_active.is_(True))
        .order_by(
            func.coalesce(User.staff_name, User.username)
        )
        .all()
    )

    return render_template(
        "absence_notices.html",
        title="Absence Notices",
        notices=notices,
        users=users,
        status_filter=status_filter,
        today=date.today(),
    )


def send_absence_notice_reminder(notice):
    employee = notice.user
    if not employee.email:
        return False, "The employee does not have an email address in the system."

    date_text = notice.start_date.strftime("%B %d, %Y")
    if notice.end_date != notice.start_date:
        date_text += f" through {notice.end_date.strftime('%B %d, %Y')}"

    time_text = ""
    if notice.start_time and notice.end_time:
        time_text = (
            f"\nTime: {h12_filter(notice.start_time)} "
            f"to {h12_filter(notice.end_time)}"
        )

    reason_text = f"\nReason on file: {notice.reason}" if notice.reason else ""

    body = (
        f"Hello {employee.staff_name or employee.username},\n\n"
        "The school office has recorded that you will be absent on the "
        f"following date(s):\n\n{date_text}{time_text}{reason_text}\n\n"
        "Please log in to the Richard Winn Academy Leave System and submit "
        "the matching leave request when you are able.\n\n"
        "Thank you,\nRichard Winn Academy"
    )

    return send_email(
        [employee.email],
        "Reminder: Please Submit Your Leave Request",
        body,
    )


@app.post("/admin/absence-notices/<int:notice_id>/remind")
@login_required
def remind_absence_notice(notice_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    notice = db.session.get(AbsenceNotice, notice_id)
    if notice is None:
        abort(404)

    sent, message = send_absence_notice_reminder(notice)
    if sent:
        notice.reminder_sent_at = datetime.utcnow()
        db.session.commit()
        flash("Reminder email sent.", "success")
    else:
        flash(f"Reminder email was not sent: {message}", "warning")

    return redirect(url_for("absence_notices"))


@app.post("/admin/absence-notices/<int:notice_id>/resolve")
@login_required
def resolve_absence_notice(notice_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    notice = db.session.get(AbsenceNotice, notice_id)
    if notice is None:
        abort(404)

    notice.status = "Resolved"
    notice.resolved_at = datetime.utcnow()
    db.session.commit()
    flash("The absence notice was marked resolved.", "success")
    return redirect(url_for("absence_notices"))


@app.post("/admin/absence-notices/<int:notice_id>/reopen")
@login_required
def reopen_absence_notice(notice_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    notice = db.session.get(AbsenceNotice, notice_id)
    if notice is None:
        abort(404)

    notice.status = "Open"
    notice.resolved_at = None
    db.session.commit()
    flash("The absence notice was reopened.", "success")
    return redirect(url_for("absence_notices", status="all"))


@app.post("/admin/absence-notices/<int:notice_id>/delete")
@login_required
def delete_absence_notice(notice_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    notice = db.session.get(AbsenceNotice, notice_id)
    if notice is None:
        abort(404)

    db.session.delete(notice)
    db.session.commit()
    flash("The absence notice was deleted.", "success")
    return redirect(url_for("absence_notices"))


# ---------- Coverage Center ----------
@app.get("/admin/coverage")
@login_required
def coverage_center():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    start_text = (request.args.get("start") or "").strip()
    end_text = (request.args.get("end") or "").strip()
    status_filter = (request.args.get("status") or "active").strip().lower()

    today = date.today()
    default_end = today + timedelta(days=45)

    try:
        start_date = (
            datetime.strptime(start_text, "%Y-%m-%d").date()
            if start_text else today
        )
        end_date = (
            datetime.strptime(end_text, "%Y-%m-%d").date()
            if end_text else default_end
        )
    except ValueError:
        flash("Please enter valid coverage dates.", "warning")
        start_date = today
        end_date = default_end

    if end_date < start_date:
        start_date, end_date = end_date, start_date

    query = LeaveRequest.query.filter(
        LeaveRequest.end_date >= start_date,
        LeaveRequest.start_date <= end_date,
    )

    if status_filter == "pending":
        query = query.filter(LeaveRequest.status == RequestStatus.pending)
    elif status_filter == "approved":
        query = query.filter(LeaveRequest.status == RequestStatus.approved)
    elif status_filter == "all":
        query = query.filter(
            LeaveRequest.status.in_([
                RequestStatus.pending,
                RequestStatus.approved,
            ])
        )
    else:
        status_filter = "active"
        query = query.filter(
            LeaveRequest.status.in_([
                RequestStatus.pending,
                RequestStatus.approved,
            ])
        )

    requests_list = query.order_by(
        LeaveRequest.start_date.asc(),
        LeaveRequest.created_at.asc(),
    ).all()

    uncovered_count = 0
    covered_count = 0
    total_sub_hours = 0.0

    for leave_request in requests_list:
        assigned_hours = sum((sub.hours or 0.0) for sub in leave_request.subs)
        leave_request.coverage_hours = normalize_hours(assigned_hours)
        leave_request.coverage_complete = (
            assigned_hours >= (leave_request.hours or 0.0)
            and (leave_request.hours or 0.0) > 0
        )
        leave_request.coverage_remaining = normalize_hours(
            max((leave_request.hours or 0.0) - assigned_hours, 0.0)
        )

        if leave_request.coverage_complete:
            covered_count += 1
        else:
            uncovered_count += 1

        total_sub_hours += assigned_hours

    return render_template(
        "coverage_center.html",
        title="Coverage Center",
        requests_list=requests_list,
        start_date=start_date,
        end_date=end_date,
        status_filter=status_filter,
        uncovered_count=uncovered_count,
        covered_count=covered_count,
        total_sub_hours=normalize_hours(total_sub_hours),
    )


# ---------- Leave Ledger Helpers ----------
def get_active_school_year():
    return (
        SchoolYear.query.filter_by(is_active=True)
        .order_by(SchoolYear.start_date.desc())
        .first()
    )


def record_leave_ledger_entry(
    *,
    user_id,
    entry_type,
    hours,
    description,
    created_by_id=None,
    created_at=None,
):
    """Create or update one uniquely identified ledger transaction."""
    school_year = get_active_school_year()
    if school_year is None:
        return None

    entry = LeaveLedger.query.filter_by(
        school_year_id=school_year.id,
        user_id=user_id,
        entry_type=entry_type,
    ).first()

    if entry is None:
        entry = LeaveLedger(
            school_year_id=school_year.id,
            user_id=user_id,
            entry_type=entry_type,
        )
        db.session.add(entry)

    entry.hours = normalize_hours(hours or 0.0)
    entry.description = description
    entry.created_by_id = created_by_id
    entry.created_at = created_at or datetime.utcnow()
    return entry


def ledger_entry_label(entry_type):
    if entry_type == "beginning_balance":
        return "Beginning Balance"
    if entry_type.startswith("approved_request_"):
        return "Approved Leave"
    if entry_type.startswith("cancelled_request_"):
        return "Cancelled Leave Restored"
    if entry_type.startswith("manual_adjustment_deleted_"):
        return "Adjustment Reversed"
    if entry_type.startswith("manual_adjustment_restored_"):
        return "Adjustment Restored"
    if entry_type.startswith("manual_adjustment_"):
        return "Manual Adjustment"
    return "Ledger Entry"


app.jinja_env.globals["ledger_entry_label"] = ledger_entry_label


# ---------- Leave Ledger ----------
@app.route("/admin/leave-ledger")
@login_required
def leave_ledger():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    years = SchoolYear.query.order_by(SchoolYear.start_date.desc()).all()
    selected_year_id = request.args.get("school_year_id", type=int)
    selected_user_id = request.args.get("user_id", type=int)

    selected_year = None
    if selected_year_id:
        selected_year = SchoolYear.query.get(selected_year_id)
    if selected_year is None:
        selected_year = get_active_school_year()
    if selected_year is None and years:
        selected_year = years[0]

    users = (
        User.query
        .filter(User.is_active.is_(True))
        .order_by(
            func.coalesce(User.staff_name, User.username)
        )
        .all()
    )

    query = LeaveLedger.query
    if selected_year is not None:
        query = query.filter(LeaveLedger.school_year_id == selected_year.id)
    if selected_user_id:
        query = query.filter(LeaveLedger.user_id == selected_user_id)

    entries = query.order_by(
        LeaveLedger.created_at.desc(),
        LeaveLedger.id.desc(),
    ).all()

    total_change = normalize_hours(sum((entry.hours or 0.0) for entry in entries))
    selected_user = User.query.get(selected_user_id) if selected_user_id else None

    return render_template(
        "leave_ledger.html",
        title="Leave Ledger",
        years=years,
        users=users,
        selected_year=selected_year,
        selected_user=selected_user,
        selected_user_id=selected_user_id,
        entries=entries,
        total_change=total_change,
    )


# ---------- School Year Setup ----------
@app.route("/admin/school-year", methods=["GET", "POST"])
@login_required
def school_year_setup():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        year_name = (request.form.get("year_name") or "").strip()

        try:
            start_date = datetime.strptime(
                request.form.get("start_date", ""), "%Y-%m-%d"
            ).date()
            end_date = datetime.strptime(
                request.form.get("end_date", ""), "%Y-%m-%d"
            ).date()
        except ValueError:
            flash("Please enter valid school-year dates.", "warning")
            return redirect(url_for("school_year_setup"))

        if not year_name:
            flash("Please enter a school-year name.", "warning")
            return redirect(url_for("school_year_setup"))

        if end_date < start_date:
            flash("The school-year end date must be after the start date.", "warning")
            return redirect(url_for("school_year_setup"))

        try:
            SchoolYear.query.update({SchoolYear.is_active: False})

            school_year = SchoolYear.query.filter_by(name=year_name).first()
            if school_year is None:
                school_year = SchoolYear(
                    name=year_name,
                    start_date=start_date,
                    end_date=end_date,
                    is_active=True,
                )
                db.session.add(school_year)
                db.session.flush()
            else:
                school_year.start_date = start_date
                school_year.end_date = end_date
                school_year.is_active = True
                db.session.flush()

            users = User.query.order_by(
                func.coalesce(User.staff_name, User.username)
            ).all()

            updated_count = 0

            for employee in users:
                raw_balance = (request.form.get(f"balance_{employee.id}") or "").strip()
                if raw_balance == "":
                    continue

                try:
                    beginning_balance = normalize_hours(raw_balance)
                except (TypeError, ValueError):
                    raise ValueError(
                        f"Invalid beginning balance for "
                        f"{employee.staff_name or employee.username}."
                    )

                note = (request.form.get(f"note_{employee.id}") or "").strip()

                saved_balance = SchoolYearBalance.query.filter_by(
                    school_year_id=school_year.id,
                    user_id=employee.id,
                ).first()

                if saved_balance is None:
                    saved_balance = SchoolYearBalance(
                        school_year_id=school_year.id,
                        user_id=employee.id,
                    )
                    db.session.add(saved_balance)

                saved_balance.beginning_balance = beginning_balance
                saved_balance.note = note or None
                saved_balance.updated_at = datetime.utcnow()

                employee.starting_balance = beginning_balance
                employee.hours_balance = beginning_balance

                ledger_entry = LeaveLedger.query.filter_by(
                    school_year_id=school_year.id,
                    user_id=employee.id,
                    entry_type="beginning_balance",
                ).first()

                description = f"Beginning balance for {school_year.name}"
                if note:
                    description += f" — {note}"

                if ledger_entry is None:
                    ledger_entry = LeaveLedger(
                        school_year_id=school_year.id,
                        user_id=employee.id,
                        entry_type="beginning_balance",
                        created_by_id=current_user.id,
                    )
                    db.session.add(ledger_entry)

                ledger_entry.hours = beginning_balance
                ledger_entry.description = description
                ledger_entry.created_by_id = current_user.id
                ledger_entry.created_at = datetime.utcnow()

                updated_count += 1

            db.session.commit()
            flash(
                f"{school_year.name} was saved. "
                f"{updated_count} beginning balance(s) were updated.",
                "success",
            )
            return redirect(url_for("school_year_setup"))

        except ValueError as exc:
            db.session.rollback()
            flash(str(exc), "warning")
        except Exception as exc:
            db.session.rollback()
            app.logger.exception("School-year setup failed")
            flash(f"School-year setup could not be saved: {exc}", "danger")

    active_year = (
        SchoolYear.query.filter_by(is_active=True)
        .order_by(SchoolYear.start_date.desc())
        .first()
    )

    users = (
        User.query
        .filter(User.is_active.is_(True))
        .order_by(
            func.coalesce(User.staff_name, User.username)
        )
        .all()
    )

    balances = {}
    if active_year:
        balances = {
            row.user_id: row
            for row in SchoolYearBalance.query.filter_by(
                school_year_id=active_year.id
            ).all()
        }

    return render_template(
        "school_year_setup.html",
        title="School Year Setup",
        active_year=active_year,
        users=users,
        balances=balances,
    )

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
        if mode not in (RequestMode.hourly, RequestMode.daily):
            flash("Please choose Full Day(s) or Hourly leave.", "warning")
            return render_template("new_request.html", title="New Request", workday=WORKDAY_HOURS)

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

        else:  # Full Day(s) / daily
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
            hours=normalize_hours(hours),
            reason=reason,
            is_school_related=is_school,
        )
        db.session.add(req)
        db.session.commit()

        # Notify admins
        subj = "New Leave Request Submitted"
        body = (
            f"User: {current_user.username}\n"
            f"Kind: {kind}\nMode: {'Full Day(s)' if mode == RequestMode.daily else 'Hourly'}\nHours: {hours:.2f}\n"
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

    # Base query (uses your existing helper)
    q = _filtered_requests_for(is_admin)

    # Admin: optional "day" filter in America/New_York timezone
    selected_day_str = None
    if is_admin:
        tz = ZoneInfo("America/New_York")
        selected_day_str = request.args.get("day") or datetime.now(tz).date().isoformat()

        if selected_day_str != "all":
            try:
                d = datetime.strptime(selected_day_str, "%Y-%m-%d").date()
            except Exception:
                d = datetime.now(tz).date()
                selected_day_str = d.isoformat()

            start_dt = datetime.combine(d, datetime.min.time(), tzinfo=tz)
            end_dt = start_dt + timedelta(days=1)

            # Convert to UTC-naive for DB filter (if DB stores UTC/naive)
            start_utc = start_dt.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)
            end_utc = end_dt.astimezone(ZoneInfo("UTC")).replace(tzinfo=None)

            q = q.filter(
                LeaveRequest.created_at >= start_utc,
                LeaveRequest.created_at < end_utc
            )

    # Final list of requests for this view
    reqs = q.all()

    # -----------------------------
    # Admin overview (Employees at a Glance)
    # -----------------------------
    staff_overview = None
    if is_admin:
        # All users, sorted by username
        users = User.query.order_by(User.username.asc()).all()

        # Who has at least one pending request?
        pending_user_ids = {
            uid for (uid,) in db.session.query(LeaveRequest.user_id)
            .filter(LeaveRequest.status == RequestStatus.pending)
            .distinct()
        }

        # Manual adjustment totals per user
        adjustments_map = _manual_adjust_totals_for([u.id for u in users])

        # Build simple dicts for the template
        staff_overview = []
        for u in users:
            hb = normalize_hours(u.hours_balance or 0.0)
            start_bal = normalize_hours(getattr(u, "starting_balance", 0.0) or 0.0)

            manual_total = normalize_hours(adjustments_map.get(u.id, 0.0))
            expected = expected_balance_for_user(u)
            diff = normalize_hours(hb - expected)

            staff_overview.append({
                "id": u.id,
                "username": u.username,
                "hours_balance": hb,
                "starting_balance": start_bal,
                "adjust_total": manual_total,
                "expected_balance": expected,
                "diff": diff,
                "has_pending": (u.id in pending_user_ids),
            })

  
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
        flash("Admins only.", "warning")
        return redirect(url_for("my_requests"))

    r = LeaveRequest.query.get_or_404(req_id)

    if r.status != RequestStatus.pending:
        flash("Request not pending.", "warning")
        return redirect(url_for("my_requests"))

    u = User.query.get(r.user_id)

    # Deduct hours only if NOT school-related
    if not r.is_school_related:
        u.hours_balance = normalize_hours(
            (u.hours_balance or 0.0) - (r.hours or 0.0)
        )

    # Always approve the request
    r.status = RequestStatus.approved
    r.decided_at = datetime.utcnow()

    ledger_hours = 0.0 if r.is_school_related else -(r.hours or 0.0)
    ledger_description = (
        f"Approved leave request #{r.id}: "
        f"{r.start_date} to {r.end_date}"
    )
    if r.is_school_related:
        ledger_description += " — School related; no leave deducted."

    record_leave_ledger_entry(
        user_id=u.id,
        entry_type=f"approved_request_{r.id}",
        hours=ledger_hours,
        description=ledger_description,
        created_by_id=current_user.id,
        created_at=r.decided_at,
    )

    db.session.commit()

    subs_text = "; ".join(
        [f"{s.name} ({s.hours:.2f}h)" for s in r.subs]
    ) or (r.substitute or "(none)")

    subj = "Leave Request Approved"
    body = (
        f"Hello {u.username},\n\n"
        f"Your leave request has been APPROVED.\n"
        f"Kind: {r.kind}\n"
        f"Mode: {r.mode}\n"
        f"Hours: {r.hours:.2f}\n"
        f"School-related: {'Yes' if r.is_school_related else 'No'}\n"
        f"Substitutes: {subs_text}\n"
        f"Dates: {r.start_date} to {r.end_date}\n\n"
        f"Remaining balance: {u.hours_balance:.2f} hours\n"
    )

    send_email([u.email] + admin_emails(), subj, body)

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

    # Add hours back ONLY if it was approved and not school-related
    if r.status == RequestStatus.approved and not r.is_school_related:
        u.hours_balance = normalize_hours(
            (u.hours_balance or 0.0) + (r.hours or 0.0)
        )

    was_approved_and_deducted = (
        r.status == RequestStatus.approved and not r.is_school_related
    )

    r.status = RequestStatus.cancelled
    r.decided_at = datetime.utcnow()

    if was_approved_and_deducted:
        record_leave_ledger_entry(
            user_id=u.id,
            entry_type=f"cancelled_request_{r.id}",
            hours=(r.hours or 0.0),
            description=(
                f"Hours restored for cancelled leave request #{r.id}: "
                f"{r.start_date} to {r.end_date}"
            ),
            created_by_id=current_user.id,
            created_at=r.decided_at,
        )

    db.session.commit()

    subs_text = "; ".join(
        [f"{s.name} ({s.hours:.2f}h)" for s in r.subs]
    ) or (r.substitute or "(none)")

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

# =========================================================
# Manual Adjustments for Admins — Add / Delete / Undo
# =========================================================
from flask import session

@app.route("/user/<int:user_id>/adjust", methods=["POST"])
@login_required
def add_manual_adjustment_for_user(user_id):
    # Only admins can manually adjust time
    if getattr(current_user, "role", "") != "admin":
        abort(403)

    user = User.query.get_or_404(user_id)
    hours = request.form.get("hours", type=float)
    note = (request.form.get("note") or "").strip()

    if hours is None or note == "":
        flash("Please enter both hours and a note.", "warning")
        return redirect(url_for("user_requests", user_id=user.id))

    # Create the manual adjustment entry
    adj = ManualAdjustment(
        user_id=user.id,
        admin_id=current_user.id,
        hours=hours,
        note=note,
        timestamp=datetime.utcnow(),
    )
    db.session.add(adj)

    # Update the user's balance (normalized)
    user.hours_balance = normalize_hours((user.hours_balance or 0.0) + float(hours or 0.0))
    db.session.flush()

    record_leave_ledger_entry(
        user_id=user.id,
        entry_type=f"manual_adjustment_{adj.id}",
        hours=hours,
        description=f"Manual adjustment: {note}",
        created_by_id=current_user.id,
        created_at=adj.timestamp,
    )

    db.session.commit()

    # Optional email notifications (won't break page if it fails)
    try:
        if user.email:
            subject_emp = "Your leave balance has been adjusted"
            body_emp = (
                f"Hi {user.staff_name or user.username},\n\n"
                f"Your leave balance has been adjusted by {hours:+.2f} hours.\n\n"
                f"Reason: {note}\n"
                f"New balance: {user.hours_balance:.2f} hours\n\n"
                f"This change was made on {adj.timestamp.strftime('%b %d, %Y at %I:%M %p')} "
                f"by {current_user.staff_name or current_user.username}.\n"
            )
            send_email([user.email], subject_emp, body_emp)

        admin_recipients = admin_emails()
        if admin_recipients:
            subject_admin = f"Manual leave adjustment for {user.staff_name or user.username}"
            body_admin = (
                "Manual leave adjustment recorded.\n\n"
                f"Employee: {user.staff_name or user.username} (username: {user.username})\n"
                f"Changed by: {current_user.staff_name or current_user.username}\n\n"
                f"Amount: {hours:+.2f} hours\n"
                f"Reason: {note}\n"
                f"New balance: {user.hours_balance:.2f} hours\n"
                f"When: {adj.timestamp.strftime('%b %d, %Y at %I:%M %p')}\n"
            )
            send_email(admin_recipients, subject_admin, body_admin)

    except Exception as e:
        app.logger.error(f"Error sending manual adjustment emails: {e}")

    flash(f"Manual adjustment of {hours:+.2f}h added for {user.username}.", "success")
    return redirect(url_for("user_requests", user_id=user.id))

# =========================================================
# Delete Manual Adjustment (with Undo support)
# =========================================================
@app.route("/user/<int:user_id>/adjust/<int:adj_id>/delete", methods=["POST"])
@login_required
def delete_adjustment(user_id, adj_id):
    """Admin deletes a manual adjustment with undo option."""
    if getattr(current_user, "role", "") != "admin":
        abort(403)

    adj = ManualAdjustment.query.get_or_404(adj_id)
    user = User.query.get_or_404(user_id)

    # Save deleted record for undo
    session["last_deleted_adjustment"] = {
        "user_id": user.id,
        "adj_id": adj.id,
        "hours": adj.hours,
        "note": adj.note,
        "timestamp": adj.timestamp.isoformat() if adj.timestamp else None,
        "admin_id": adj.admin_id,
    }

    try:
        user.hours_balance = normalize_hours(
        (user.hours_balance or 0.0) - adj.hours
        )
        db.session.delete(adj)
        db.session.commit()
        flash(
            f"Adjustment ({adj.hours:+.2f}h) deleted. "
            f"<a href='{url_for('undo_delete_adjustment')}' class='alert-link'>Undo</a>",
            "warning"
        )
    except Exception as e:
        db.session.rollback()
        flash(f"Error deleting adjustment: {e}", "danger")

    return redirect(url_for("user_requests", user_id=user.id))


# =========================================================
# Undo Restore Adjustment
# =========================================================
@app.route("/undo_delete_adjustment")
@login_required
def undo_delete_adjustment():
    """Restore most recently deleted manual adjustment."""
    if getattr(current_user, "role", "") != "admin":
        abort(403)

    data = session.pop("last_deleted_adjustment", None)
    if not data:
        flash("No recent adjustment to restore.", "info")
        return redirect(url_for("my_requests"))

    try:
        adj = ManualAdjustment(
            user_id=data["user_id"],
            hours=data["hours"],
            note=data["note"],
            admin_id=data.get("admin_id"),
            timestamp=datetime.fromisoformat(data["timestamp"]) if data.get("timestamp") else datetime.now(),
        )
        user = User.query.get(data["user_id"])
        user.hours_balance = normalize_hours(
        (user.hours_balance or 0.0) + adj.hours
        )
        db.session.add(adj)
        db.session.flush()

        record_leave_ledger_entry(
            user_id=user.id,
            entry_type=f"manual_adjustment_restored_{adj.id}",
            hours=adj.hours,
            description=f"Manual adjustment restored: {adj.note}",
            created_by_id=current_user.id,
        )

        db.session.commit()
        flash("Deleted adjustment has been restored.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error restoring adjustment: {e}", "danger")

    return redirect(url_for("user_requests", user_id=data["user_id"]))


# ---------- Employee School-Year Archive ----------
@app.get("/my-history")
@login_required
def my_school_year_history():
    years = (
        SchoolYear.query
        .order_by(SchoolYear.start_date.desc())
        .all()
    )

    selected_year_id = request.args.get("school_year_id", type=int)
    selected_year = None

    if selected_year_id:
        selected_year = SchoolYear.query.get(selected_year_id)

    if selected_year is None:
        selected_year = (
            SchoolYear.query
            .filter(SchoolYear.is_active.is_(False))
            .order_by(SchoolYear.start_date.desc())
            .first()
        )

    requests_list = []
    adjustments = []
    ledger_entries = []
    beginning_balance = 0.0

    if selected_year:
        requests_list = (
            LeaveRequest.query
            .filter(
                LeaveRequest.user_id == current_user.id,
                LeaveRequest.start_date <= selected_year.end_date,
                LeaveRequest.end_date >= selected_year.start_date,
            )
            .order_by(
                LeaveRequest.start_date.desc(),
                LeaveRequest.created_at.desc(),
            )
            .all()
        )

        adjustments = (
            ManualAdjustment.query
            .filter(
                ManualAdjustment.user_id == current_user.id,
                ManualAdjustment.timestamp
                >= datetime.combine(
                    selected_year.start_date,
                    datetime.min.time(),
                ),
                ManualAdjustment.timestamp
                < datetime.combine(
                    selected_year.end_date + timedelta(days=1),
                    datetime.min.time(),
                ),
            )
            .order_by(ManualAdjustment.timestamp.desc())
            .all()
        )

        ledger_entries = (
            LeaveLedger.query
            .filter_by(
                school_year_id=selected_year.id,
                user_id=current_user.id,
            )
            .order_by(
                LeaveLedger.created_at.desc(),
                LeaveLedger.id.desc(),
            )
            .all()
        )

        year_balance = SchoolYearBalance.query.filter_by(
            school_year_id=selected_year.id,
            user_id=current_user.id,
        ).first()

        if year_balance:
            beginning_balance = normalize_hours(
                year_balance.beginning_balance or 0.0
            )

    return render_template(
        "school_year_archive.html",
        title="Previous School Years",
        years=years,
        selected_year=selected_year,
        requests_list=requests_list,
        adjustments=adjustments,
        ledger_entries=ledger_entries,
        beginning_balance=beginning_balance,
        me=current_user,
    )


# ---------- Employee Self-Service Ledger ----------
@app.get("/my-ledger")
@login_required
def my_leave_ledger():
    active_year = get_active_school_year()
    entries = []
    beginning_balance = 0.0

    if active_year:
        year_balance = SchoolYearBalance.query.filter_by(
            school_year_id=active_year.id,
            user_id=current_user.id,
        ).first()

        if year_balance:
            beginning_balance = normalize_hours(
                year_balance.beginning_balance or 0.0
            )

        entries = (
            LeaveLedger.query
            .filter_by(
                school_year_id=active_year.id,
                user_id=current_user.id,
            )
            .order_by(
                LeaveLedger.created_at.desc(),
                LeaveLedger.id.desc(),
            )
            .all()
        )

    return render_template(
        "my_leave_ledger.html",
        title="My Leave Ledger",
        active_year=active_year,
        beginning_balance=beginning_balance,
        entries=entries,
        me=current_user,
    )


# =========================================================
# Show individual user's leave history (admin only)
# =========================================================
@app.route("/user/<int:user_id>/requests")
@login_required
def user_requests(user_id):
    is_admin = current_user.role == Role.admin

    if not is_admin and current_user.id != user_id:
        abort(403)

    user = User.query.get_or_404(user_id)

    reqs = (
        LeaveRequest.query
        .filter_by(user_id=user.id)
        .order_by(
            LeaveRequest.start_date.desc(),
            LeaveRequest.created_at.desc(),
        )
        .all()
    )

    adjustments = (
        ManualAdjustment.query
        .filter_by(user_id=user.id)
        .order_by(ManualAdjustment.timestamp.desc())
        .all()
    )

    active_year = get_active_school_year()
    ledger_entries = []
    beginning_balance = 0.0

    if active_year:
        year_balance = SchoolYearBalance.query.filter_by(
            school_year_id=active_year.id,
            user_id=user.id,
        ).first()

        if year_balance:
            beginning_balance = normalize_hours(
                year_balance.beginning_balance or 0.0
            )

        ledger_entries = (
            LeaveLedger.query
            .filter_by(
                school_year_id=active_year.id,
                user_id=user.id,
            )
            .order_by(
                LeaveLedger.created_at.desc(),
                LeaveLedger.id.desc(),
            )
            .all()
        )

    return render_template(
        "user_requests.html",
        user=user,
        requests=reqs,
        adjustments=adjustments,
        ledger_entries=ledger_entries,
        beginning_balance=beginning_balance,
        active_year=active_year,
        is_admin=is_admin,
        me=current_user,
    )

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
    status_filter = (
        request.args.get("status") or "active"
    ).strip().lower()

    query = User.query

    if qtxt:
        search_text = f"%{qtxt}%"
        query = query.filter(
            db.or_(
                User.username.ilike(search_text),
                User.staff_name.ilike(search_text),
                User.email.ilike(search_text),
            )
        )

    if status_filter == "inactive":
        query = query.filter(User.is_active.is_(False))
    elif status_filter == "all":
        pass
    else:
        status_filter = "active"
        query = query.filter(User.is_active.is_(True))

    users = query.order_by(
        func.coalesce(User.staff_name, User.username)
    ).all()

    active_count = User.query.filter(
        User.is_active.is_(True)
    ).count()

    inactive_count = User.query.filter(
        User.is_active.is_(False)
    ).count()

    admin_count = User.query.filter(
        User.role == Role.admin,
        User.is_active.is_(True),
    ).count()

    return render_template(
        "manage_users.html",
        title="Manage Employees",
        users=users,
        q=qtxt,
        status_filter=status_filter,
        active_count=active_count,
        inactive_count=inactive_count,
        admin_count=admin_count,
    )


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
    hours_str = (
        request.form.get("hours_balance") or ""
    ).strip()
    password = (request.form.get("password") or "").strip()
    hire_date_text = (
        request.form.get("hire_date") or ""
    ).strip()

    if not username or not password:
        flash(
            "Username and password are required.",
            "warning",
        )
        return redirect(url_for("manage_users"))

    if User.query.filter(User.username.ilike(username)).first():
        flash("Username already exists.", "danger")
        return redirect(url_for("manage_users"))

    try:
        hours_balance = normalize_hours(
            float(hours_str) if hours_str else 160.0
        )
    except Exception:
        hours_balance = 160.0

    hire_date = None
    if hire_date_text:
        try:
            hire_date = datetime.strptime(
                hire_date_text,
                "%Y-%m-%d",
            ).date()
        except ValueError:
            flash(
                "The hire date was not valid and was not saved.",
                "warning",
            )

    user = User(
        username=username,
        staff_name=staff_name or None,
        password_hash=generate_password_hash(password),
        role=(
            role
            if role in (Role.admin, Role.staff)
            else Role.staff
        ),
        hours_balance=hours_balance,
        email=email or None,
        hire_date=hire_date,
        is_active=True,
    )

    db.session.add(user)
    db.session.commit()

    flash(
        f"Employee '{staff_name or username}' was created.",
        "success",
    )
    return redirect(url_for("manage_users"))


@app.post("/admin/users/<int:user_id>/update")
@login_required
def admin_update_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))

    user = User.query.get_or_404(user_id)

    staff_name = (
        request.form.get("staff_name") or ""
    ).strip()
    email = (request.form.get("email") or "").strip()
    role = (request.form.get("role") or "").strip()
    balance_text = (
        request.form.get("hours_balance") or ""
    ).strip()
    hire_date_text = (
        request.form.get("hire_date") or ""
    ).strip()

    user.staff_name = staff_name or None
    user.email = email or None

    if role in (Role.admin, Role.staff):
        user.role = role

    try:
        if balance_text != "":
            user.hours_balance = normalize_hours(
                float(balance_text)
            )
    except Exception:
        flash("The balance value was not valid.", "warning")

    if hire_date_text:
        try:
            user.hire_date = datetime.strptime(
                hire_date_text,
                "%Y-%m-%d",
            ).date()
        except ValueError:
            flash("The hire date was not valid.", "warning")
    else:
        user.hire_date = None

    db.session.commit()
    flash(
        f"Updated {user.staff_name or user.username}.",
        "success",
    )
    return redirect(
        url_for(
            "manage_users",
            status=request.args.get("status", "active"),
        )
    )


@app.post("/admin/users/<int:user_id>/reset")
@login_required
def admin_reset_password(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))

    new_password = (
        request.form.get("new_password") or ""
    ).strip()

    if not new_password:
        flash("Password cannot be empty.", "warning")
        return redirect(url_for("manage_users"))

    user = User.query.get_or_404(user_id)
    user.password_hash = generate_password_hash(new_password)
    db.session.commit()

    flash(
        f"Password updated for "
        f"{user.staff_name or user.username}.",
        "success",
    )
    return redirect(url_for("manage_users"))


@app.post("/admin/users/<int:user_id>/deactivate")
@login_required
def admin_deactivate_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))

    user = User.query.get_or_404(user_id)

    if user.id == current_user.id:
        flash(
            "You cannot deactivate your own account.",
            "warning",
        )
        return redirect(url_for("manage_users"))

    active_admin_count = User.query.filter(
        User.role == Role.admin,
        User.is_active.is_(True),
    ).count()

    if (
        user.role == Role.admin
        and active_admin_count <= 1
    ):
        flash(
            "At least one active administrator must remain.",
            "warning",
        )
        return redirect(url_for("manage_users"))

    reason = (
        request.form.get("inactive_reason") or ""
    ).strip()

    user.is_active = False
    user.inactive_at = datetime.utcnow()
    user.inactive_reason = reason or "No reason entered"

    db.session.commit()

    flash(
        f"{user.staff_name or user.username} was deactivated. "
        "All historical records were preserved.",
        "success",
    )
    return redirect(
        url_for("manage_users", status="inactive")
    )


@app.post("/admin/users/<int:user_id>/reactivate")
@login_required
def admin_reactivate_user(user_id):
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("manage_users"))

    user = User.query.get_or_404(user_id)
    user.is_active = True
    user.inactive_at = None
    user.inactive_reason = None
    db.session.commit()

    flash(
        f"{user.staff_name or user.username} was reactivated.",
        "success",
    )
    return redirect(
        url_for("manage_users", status="active")
    )


# Keep the old endpoint safe for bookmarks or older forms.
@app.post("/admin/users/<int:user_id>/delete")
@login_required
def admin_delete_user(user_id):
    flash(
        "Employees are no longer permanently deleted. "
        "Use Deactivate to preserve historical records.",
        "warning",
    )
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

# ---------- Google Sheets Reports ----------
@app.route("/admin/google-sheets", methods=["GET", "POST"])
@login_required
def google_sheets_reports():
    if current_user.role != Role.admin:
        flash("Admins only.", "warning")
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        report_year = request.form.get("report_year", type=int)
        report_month = request.form.get("report_month", type=int)

        if not report_year or not report_month or report_month not in range(1, 13):
            flash("Please select a valid report month and year.", "warning")
            return redirect(url_for("google_sheets_reports"))

        if not google_sheets_configuration_status()["ready"]:
            flash("Google Sheets is disabled or its Render settings are incomplete.", "warning")
            return redirect(url_for("google_sheets_reports"))

        try:
            result = create_google_sheet_report(report_year, report_month)
            report = GoogleSheetReport(
                report_year=report_year,
                report_month=report_month,
                spreadsheet_id=result["spreadsheet_id"],
                spreadsheet_url=result["spreadsheet_url"],
                title=result["title"],
                generated_by_id=current_user.id,
            )
            db.session.add(report)
            db.session.commit()
            flash("Google Sheet report created successfully.", "success")
            return redirect(result["spreadsheet_url"])
        except Exception as exc:
            db.session.rollback()
            app.logger.exception("Google Sheet report generation failed")
            flash(f"Google Sheet report could not be created: {exc}", "danger")
            return redirect(url_for("google_sheets_reports"))

    today = date.today()
    reports = GoogleSheetReport.query.order_by(
        GoogleSheetReport.generated_at.desc()
    ).limit(25).all()

    return render_template(
        "google_sheets_reports.html",
        title="Google Sheets Reports",
        google_status=google_sheets_configuration_status(),
        reports=reports,
        default_year=today.year,
        default_month=today.month,
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

# Dev server entry (ignored by gunicorn)
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
