"""App configuration."""
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE_PATH = os.path.join(BASE_DIR, "timesheet.db")
SECRET_KEY = os.environ.get("TIMESHEET_SECRET_KEY", "change-me-in-production-use-env-var")

# Master password: if set, Administrator can log in as any employee by entering that
# employee's full name and the master password. Set via env TIMESHEET_MASTER_PASSWORD.
MASTER_PASSWORD = os.environ.get("TIMESHEET_MASTER_PASSWORD", "")

# Work week: Monday = 0, Sunday = 6
WEEK_STARTS_ON = 0  # Monday

# Overtime: hours over this per day (after lunch deduction) count as overtime
REGULAR_HOURS_PER_DAY = 8

# Graveyard shift: work between these hours (24h) is considered graveyard
GRAVEYARD_START_HOUR = 22   # 10 PM
GRAVEYARD_END_HOUR = 6      # 6 AM (next day)

# Time-off notification: email sent when an employee requests time off
TIMEOFF_NOTIFY_EMAIL = os.environ.get("TIMESHEET_TIMEOFF_NOTIFY_EMAIL", "phuong.pham@fii-na.com")
# SMTP (optional): set to send time-off emails. If not set, notification is skipped.
# You can set env vars TIMESHEET_SMTP_* or use a file: create email_config.env in this folder
# with lines like: TIMESHEET_SMTP_HOST=smtp.office365.com
SMTP_HOST = os.environ.get("TIMESHEET_SMTP_HOST", "")
SMTP_PORT = int(os.environ.get("TIMESHEET_SMTP_PORT", "587"))
SMTP_USER = os.environ.get("TIMESHEET_SMTP_USER", "")
SMTP_PASSWORD = os.environ.get("TIMESHEET_SMTP_PASSWORD", "")
SMTP_USE_TLS = os.environ.get("TIMESHEET_SMTP_USE_TLS", "1").strip().lower() in ("1", "true", "yes")
SMTP_FROM = os.environ.get("TIMESHEET_SMTP_FROM", "") or os.environ.get("TIMESHEET_SMTP_USER", "") or "timesheet@localhost"

# Optional: load SMTP from email_config.env (same folder as config.py). One key=value per line. Do not commit this file.
_email_config_path = os.path.join(BASE_DIR, "email_config.env")
if os.path.isfile(_email_config_path):
    try:
        with open(_email_config_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, _, v = line.partition("=")
                k, v = k.strip(), v.strip()
                if k == "TIMESHEET_SMTP_HOST":
                    SMTP_HOST = v
                elif k == "TIMESHEET_SMTP_PORT":
                    SMTP_PORT = int(v) if v else 587
                elif k == "TIMESHEET_SMTP_USER":
                    SMTP_USER = v
                elif k == "TIMESHEET_SMTP_PASSWORD":
                    SMTP_PASSWORD = v
                elif k == "TIMESHEET_SMTP_FROM":
                    SMTP_FROM = v
                elif k == "TIMESHEET_SMTP_USE_TLS":
                    SMTP_USE_TLS = v.strip().lower() in ("1", "true", "yes")
                elif k == "TIMESHEET_TIMEOFF_NOTIFY_EMAIL":
                    TIMEOFF_NOTIFY_EMAIL = v
    except Exception:
        pass
    if SMTP_USER and (not SMTP_FROM or SMTP_FROM == "timesheet@localhost"):
        SMTP_FROM = SMTP_USER
