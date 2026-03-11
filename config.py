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

# Shift windows (hours as float: 15.5 = 3:30 PM). Shift = whichever window contains most of the work time.
# Day: e.g. 7:00 AM - 3:30 PM
DAY_SHIFT_START_HOUR = float(os.environ.get("TIMESHEET_DAY_SHIFT_START", "7"))
DAY_SHIFT_END_HOUR = float(os.environ.get("TIMESHEET_DAY_SHIFT_END", "15.5"))
# Swing: e.g. 3:00 PM - 11:45 PM
SWING_SHIFT_START_HOUR = float(os.environ.get("TIMESHEET_SWING_SHIFT_START", "15"))
SWING_SHIFT_END_HOUR = float(os.environ.get("TIMESHEET_SWING_SHIFT_END", "23.75"))
# Graveyard: 10 PM - 6 AM (next day); most time in this window = graveyard shift
GRAVEYARD_START_HOUR = int(os.environ.get("TIMESHEET_GRAVEYARD_START", "22"))
GRAVEYARD_END_HOUR = int(os.environ.get("TIMESHEET_GRAVEYARD_END", "6"))

# Time-off notification: only Admin → Settings (no default email in code)
TIMEOFF_NOTIFY_EMAIL = os.environ.get("TIMESHEET_TIMEOFF_NOTIFY_EMAIL", "")
# SMTP (optional): set to send time-off emails. Default account committed for convenience; override with env TIMESHEET_SMTP_* or email_config.env.
SMTP_HOST = os.environ.get("TIMESHEET_SMTP_HOST", "smtp.office365.com")
SMTP_PORT = int(os.environ.get("TIMESHEET_SMTP_PORT", "587"))
SMTP_USER = os.environ.get("TIMESHEET_SMTP_USER", "FA.epd2@fii-na.com")
SMTP_PASSWORD = os.environ.get("TIMESHEET_SMTP_PASSWORD", "FA-op-8299")
SMTP_USE_TLS = os.environ.get("TIMESHEET_SMTP_USE_TLS", "1").strip().lower() in ("1", "true", "yes")
SMTP_FROM = os.environ.get("TIMESHEET_SMTP_FROM", "") or os.environ.get("TIMESHEET_SMTP_USER", "") or "FA.epd2@fii-na.com"

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
                    TIMEOFF_NOTIFY_EMAIL = v or ""
                elif k == "TIMESHEET_DAY_SHIFT_START":
                    DAY_SHIFT_START_HOUR = float(v) if v else 7
                elif k == "TIMESHEET_DAY_SHIFT_END":
                    DAY_SHIFT_END_HOUR = float(v) if v else 15.5
                elif k == "TIMESHEET_SWING_SHIFT_START":
                    SWING_SHIFT_START_HOUR = float(v) if v else 15
                elif k == "TIMESHEET_SWING_SHIFT_END":
                    SWING_SHIFT_END_HOUR = float(v) if v else 23.75
                elif k == "TIMESHEET_GRAVEYARD_START":
                    GRAVEYARD_START_HOUR = int(v) if v else 22
                elif k == "TIMESHEET_GRAVEYARD_END":
                    GRAVEYARD_END_HOUR = int(v) if v else 6
    except Exception:
        pass
    if SMTP_USER and (not SMTP_FROM or SMTP_FROM == "timesheet@localhost"):
        SMTP_FROM = SMTP_USER
