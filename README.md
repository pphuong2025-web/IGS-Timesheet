# Timesheet Web Application

A simple timesheet system with employee login, admin management, weekly times (Monday–Sunday), automatic overtime and graveyard detection, and Excel export.

## Features

- **Authentication**: Log in with full employee name and password. Session-based login.
- **Admin**: Add, edit, and delete employees. Only admins can access the Employees page.
- **Timesheet**: Each employee enters clock-in/clock-out per day. Work week is Monday–Sunday. You can view and edit previous weeks.
- **Overtime**: Automatically calculated at over 8 hours per day (after deducting lunch).
- **Graveyard shift**: Any shift that includes work between 10 PM and 6 AM is flagged.
- **Export**: Export any week’s timesheet data to an Excel file (from the timesheet page or as admin for all employees).

## Setup

```bash
cd timesheet
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5050 (or the port shown). On first run, a default admin account is created:

- **Full name:** `admin`  
- **Password:** `admin` (Administrator privileges)  

Change this password after first login (Admin → Employees → Edit admin).

## Usage

1. **Log in** with your full name and password.
2. **Timesheet**: Use “Previous week” / “Next week” to move between weeks. Enter clock-in and clock-out (and optional notes), then click **Save** for that row. Regular/overtime and graveyard update when you reload.
3. **Export to Excel**: On the timesheet page, click **Export to Excel** to download the current week’s data.
4. **Admins**: Go to **Employees** to add, edit, or delete employees. New employees can then log in with the credentials you set.

## Configuration

Edit `config.py` to change:

- `REGULAR_HOURS_PER_DAY` (default 8; hours over this per day = overtime)
- `GRAVEYARD_START_HOUR` / `GRAVEYARD_END_HOUR` (default 22 and 6)
- `SECRET_KEY` (set via env `TIMESHEET_SECRET_KEY` in production)

## Data

- SQLite database: `timesheet/timesheet.db` (created on first run).
- No automatic backup; copy `timesheet.db` to back up.
