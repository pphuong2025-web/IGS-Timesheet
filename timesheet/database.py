"""
SQLite schema and helpers for timesheet: employees and time entries.
Work week: Monday–Sunday. Entries store clock-in/out per day.
"""
import sqlite3
from contextlib import contextmanager
from datetime import datetime, date, time, timedelta

import config


def init_db():
    """Create tables if they don't exist."""
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT NOT NULL,
                is_admin INTEGER NOT NULL DEFAULT 0,
                shift TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        conn.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_employees_username ON employees(username)")

        conn.execute("""
            CREATE TABLE IF NOT EXISTS time_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
                work_date TEXT NOT NULL,
                clock_in TEXT,
                clock_out TEXT,
                lunch_start TEXT,
                lunch_end TEXT,
                regular_hours REAL NOT NULL DEFAULT 0,
                overtime_hours REAL NOT NULL DEFAULT 0,
                is_graveyard INTEGER NOT NULL DEFAULT 0,
                notes TEXT,
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                UNIQUE(employee_id, work_date)
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_time_entries_employee ON time_entries(employee_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_time_entries_work_date ON time_entries(work_date)")
        # Migration: add lunch columns if table existed without them
        cur = conn.execute("PRAGMA table_info(time_entries)")
        cols = [row[1] for row in cur.fetchall()]
        if "lunch_start" not in cols:
            conn.execute("ALTER TABLE time_entries ADD COLUMN lunch_start TEXT")
        if "lunch_end" not in cols:
            conn.execute("ALTER TABLE time_entries ADD COLUMN lunch_end TEXT")
        # Migration: add shift column to employees (day, swing, graveyard)
        cur = conn.execute("PRAGMA table_info(employees)")
        emp_cols = [row[1] for row in cur.fetchall()]
        if "shift" not in emp_cols:
            conn.execute("ALTER TABLE employees ADD COLUMN shift TEXT")
        if "employment_type" not in emp_cols:
            conn.execute("ALTER TABLE employees ADD COLUMN employment_type TEXT")
        conn.execute("""
            CREATE TABLE IF NOT EXISTS time_off_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
                from_date TEXT NOT NULL,
                to_date TEXT NOT NULL,
                notes TEXT NOT NULL,
                hours_per_day REAL NOT NULL DEFAULT 8,
                status TEXT NOT NULL DEFAULT 'pending',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_time_off_requests_employee ON time_off_requests(employee_id)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_time_off_requests_status ON time_off_requests(status)")
        conn.commit()


@contextmanager
def _conn():
    conn = sqlite3.connect(config.DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


# --- Employees ---

def create_employee(conn, username, password_hash, full_name, is_admin=False, shift=None, employment_type=None):
    now = datetime.utcnow().isoformat() + "Z"
    shift_val = (shift or "").strip().lower() or None
    if shift_val and shift_val not in ("day", "swing", "graveyard"):
        shift_val = None
    emp_type = (employment_type or "full_time").strip().lower() if employment_type else "full_time"
    if emp_type not in ("full_time", "contractor"):
        emp_type = "full_time"
    conn.execute(
        "INSERT INTO employees (username, password_hash, full_name, is_admin, shift, employment_type, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (username, password_hash, full_name, 1 if is_admin else 0, shift_val, emp_type, now, now),
    )
    conn.commit()
    return conn.execute("SELECT last_insert_rowid()").fetchone()[0]


def get_employee_by_id(employee_id):
    with _conn() as conn:
        row = conn.execute("SELECT * FROM employees WHERE id = ?", (employee_id,)).fetchone()
        return dict(row) if row else None


def get_employee_by_username(username):
    with _conn() as conn:
        row = conn.execute("SELECT * FROM employees WHERE username = ?", (username,)).fetchone()
        return dict(row) if row else None


def get_employee_by_full_name(full_name):
    """Look up employee by full name (exact match). If multiple exist, returns first."""
    with _conn() as conn:
        row = conn.execute("SELECT * FROM employees WHERE full_name = ?", (full_name.strip(),)).fetchone()
        return dict(row) if row else None


def list_employees():
    with _conn() as conn:
        rows = conn.execute(
            "SELECT id, full_name, is_admin, shift, employment_type, created_at, updated_at FROM employees ORDER BY full_name"
        ).fetchall()
        return [dict(r) for r in rows]


def list_employees_by_shift():
    """Return employees grouped by shift: day, swing, graveyard, unassigned."""
    all_emp = list_employees()
    by_shift = {"day": [], "swing": [], "graveyard": [], "unassigned": []}
    for e in all_emp:
        s = (e.get("shift") or "").strip().lower()
        if s in ("day", "swing", "graveyard"):
            by_shift[s].append(e)
        else:
            by_shift["unassigned"].append(e)
    return by_shift


def list_employees_for_shift(shift):
    """Return employees in the given shift (day, swing, graveyard). Empty list if shift invalid."""
    if not shift or (shift or "").strip().lower() not in ("day", "swing", "graveyard"):
        return []
    by_shift = list_employees_by_shift()
    return by_shift.get((shift or "").strip().lower(), [])


def update_employee(conn, employee_id, full_name=None, password_hash=None, is_admin=None, shift=None, employment_type=None):
    updates = ["updated_at = ?"]
    args = [datetime.utcnow().isoformat() + "Z"]
    if full_name is not None:
        updates.append("full_name = ?")
        args.append(full_name)
    if password_hash is not None:
        updates.append("password_hash = ?")
        args.append(password_hash)
    if is_admin is not None:
        updates.append("is_admin = ?")
        args.append(1 if is_admin else 0)
    if shift is not None:
        shift_val = (shift if isinstance(shift, str) else "").strip().lower() or None
        if shift_val and shift_val not in ("day", "swing", "graveyard"):
            shift_val = None
        updates.append("shift = ?")
        args.append(shift_val)
    if employment_type is not None:
        emp_type = (employment_type if isinstance(employment_type, str) else "").strip().lower()
        if emp_type not in ("full_time", "contractor"):
            emp_type = "full_time"
        updates.append("employment_type = ?")
        args.append(emp_type)
    args.append(employee_id)
    conn.execute(
        f"UPDATE employees SET {', '.join(updates)} WHERE id = ?",
        args,
    )
    conn.commit()


def delete_employee(conn, employee_id):
    conn.execute("DELETE FROM employees WHERE id = ?", (employee_id,))
    conn.execute("DELETE FROM time_entries WHERE employee_id = ?", (employee_id,))
    conn.commit()


# --- Time entries ---

def get_week_start(d):
    """Return Monday of the week containing d (date)."""
    if isinstance(d, str):
        d = date.fromisoformat(d)
    # Monday = 0
    weekday = d.weekday()
    return d - timedelta(days=weekday)


def get_week_range(week_start):
    """Return (week_start, week_end) for the given Monday."""
    if isinstance(week_start, str):
        week_start = date.fromisoformat(week_start)
    week_end = week_start + timedelta(days=6)
    return week_start, week_end


def upsert_time_entry(conn, employee_id, work_date, clock_in=None, clock_out=None, lunch_start=None, lunch_end=None, notes=None,
                      regular_hours=0, overtime_hours=0, is_graveyard=0):
    now = datetime.utcnow().isoformat() + "Z"
    if isinstance(work_date, date):
        work_date = work_date.isoformat()
    conn.execute("""
        INSERT INTO time_entries (employee_id, work_date, clock_in, clock_out, lunch_start, lunch_end, regular_hours, overtime_hours, is_graveyard, notes, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(employee_id, work_date) DO UPDATE SET
            clock_in = excluded.clock_in,
            clock_out = excluded.clock_out,
            lunch_start = excluded.lunch_start,
            lunch_end = excluded.lunch_end,
            regular_hours = excluded.regular_hours,
            overtime_hours = excluded.overtime_hours,
            is_graveyard = excluded.is_graveyard,
            notes = excluded.notes,
            updated_at = excluded.updated_at
    """, (employee_id, work_date, clock_in, clock_out, lunch_start, lunch_end, regular_hours, overtime_hours, is_graveyard, notes or "", now, now))
    conn.commit()


def get_entries_for_week(employee_id, week_start):
    """Get all time entries for one employee for the week (Monday–Sunday)."""
    if isinstance(week_start, str):
        week_start = date.fromisoformat(week_start)
    week_end = week_start + timedelta(days=6)
    start_str = week_start.isoformat()
    end_str = week_end.isoformat()
    with _conn() as conn:
        rows = conn.execute("""
            SELECT * FROM time_entries
            WHERE employee_id = ? AND work_date >= ? AND work_date <= ?
            ORDER BY work_date
        """, (employee_id, start_str, end_str)).fetchall()
        return [dict(r) for r in rows]


def get_all_entries_for_week(week_start):
    """Get all time entries for all employees for the week (for export)."""
    if isinstance(week_start, str):
        week_start = date.fromisoformat(week_start)
    week_end = week_start + timedelta(days=6)
    start_str = week_start.isoformat()
    end_str = week_end.isoformat()
    with _conn() as conn:
        rows = conn.execute("""
            SELECT e.id AS employee_id, e.full_name,
                   t.id AS entry_id, t.work_date, t.clock_in, t.clock_out,
                   t.regular_hours, t.overtime_hours, t.is_graveyard, t.notes
            FROM employees e
            LEFT JOIN time_entries t ON e.id = t.employee_id AND t.work_date >= ? AND t.work_date <= ?
            WHERE e.id IN (SELECT DISTINCT employee_id FROM time_entries WHERE work_date >= ? AND work_date <= ?)
               OR e.id IN (SELECT id FROM employees)
            ORDER BY e.full_name, t.work_date
        """, (start_str, end_str, start_str, end_str)).fetchall()
        # Simpler: get entries in range, then join employee names
        rows = conn.execute("""
            SELECT t.*, e.full_name, e.is_admin
            FROM time_entries t
            JOIN employees e ON e.id = t.employee_id
            WHERE t.work_date >= ? AND t.work_date <= ?
            ORDER BY e.full_name, t.work_date
        """, (start_str, end_str)).fetchall()
        return [dict(r) for r in rows]


def get_entries_for_employee_all_weeks(employee_id, limit_weeks=52):
    """Get recent weeks of entries for an employee (for history)."""
    with _conn() as conn:
        rows = conn.execute("""
            SELECT * FROM time_entries
            WHERE employee_id = ?
            ORDER BY work_date DESC
            LIMIT ?
        """, (employee_id, limit_weeks * 7)).fetchall()
        return [dict(r) for r in rows]


# Time off = notes in ('Sick leave', 'PTO', 'Non Pay')
TIME_OFF_NOTES = ("Sick leave", "PTO", "Non Pay")


def get_timeoff_entries(start_date, end_date, exclude_admin=True):
    """Get all time-off entries (Sick leave) in the date range. Returns list of dicts with employee_id, full_name, work_date, notes (type), shift, regular_hours."""
    if isinstance(start_date, str):
        start_date = date.fromisoformat(start_date)
    if isinstance(end_date, str):
        end_date = date.fromisoformat(end_date)
    start_str = start_date.isoformat()
    end_str = end_date.isoformat()
    placeholders = ",".join("?" * len(TIME_OFF_NOTES))
    with _conn() as conn:
        if exclude_admin:
            rows = conn.execute("""
                SELECT t.employee_id, e.full_name, t.work_date, t.notes, e.shift, t.regular_hours
                FROM time_entries t
                JOIN employees e ON e.id = t.employee_id
                WHERE t.work_date >= ? AND t.work_date <= ?
                  AND t.notes IN (""" + placeholders + """)
                  AND (e.is_admin IS NULL OR e.is_admin = 0)
                ORDER BY e.full_name, t.work_date
            """, [start_str, end_str] + list(TIME_OFF_NOTES)).fetchall()
        else:
            rows = conn.execute("""
                SELECT t.employee_id, e.full_name, t.work_date, t.notes, e.shift, t.regular_hours
                FROM time_entries t
                JOIN employees e ON e.id = t.employee_id
                WHERE t.work_date >= ? AND t.work_date <= ?
                  AND t.notes IN (""" + placeholders + """)
                ORDER BY e.full_name, t.work_date
            """, [start_str, end_str] + list(TIME_OFF_NOTES)).fetchall()
        return [dict(r) for r in rows]


def submit_timeoff(employee_id, from_date, to_date, notes, hours_per_day=8):
    """Create or update time-off entries for each day in [from_date, to_date]. notes must be one of TIME_OFF_NOTES. hours_per_day is stored as regular_hours (e.g. 8 for full day)."""
    if notes not in TIME_OFF_NOTES:
        return
    if isinstance(from_date, str):
        from_date = date.fromisoformat(from_date)
    if isinstance(to_date, str):
        to_date = date.fromisoformat(to_date)
    if from_date > to_date:
        return
    with _conn() as conn:
        d = from_date
        while d <= to_date:
            upsert_time_entry(
                conn, employee_id, d,
                clock_in=None, clock_out=None, lunch_start=None, lunch_end=None,
                notes=notes, regular_hours=hours_per_day, overtime_hours=0, is_graveyard=0,
            )
            d += timedelta(days=1)


def create_timeoff_request(employee_id, from_date, to_date, notes, hours_per_day=8):
    """Create a time-off request with status pending. Returns the new row id or None."""
    if notes not in TIME_OFF_NOTES:
        return None
    if isinstance(from_date, str):
        from_date = date.fromisoformat(from_date)
    if isinstance(to_date, str):
        to_date = date.fromisoformat(to_date)
    if from_date > to_date:
        return None
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        cur = conn.execute("""
            INSERT INTO time_off_requests (employee_id, from_date, to_date, notes, hours_per_day, status, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, 'pending', ?, ?)
        """, (employee_id, from_date.isoformat(), to_date.isoformat(), notes, hours_per_day, now, now))
        conn.commit()
        return cur.lastrowid


def get_pending_timeoff_requests():
    """Return list of pending time-off requests with employee full_name."""
    with _conn() as conn:
        rows = conn.execute("""
            SELECT r.id, r.employee_id, r.from_date, r.to_date, r.notes, r.hours_per_day, r.status, r.created_at,
                   e.full_name
            FROM time_off_requests r
            JOIN employees e ON e.id = r.employee_id
            WHERE r.status = 'pending'
            ORDER BY r.created_at ASC
        """).fetchall()
        return [dict(r) for r in rows]


def get_all_timeoff_requests():
    """Return all time-off requests (pending, approved, rejected) with employee full_name, ordered by created_at DESC."""
    with _conn() as conn:
        rows = conn.execute("""
            SELECT r.id, r.employee_id, r.from_date, r.to_date, r.notes, r.hours_per_day, r.status, r.created_at,
                   e.full_name
            FROM time_off_requests r
            JOIN employees e ON e.id = r.employee_id
            ORDER BY r.created_at DESC
        """).fetchall()
        return [dict(r) for r in rows]


def get_timeoff_request_by_id(request_id):
    """Return a single time-off request by id or None."""
    with _conn() as conn:
        row = conn.execute("""
            SELECT r.id, r.employee_id, r.from_date, r.to_date, r.notes, r.hours_per_day, r.status,
                   e.full_name
            FROM time_off_requests r
            JOIN employees e ON e.id = r.employee_id
            WHERE r.id = ?
        """, (request_id,)).fetchone()
        return dict(row) if row else None


def set_timeoff_request_status(request_id, status):
    """Set request status to 'approved' or 'rejected'. If approved, apply time off to timesheet. Returns True on success."""
    if status not in ("approved", "rejected"):
        return False
    req = get_timeoff_request_by_id(request_id)
    if not req or req.get("status") != "pending":
        return False
    now = datetime.utcnow().isoformat()
    with _conn() as conn:
        conn.execute(
            "UPDATE time_off_requests SET status = ?, updated_at = ? WHERE id = ?",
            (status, now, request_id),
        )
        conn.commit()
    if status == "approved":
        from_d = date.fromisoformat(req["from_date"])
        to_d = date.fromisoformat(req["to_date"])
        submit_timeoff(req["employee_id"], from_d, to_d, req["notes"], hours_per_day=req.get("hours_per_day") or 8)
    return True


def get_employee_timeoff_requests(employee_id):
    """Return all time-off requests for an employee (for status box on Request time off page)."""
    with _conn() as conn:
        rows = conn.execute("""
            SELECT id, from_date, to_date, notes, hours_per_day, status, created_at
            FROM time_off_requests
            WHERE employee_id = ?
            ORDER BY created_at DESC
        """, (employee_id,)).fetchall()
        return [dict(r) for r in rows]

