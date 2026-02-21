"""
Overtime and graveyard shift calculation.
Work week: Monday–Sunday. Overtime = hours over 8 per day (after deducting lunch).
Graveyard = shift that includes work between 22:00 and 06:00.
"""
from datetime import datetime, date, time, timedelta

import config


def parse_time(s):
    """Parse 'HH:MM' or 'HH:MM:SS' to time. Returns None if invalid."""
    if not s:
        return None
    try:
        parts = s.strip().split(":")
        if len(parts) >= 2:
            h, m = int(parts[0]), int(parts[1])
            sec = int(parts[2]) if len(parts) > 2 else 0
            return time(h, m, sec)
    except (ValueError, IndexError):
        pass
    return None


def time_to_minutes(t):
    """Convert time to minutes since midnight."""
    return t.hour * 60 + t.minute + t.second / 60


def minutes_to_hours(minutes):
    return round(minutes / 60, 2)


def is_graveyard_shift(clock_in_str, clock_out_str):
    """
    Graveyard shift: any work between GRAVEYARD_START_HOUR (22) and GRAVEYARD_END_HOUR (6).
    If clock_in/out span that window, the shift is graveyard.
    """
    clock_in = parse_time(clock_in_str)
    clock_out = parse_time(clock_out_str)
    if not clock_in or not clock_out:
        return False
    start_min = config.GRAVEYARD_START_HOUR * 60   # 22:00 = 1320 min
    end_min = config.GRAVEYARD_END_HOUR * 60       # 06:00 = 360 min
    in_min = time_to_minutes(clock_in)
    out_min = time_to_minutes(clock_out)
    # Overnight shift: e.g. out < in (clock out next day)
    if out_min <= in_min:
        out_min += 24 * 60
    # Check if range [in_min, out_min] overlaps [end_min, start_min] (6am to 10pm) complement
    # Graveyard window in same-day terms: 0..360 (0-6am) and 1320..24*60 (10pm-midnight)
    if in_min < end_min and out_min > in_min:
        return True   # worked into 0-6am
    if in_min < 24 * 60 and out_min > start_min:
        return True   # worked 10pm+
    if in_min >= start_min or out_min <= end_min:
        return True
    return False


def day_hours(clock_in_str, clock_out_str, lunch_start_str=None, lunch_end_str=None):
    """Compute total working hours for one day: (clock_out - clock_in) minus lunch. Returns (total_hours, 0.0)."""
    clock_in = parse_time(clock_in_str)
    clock_out = parse_time(clock_out_str)
    if not clock_in or not clock_out:
        return 0.0, 0.0
    in_min = time_to_minutes(clock_in)
    out_min = time_to_minutes(clock_out)
    if out_min <= in_min:
        out_min += 24 * 60
    total_min = out_min - in_min
    # Deduct lunch if both start and end are provided
    lunch_start = parse_time(lunch_start_str) if lunch_start_str else None
    lunch_end = parse_time(lunch_end_str) if lunch_end_str else None
    if lunch_start and lunch_end:
        ls_min = time_to_minutes(lunch_start)
        le_min = time_to_minutes(lunch_end)
        if le_min <= ls_min:
            le_min += 24 * 60
        lunch_min = le_min - ls_min
        total_min = max(0, total_min - lunch_min)
    total_hours = minutes_to_hours(total_min)
    return total_hours, 0.0  # per-day we don't split; weekly does


def _day_total_hours(entry):
    """Total hours for one entry from clock times (minus lunch) or stored hours. Non Pay = 0."""
    if (entry.get("notes") or "").strip() == "Non Pay":
        return 0.0
    r = entry.get("regular_hours") or 0.0
    o = entry.get("overtime_hours") or 0.0
    if r + o > 0:
        return r + o
    if entry.get("clock_in") and entry.get("clock_out"):
        h, _ = day_hours(
            entry.get("clock_in"),
            entry.get("clock_out"),
            entry.get("lunch_start"),
            entry.get("lunch_end"),
        )
        return h
    return 0.0


def compute_weekly_overtime(entries):
    """
    entries: list of dicts for one employee, one week, sorted by work_date.
    Returns list of entries with regular_hours, overtime_hours, is_graveyard set.
    Daily rule: working hours (after lunch deduction) over REGULAR_HOURS_PER_DAY = overtime for that day.
    """
    cap = config.REGULAR_HOURS_PER_DAY
    result = []
    for e in entries:
        # Day total = clock_out - clock_in - lunch (already used in _day_total_hours)
        day_total = _day_total_hours(e)
        is_grav = bool(e.get("is_graveyard"))
        if not is_grav and (e.get("clock_in") and e.get("clock_out")):
            is_grav = is_graveyard_shift(e.get("clock_in"), e.get("clock_out"))
        regular = round(min(day_total, cap), 2)
        overtime = round(max(0, day_total - cap), 2)
        out = {**e, "regular_hours": regular, "overtime_hours": overtime, "is_graveyard": 1 if is_grav else 0}
        result.append(out)
    return result
