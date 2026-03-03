"""
Overtime and shift calculation (day / swing / graveyard).
Work week: Monday–Sunday. Overtime = hours over 8 per day (after deducting lunch).
Shift = whichever window (day, swing, graveyard) contains the most work minutes.
"""
from datetime import datetime, date, time, timedelta

import config

MINUTES_PER_DAY = 24 * 60


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


def _overlap_minutes(seg_lo, seg_hi, w_lo, w_hi):
    """Minutes of segment [seg_lo, seg_hi] that fall in window [w_lo, w_hi]. All in 0..1440."""
    return max(0, min(seg_hi, w_hi) - max(seg_lo, w_lo))


def _minutes_in_window(in_min, out_min, w_start_min, w_end_min):
    """Minutes of shift [in_min, out_min] that fall in same-day window [w_start_min, w_end_min]. Handles overnight (out_min > 1440)."""
    total = 0
    if out_min <= MINUTES_PER_DAY:
        total += _overlap_minutes(in_min, out_min, w_start_min, w_end_min)
    else:
        total += _overlap_minutes(in_min, MINUTES_PER_DAY, w_start_min, w_end_min)
        total += _overlap_minutes(0, out_min - MINUTES_PER_DAY, w_start_min, w_end_min)
    return total


def classify_shift(clock_in_str, clock_out_str):
    """
    Classify shift as day, swing, or graveyard by which window contains the most work minutes.
    Day: DAY_SHIFT_START to DAY_SHIFT_END (e.g. 7:00–15:30).
    Swing: SWING_SHIFT_START to SWING_SHIFT_END (e.g. 15:00–23:45).
    Graveyard: GRAVEYARD_START to GRAVEYARD_END (22:00–06:00).
    Returns "day", "swing", or "graveyard". Returns None if no clock in/out.
    """
    clock_in = parse_time(clock_in_str)
    clock_out = parse_time(clock_out_str)
    if not clock_in or not clock_out:
        return None
    in_min = time_to_minutes(clock_in)
    out_min = time_to_minutes(clock_out)
    if out_min <= in_min:
        out_min += MINUTES_PER_DAY
    day_start = int(config.DAY_SHIFT_START_HOUR * 60)
    day_end = int(config.DAY_SHIFT_END_HOUR * 60)
    swing_start = int(config.SWING_SHIFT_START_HOUR * 60)
    swing_end = int(config.SWING_SHIFT_END_HOUR * 60)
    grav_start = config.GRAVEYARD_START_HOUR * 60
    grav_end = config.GRAVEYARD_END_HOUR * 60
    day_min = _minutes_in_window(in_min, out_min, day_start, day_end)
    swing_min = _minutes_in_window(in_min, out_min, swing_start, swing_end)
    grav_min = _minutes_in_window(in_min, out_min, 0, grav_end) + _minutes_in_window(in_min, out_min, grav_start, MINUTES_PER_DAY)
    if grav_min >= day_min and grav_min >= swing_min:
        return "graveyard"
    if swing_min >= day_min:
        return "swing"
    return "day"


def is_graveyard_shift(clock_in_str, clock_out_str):
    """True if classify_shift returns graveyard. Kept for backward compatibility."""
    return classify_shift(clock_in_str, clock_out_str) == "graveyard"


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
    Returns list of entries with regular_hours, overtime_hours, shift, is_graveyard set.
    Daily rule: working hours (after lunch deduction) over REGULAR_HOURS_PER_DAY = overtime for that day.
    Shift = day/swing/graveyard from classify_shift (most work minutes in that window).
    """
    cap = config.REGULAR_HOURS_PER_DAY
    result = []
    for e in entries:
        day_total = _day_total_hours(e)
        shift = (e.get("shift") or "").strip().lower() or None
        if (e.get("clock_in") and e.get("clock_out")):
            classified = classify_shift(e.get("clock_in"), e.get("clock_out"))
            if classified:
                shift = classified
        if shift not in ("day", "swing", "graveyard"):
            shift = None
        regular = round(min(day_total, cap), 2)
        overtime = round(max(0, day_total - cap), 2)
        is_grav = 1 if shift == "graveyard" else 0
        out = {**e, "regular_hours": regular, "overtime_hours": overtime, "shift": shift, "is_graveyard": is_grav}
        result.append(out)
    return result
