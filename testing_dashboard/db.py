"""
SQLite schema and helpers for L10 test results.
"""
import sqlite3
import os
from datetime import datetime
from contextlib import contextmanager

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tests.db")


def init_db():
    """Create tables if they don't exist."""
    with _conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS test_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                folder_id TEXT NOT NULL,
                year INTEGER NOT NULL,
                month INTEGER NOT NULL,
                day INTEGER NOT NULL,
                model TEXT NOT NULL,
                serial TEXT NOT NULL,
                result TEXT NOT NULL,
                station TEXT NOT NULL,
                zip_filename TEXT NOT NULL,
                zip_timestamp_taiwan TEXT,
                folder_created_utc REAL,
                zip_created_utc REAL,
                ingested_at TEXT NOT NULL,
                UNIQUE(folder_id, zip_filename)
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_result ON test_results(result)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_station ON test_results(station)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_model ON test_results(model)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ymd ON test_results(year, month, day)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_zip_created ON test_results(zip_created_utc)")
        conn.commit()


@contextmanager
def _conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
    finally:
        conn.close()


def insert_result(conn, row):
    """Insert one test result. Ignore if duplicate (folder_id + zip_filename)."""
    conn.execute("""
        INSERT OR IGNORE INTO test_results (
            folder_id, year, month, day, model, serial, result, station,
            zip_filename, zip_timestamp_taiwan, folder_created_utc, zip_created_utc, ingested_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row["folder_id"],
        row["year"],
        row["month"],
        row["day"],
        row["model"],
        row["serial"],
        row["result"],
        row["station"],
        row["zip_filename"],
        row.get("zip_timestamp_taiwan"),
        row.get("folder_created_utc"),
        row.get("zip_created_utc"),
        row["ingested_at"],
    ))


def get_stats(from_ts=None, to_ts=None):
    """Pass/fail counts and by station/model. Optional time filter on zip_created_utc."""
    with _conn() as conn:
        args = []
        where = ""
        if from_ts is not None:
            where += " AND zip_created_utc >= ?"
            args.append(from_ts)
        if to_ts is not None:
            where += " AND zip_created_utc <= ?"
            args.append(to_ts)

        # Overall pass/fail
        cursor = conn.execute(
            f"SELECT result, COUNT(*) as cnt FROM test_results WHERE 1=1 {where} GROUP BY result",
            args,
        )
        by_result = {row["result"]: row["cnt"] for row in cursor.fetchall()}

        # By station
        cursor = conn.execute(
            f"SELECT station, result, COUNT(*) as cnt FROM test_results WHERE 1=1 {where} GROUP BY station, result",
            args,
        )
        by_station = {}
        for row in cursor.fetchall():
            st = row["station"]
            if st not in by_station:
                by_station[st] = {"P": 0, "F": 0}
            by_station[st][row["result"]] = row["cnt"]

        # By model
        cursor = conn.execute(
            f"SELECT model, result, COUNT(*) as cnt FROM test_results WHERE 1=1 {where} GROUP BY model, result",
            args,
        )
        by_model = {}
        for row in cursor.fetchall():
            m = row["model"]
            if m not in by_model:
                by_model[m] = {"P": 0, "F": 0}
            by_model[m][row["result"]] = row["cnt"]

        # Tests per hour (bucket by hour in UTC, then we can convert in frontend if needed)
        cursor = conn.execute(
            f"""
            SELECT CAST(strftime('%Y-%m-%d %H:00', zip_created_utc, 'unixepoch') AS TEXT) as hour_utc,
                   COUNT(*) as cnt
            FROM test_results WHERE zip_created_utc IS NOT NULL {where}
            GROUP BY hour_utc ORDER BY hour_utc
            """,
            args,
        )
        tests_per_hour = [{"hour_utc": row["hour_utc"], "count": row["cnt"]} for row in cursor.fetchall()]

    return {
        "by_result": by_result,
        "by_station": by_station,
        "by_model": by_model,
        "tests_per_hour": tests_per_hour,
    }


def get_recent(limit=100, from_ts=None, to_ts=None):
    """Recent tests list. Optional time filter."""
    with _conn() as conn:
        where = ""
        args = []
        if from_ts is not None:
            where += " AND zip_created_utc >= ?"
            args.append(from_ts)
        if to_ts is not None:
            where += " AND zip_created_utc <= ?"
            args.append(to_ts)
        args.append(limit)

        cursor = conn.execute(
            f"""
            SELECT folder_id, year, month, day, model, serial, result, station,
                   zip_filename, zip_timestamp_taiwan, folder_created_utc, zip_created_utc, ingested_at
            FROM test_results WHERE 1=1 {where}
            ORDER BY zip_created_utc DESC, id DESC
            LIMIT ?
            """,
            args,
        )
        rows = cursor.fetchall()
    return [dict(r) for r in rows]


def seed_sample_data():
    """Insert sample test rows so the dashboard shows data when the scanner has not run yet."""
    import time
    now = time.time()
    # Spread samples across the last 24 hours so "Last 24 hours" and "Today" show them
    samples = [
        ("104727", 2026, 2, 5, "675-24109-0002-TS2", "1830326000021", "F", "FLA",
         "IGSJ_PB-65984_675-24109-0002-TS2_1830326000021_F_FLA_20260204T161044Z.zip",
         "20260204T161044Z", now - 3600 * 2, now - 3600 * 2),
        ("104845", 2026, 2, 5, "675-24109-0002-TS1", "1830226000123", "F", "FLA",
         "IGSJ_675-24109-0002-TS1_1830226000123_F_FLA_20260205T102044Z.zip",
         "20260205T102044Z", now - 3600 * 5, now - 3600 * 4),
        ("105012", 2026, 2, 5, "675-24109-0010-TS2", "1830526000035", "P", "FLB",
         "IGSJ_675-24109-0010-TS2_1830526000035_P_FLB_20260205T120000Z.zip",
         "20260205T120000Z", now - 3600 * 3, now - 3600 * 2.5),
        ("105123", 2026, 2, 5, "675-24109-0002-TS2", "1830326000022", "P", "FLA",
         "IGSJ_675-24109-0002-TS2_1830326000022_P_FLA_20260205T131000Z.zip",
         "20260205T131000Z", now - 3600 * 1.5, now - 3600 * 1),
        ("105234", 2026, 2, 5, "675-24109-0000-TS1", "1830125000269", "F", "FCT",
         "IGSJ_675-24109-0000-TS1_1830125000269_F_FCT_20260205T140000Z.zip",
         "20260205T140000Z", now - 3600 * 0.5, now - 300),
    ]
    ingested = datetime.utcnow().isoformat() + "Z"
    with _conn() as conn:
        for folder_id, year, month, day, model, serial, result, station, zip_filename, ts_tw, folder_utc, zip_utc in samples:
            conn.execute("""
                INSERT OR IGNORE INTO test_results (
                    folder_id, year, month, day, model, serial, result, station,
                    zip_filename, zip_timestamp_taiwan, folder_created_utc, zip_created_utc, ingested_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (folder_id, year, month, day, model, serial, result, station, zip_filename, ts_tw, folder_utc, zip_utc, ingested))
        conn.commit()
        cursor = conn.execute("SELECT COUNT(*) FROM test_results")
        count = cursor.fetchone()[0]
    return count
