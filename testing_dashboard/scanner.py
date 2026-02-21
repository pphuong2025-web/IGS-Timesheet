"""
Scan /mnt/L10 on the remote server for today and yesterday.
Parse 6-digit folders and zip filenames; store test results in SQLite.
"""
import re
import os
import stat
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import paramiko
import db

# Zip filename: PREFIX_MODEL_SERIAL_RESULT_STATION_TIMESTAMP.zip
# Model is the segment before the 13-digit serial (last part of prefix).
ZIP_PATTERN = re.compile(
    r"^(.+)_(\d{13})_(P|F)_([A-Z0-9]+)_(\d{8}T\d{6}Z?)\.zip$",
    re.IGNORECASE,
)


def parse_zip_filename(name):
    """
    Return dict with model, serial, result, station, zip_timestamp_taiwan
    or None if not a valid test zip.
    """
    m = ZIP_PATTERN.match(name)
    if not m:
        return None
    prefix, serial, result, station, ts = m.groups()
    # Model = last segment of prefix (after final underscore)
    model = prefix.split("_")[-1] if "_" in prefix else prefix
    return {
        "model": model,
        "serial": serial,
        "result": result.upper(),
        "station": station,
        "zip_timestamp_taiwan": ts,
    }


def load_config():
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    if not os.path.exists(path):
        raise FileNotFoundError("config.json not found. Copy config.example.json to config.json and set credentials.")
    import json
    with open(path, "r") as f:
        return json.load(f)


def get_date_paths():
    """Return (year, month, day) for today and yesterday in server's time (Taiwan)."""
    taiwan = ZoneInfo("Asia/Taipei")
    out = []
    for delta in (0, 1):
        d = datetime.now(taiwan).date() - timedelta(days=delta)
        out.append((d.year, d.month, d.day))
    return out


def scan_once(config):
    """Connect via SFTP, scan today and yesterday, insert new results."""
    cfg = config["server"]
    base = cfg["base_path"].rstrip("/")
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        ssh.connect(
            cfg["host"],
            port=cfg.get("port", 22),
            username=cfg["username"],
            password=cfg["password"],
            timeout=30,
        )
    except Exception as e:
        print(f"SSH connect failed: {e}")
        return

    sftp = ssh.open_sftp()
    ingested_at = datetime.utcnow().isoformat() + "Z"

    try:
        with db._conn() as conn:
            for year, month, day in get_date_paths():
                day_path = f"{base}/{year}/{month:02d}/{day:02d}"
                try:
                    entries = sftp.listdir_attr(day_path)
                except FileNotFoundError:
                    continue
                except Exception as e:
                    print(f"List {day_path}: {e}")
                    continue

                for entry in entries:
                    if getattr(entry, "st_mode", None) and not stat.S_ISDIR(entry.st_mode):
                        continue
                    name = entry.filename
                    if not re.match(r"^\d{6}$", name):
                        continue
                    folder_id = name
                    folder_path = f"{day_path}/{folder_id}"
                    try:
                        files = sftp.listdir_attr(folder_path)
                    except Exception as e:
                        print(f"List {folder_path}: {e}")
                        continue

                    try:
                        folder_mtime = entry.st_mtime
                    except Exception:
                        folder_mtime = None

                    for f in files:
                        if not f.filename.lower().endswith(".zip"):
                            continue
                        parsed = parse_zip_filename(f.filename)
                        if not parsed:
                            continue
                        try:
                            zip_mtime = f.st_mtime
                        except Exception:
                            zip_mtime = None

                        row = {
                            "folder_id": folder_id,
                            "year": year,
                            "month": month,
                            "day": day,
                            "model": parsed["model"],
                            "serial": parsed["serial"],
                            "result": parsed["result"],
                            "station": parsed["station"],
                            "zip_filename": f.filename,
                            "zip_timestamp_taiwan": parsed.get("zip_timestamp_taiwan"),
                            "folder_created_utc": folder_mtime,
                            "zip_created_utc": zip_mtime,
                            "ingested_at": ingested_at,
                        }
                        db.insert_result(conn, row)
            conn.commit()
    finally:
        sftp.close()
        ssh.close()
