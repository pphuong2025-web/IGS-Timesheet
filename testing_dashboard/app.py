"""
Flask app: dashboard UI and API. Runs scanner every 5 minutes.
Bind to 0.0.0.0 so others on the network can access.
"""
from nt import lseek
import os
from re import L
from flask import Flask, send_from_directory, request, jsonify
from apscheduler.schedulers.background import BackgroundScheduler

import db
import scanner

app = Flask(__name__, static_folder="static")
BASE = os.path.dirname(os.path.abspath(__file__))


def run_scan():
    try:
        config = scanner.load_config()
        scanner.scan_once(config)
    except Exception as e:
        print(f"Scan error: {e}")


# Run scanner every 5 minutes when config exists
try:
    scanner.load_config()
    scheduler = BackgroundScheduler()
    scheduler.add_job(run_scan, "interval", minutes=5, id="l10_scan")
    scheduler.start()
    from threading import Timer
    Timer(10, run_scan).start()
except FileNotFoundError:
    scheduler = None  # no config: dashboard still works; use "Load sample data" to see data


@app.route("/")
def index():
    return send_from_directory(BASE, "index.html")


@app.route("/api/stats")
def api_stats():
    from_ts = request.args.get("from")
    to_ts = request.args.get("to")
    if from_ts:
        try:
            from_ts = float(from_ts)
        except ValueError:
            from_ts = None
    if to_ts:
        try:
            to_ts = float(to_ts)
        except ValueError:
            to_ts = None
    data = db.get_stats(from_ts=from_ts, to_ts=to_ts)
    return jsonify(data)


@app.route("/api/seed-sample", methods=["POST"])
def api_seed_sample():
    """Insert sample data so the dashboard shows something when the DB is empty."""
    try:
        count = db.seed_sample_data()
        return jsonify({"ok": True, "count": count})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/recent")
def api_recent():
    limit = request.args.get("limit", 100, type=int)
    limit = min(max(limit, 1), 2000)
    from_ts = request.args.get("from")
    to_ts = request.args.get("to")
    if from_ts:
        try:
            from_ts = float(from_ts)
        except ValueError:
            from_ts = None
    if to_ts:
        try:
            to_ts = float(to_ts)
        except ValueError:
            to_ts = None
    rows = db.get_recent(limit=limit, from_ts=from_ts, to_ts=to_ts)
    return jsonify(rows)


if __name__ == "__main__":
    db.init_db()
    cfg = {}
    try:
        cfg = scanner.load_config()
    except FileNotFoundError:
        pass
    host = (cfg.get("dashboard") or {}).get("host", "0.0.0.0")
    port = (cfg.get("dashboard") or {}).get("port", 5000)
    print(f"Dashboard: http://{host}:{port} (network: http://<this-pc-ip>:{port})")
    app.run(host=host, port=port, threaded=True)
cls
l
