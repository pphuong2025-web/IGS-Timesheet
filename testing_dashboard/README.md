# L10 Testing Dashboard

Dashboard for L10 test results. Data is scanned from a remote server (`/mnt/L10`) every 5 minutes and stored in SQLite. All times in the UI are shown in **Pacific (PST/PDT)**. The zip filename timestamp is in Taiwan time; the dashboard uses **server file creation time** (zip/folder mtime) for display, which is converted to Pacific.

## Setup

1. **Python 3.9+** (for `zoneinfo`).

2. **Install dependencies:**
   ```bash
   cd testing_dashboard
   pip install -r requirements.txt
   ```

3. **Config:** Copy `config.example.json` to `config.json` and set your server credentials:
   ```json
   {
     "server": {
       "host": "172.10.16.67",
       "port": 22,
       "username": "YOUR_USERNAME",
       "password": "YOUR_PASSWORD",
       "base_path": "/mnt/L10"
     },
     "dashboard": {
       "host": "0.0.0.0",
       "port": 5000
     }
   }
   ```
   Do not commit `config.json` (it is in `.gitignore`).

## Run

From the `testing_dashboard` folder:

```bash
python app.py
```

- Open **on this PC:** http://127.0.0.1:5000  
- Open **from another machine on the network:** http://\<this-PC-IP\>:5000  

To find this PC’s IP (Windows): `ipconfig` and use the IPv4 address (e.g. `192.168.1.100`).

## Behavior

- **Scanner:** Runs at startup (after 10 seconds) and then **every 5 minutes**. It scans **today and yesterday** (Taiwan date) under `/mnt/L10/yyyy/mm/dd/`, finds 6-digit folders and `.zip` files inside, parses filenames (model, serial, pass/fail, station), and stores new rows in `tests.db` (duplicates by folder + zip name are skipped).
- **Dashboard:** Summary (total, pass, fail, pass rate), tests per hour (Pacific), by station, by model, and a recent-tests table. Use the time range filter (e.g. Last 24 hours, Today Pacific) and click **Apply**.

## Files

| File | Purpose |
|------|--------|
| `config.json` | Server credentials and dashboard host/port (create from `config.example.json`) |
| `app.py` | Flask app + 5‑min scheduler |
| `scanner.py` | SFTP scan and zip parsing |
| `db.py` | SQLite schema and queries |
| `index.html` | Dashboard UI |
| `tests.db` | SQLite DB (created automatically) |

## Time zones

- **Zip filename timestamp:** Taiwan (Asia/Taipei). Stored as-is; optional for display later.
- **Folder/zip creation time:** From server (SFTP mtime, typically UTC). Stored as Unix timestamp and shown in the dashboard in **Pacific (America/Los_Angeles)**.
