#!/usr/bin/env python3
"""
disney_sync.py — Disney Calendar Sync
=======================================
Downloads the Disney XLSX and parser script from Google Drive, runs the
parser, and syncs results to the "New Releases – Webstores Clients" Calendar.

Usage:
    python disney_sync.py [--dry-run] [--output-json PATH]

Authentication:
    Set GOOGLE_OAUTH_CREDENTIALS=/path/to/credentials.json
    Set GOOGLE_OAUTH_TOKEN=/path/to/token.json

    NOTE: Requires drive.readonly scope in addition to calendar.
    If upgrading from NBCU/WB token, delete token.json and re-authenticate.
"""

import argparse
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import urllib.request
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

try:
    from google.oauth2 import service_account
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    from googleapiclient.http import MediaIoBaseDownload
except ImportError:
    sys.exit("Missing dependencies. Run:  pip install -r requirements.txt")

# ── Configuration ─────────────────────────────────────────────────────────────
XLSX_FILE_ID    = os.getenv("DISNEY_XLSX_ID",    "1iVxfN6or2RObpSwOJsMy0MNE_bJt2ixt4KdnZZ6oQao")
SCRIPT_FILE_ID  = os.getenv("DISNEY_SCRIPT_ID",  "1Ev-VzOnh5hmlZa-JPfatyaXxZMhwsR-m")
CALENDAR_ID     = os.getenv("DISNEY_CALENDAR_ID",
    "c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com")
CALENDAR_TZ     = "Europe/London"
EVENT_SUMMARY   = "Disney Launches"
SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL", "")

LOOKAHEAD_DAYS  = 120
LOOKBACK_DAYS   = 730   # 2 years for historical Confluence page

# Mon=0, Thu=3, Sat=5, Sun=6 — stale TV events only deleted on these weekdays
STALE_TV_WEEKDAYS = {0, 3, 5, 6}

SCOPES = [
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/calendar",
]


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_credentials():
    sa_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if sa_path and os.path.exists(sa_path):
        return service_account.Credentials.from_service_account_file(sa_path, scopes=SCOPES)

    oauth_path = os.getenv("GOOGLE_OAUTH_CREDENTIALS", "credentials.json")
    token_path = os.getenv("GOOGLE_OAUTH_TOKEN", "token.json")

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                creds = None
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(oauth_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as fh:
            fh.write(creds.to_json())
    return creds


# ── Drive helpers ─────────────────────────────────────────────────────────────

def download_drive_file(drive_svc, file_id, dest_path, export_mime=None):
    """Download a file from Drive. Use export_mime for Google Sheets → XLSX."""
    if export_mime:
        req = drive_svc.files().export_media(fileId=file_id, mimeType=export_mime)
    else:
        req = drive_svc.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, req)
        done = False
        while not done:
            _, done = downloader.next_chunk()


# ── Parser runner ─────────────────────────────────────────────────────────────

def run_parser(script_path, xlsx_path):
    env = os.environ.copy()
    env["DISNEY_SHEET_PATH"] = xlsx_path
    env["DISNEY_USE_TODAY"] = "1"
    result = subprocess.run(
        [sys.executable, script_path, "emit"],
        capture_output=True, text=True, env=env, timeout=120,
    )
    if result.returncode != 0:
        print(f"  Parser stderr:\n{result.stderr[:3000]}")
        raise RuntimeError(f"Parser exited with code {result.returncode}")
    return json.loads(result.stdout)


# ── Calendar helpers ──────────────────────────────────────────────────────────

def dt_london(d, hour):
    tz = ZoneInfo(CALENDAR_TZ)
    return datetime(d.year, d.month, d.day, hour, 0, 0, tzinfo=tz).isoformat()


def list_calendar_events(cal_svc, start, end):
    all_events = []
    page_token = None
    time_min = datetime(start.year, start.month, start.day,
                        tzinfo=ZoneInfo(CALENDAR_TZ)).isoformat()
    time_max = datetime(end.year, end.month, end.day, 23, 59, 59,
                        tzinfo=ZoneInfo(CALENDAR_TZ)).isoformat()
    while True:
        resp = cal_svc.events().list(
            calendarId=CALENDAR_ID,
            timeMin=time_min, timeMax=time_max,
            singleEvents=True, orderBy="startTime",
            maxResults=2500, pageToken=page_token,
        ).execute()
        all_events.extend(
            e for e in resp.get("items", []) if e.get("summary") == EVENT_SUMMARY
        )
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return all_events


def event_date(event):
    start = event.get("start", {})
    dt_str = start.get("dateTime") or start.get("date")
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str).date()
    except ValueError:
        return None


def create_event(cal_svc, d, description, dry_run):
    body = {
        "summary": EVENT_SUMMARY,
        "description": description,
        "start": {"dateTime": dt_london(d, 9), "timeZone": CALENDAR_TZ},
        "end":   {"dateTime": dt_london(d, 10), "timeZone": CALENDAR_TZ},
        "reminders": {"useDefault": False, "overrides": []},
    }
    if dry_run:
        print(f"  [DRY-RUN] CREATE  {EVENT_SUMMARY!r} on {d}")
        return None
    result = cal_svc.events().insert(
        calendarId=CALENDAR_ID, body=body, sendNotifications=False
    ).execute()
    print(f"  CREATED  {EVENT_SUMMARY!r} on {d} -> {result.get('htmlLink', '')}")
    return result


def update_event(cal_svc, event_id, d, description, dry_run):
    body = {
        "summary": EVENT_SUMMARY,
        "description": description,
        "start": {"dateTime": dt_london(d, 9), "timeZone": CALENDAR_TZ},
        "end":   {"dateTime": dt_london(d, 10), "timeZone": CALENDAR_TZ},
        "reminders": {"useDefault": False, "overrides": []},
    }
    if dry_run:
        print(f"  [DRY-RUN] UPDATE  {EVENT_SUMMARY!r} on {d}")
        return None
    result = cal_svc.events().update(
        calendarId=CALENDAR_ID, eventId=event_id, body=body, sendNotifications=False
    ).execute()
    print(f"  UPDATED  {EVENT_SUMMARY!r} on {d} -> {result.get('htmlLink', '')}")
    return result


def delete_event(cal_svc, event_id, d, dry_run):
    if dry_run:
        print(f"  [DRY-RUN] DELETE  {EVENT_SUMMARY!r} on {d}")
        return
    cal_svc.events().delete(
        calendarId=CALENDAR_ID, eventId=event_id, sendNotifications=False
    ).execute()
    print(f"  DELETED  {EVENT_SUMMARY!r} on {d}")


def retry(fn, *args, retries=2, delay=30, **kwargs):
    for attempt in range(retries):
        try:
            return fn(*args, **kwargs)
        except HttpError as e:
            if attempt < retries - 1:
                print(f"  HTTP {e.status_code}, retrying in {delay}s...")
                time.sleep(delay)
            else:
                raise


# ── Stale TV detection ────────────────────────────────────────────────────────

def is_all_tv(description):
    """True if every bolded title in the HTML description ends with (TV)."""
    titles = re.findall(r"<b>(.*?)</b>", description or "")
    if not titles:
        return False
    return all(t.strip().endswith("(TV)") for t in titles)


def extract_titles(description):
    """Return set of bold titles from HTML description."""
    return set(re.findall(r"<b>(.*?)</b>", description or ""))


# ── Slack ─────────────────────────────────────────────────────────────────────

def _fmt_titles(titles):
    return ", ".join(f"`{t}`" for t in titles)


def post_slack_summary(stats, today):
    webhook_url = SLACK_WEBHOOK_URL
    if not webhook_url:
        return
    total = stats["updates"] + stats["creates"] + stats["deletes"]
    if total == 0:
        text = (f"*Disney Calendar Sync* — {today}\n"
                f"No changes needed, calendar already in sync. :white_check_mark:")
    else:
        lines = [f"*Disney Calendar Sync* — {today} :calendar:"]
        lines.append(
            f"*{stats['updates']} updated · {stats['creates']} created · {stats['deletes']} deleted*"
        )
        if stats.get("update_log"):
            lines.append("\n*Updated:*")
            for entry in stats["update_log"]:
                d_str = entry["date"].strftime("%-d %b")
                for t in entry.get("added", []):
                    lines.append(f"• {d_str}: added — `{t}`")
                for t in entry.get("removed", []):
                    lines.append(f"• {d_str}: removed — `{t}`")
        if stats.get("create_log"):
            lines.append("\n*Created:*")
            for entry in stats["create_log"]:
                d_str = entry["date"].strftime("%-d %b")
                lines.append(f"• {d_str} — {_fmt_titles(entry['titles'])}")
        if stats.get("delete_log"):
            lines.append("\n*Deleted:*")
            for entry in stats["delete_log"]:
                d_str = entry["date"].strftime("%-d %b")
                lines.append(f"• {d_str} — {_fmt_titles(entry['titles'])}")
        text = "\n".join(lines)

    payload = json.dumps({"text": text}).encode()
    req = urllib.request.Request(
        webhook_url, data=payload, headers={"Content-Type": "application/json"})
    try:
        urllib.request.urlopen(req, timeout=10)
        print("  Slack notification sent.")
    except Exception as e:
        print(f"  Slack notification failed: {e}")


# ── Main sync ─────────────────────────────────────────────────────────────────

def run_sync(dry_run=False, output_json=None):
    today        = date.today()
    window_end   = today + timedelta(days=LOOKAHEAD_DAYS)
    history_start = today - timedelta(days=LOOKBACK_DAYS)

    print(f"Disney Calendar Sync -- {today}  (future window {today} -> {window_end})")
    print(f"  Dry-run: {dry_run}\n")

    creds     = get_credentials()
    drive_svc = build("drive",    "v3", credentials=creds)
    cal_svc   = build("calendar", "v3", credentials=creds)

    # ── Download assets from Drive ────────────────────────────────────────────
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path   = os.path.join(tmpdir, "disney.xlsx")
        script_path = os.path.join(tmpdir, "disney_plan.py")

        print("Downloading Disney spreadsheet from Drive...")
        retry(download_drive_file, drive_svc, XLSX_FILE_ID, xlsx_path,
              export_mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        print(f"  OK ({os.path.getsize(xlsx_path):,} bytes)")

        print("Downloading parser script from Drive...")
        retry(download_drive_file, drive_svc, SCRIPT_FILE_ID, script_path)
        print(f"  OK ({os.path.getsize(script_path):,} bytes)")

        print("Running parser (DISNEY_USE_TODAY=1)...")
        parser_output = run_parser(script_path, xlsx_path)
        print(f"  Parser returned {len(parser_output)} release date(s).")

    # Key by date object
    plan = {date.fromisoformat(e["date"]): e for e in parser_output}

    # ── Fetch future calendar events ──────────────────────────────────────────
    print("\nFetching future calendar events...")
    future_events = retry(list_calendar_events, cal_svc, today, window_end)
    future_by_date = {}
    for ev in future_events:
        d = event_date(ev)
        if d:
            future_by_date[d] = ev
    print(f"  {len(future_by_date)} Disney Launches events in window.")

    n_updates = n_creates = n_deletes = 0
    update_log, create_log, delete_log = [], [], []

    # ── Diff ──────────────────────────────────────────────────────────────────
    print("\nDiff...")
    for d in sorted(plan.keys()):
        entry    = plan[d]
        new_desc = entry["description"]
        new_titles = extract_titles(new_desc)

        if d in future_by_date:
            existing     = future_by_date[d]
            current_desc = existing.get("description", "") or ""
            if new_desc.strip() == current_desc.strip():
                print(f"  {d}  already in sync, skip.")
            else:
                old_titles = extract_titles(current_desc)
                added   = sorted(new_titles - old_titles)
                removed = sorted(old_titles - new_titles)
                retry(update_event, cal_svc, existing["id"], d, new_desc, dry_run)
                n_updates += 1
                update_log.append({"date": d, "added": added, "removed": removed})
        else:
            retry(create_event, cal_svc, d, new_desc, dry_run)
            n_creates += 1
            create_log.append({"date": d, "titles": sorted(new_titles)})

    # ── Stale TV cleanup ──────────────────────────────────────────────────────
    print("\nStale TV cleanup...")
    for d, ev in sorted(future_by_date.items()):
        if d not in plan:
            desc = ev.get("description", "") or ""
            if d.weekday() in STALE_TV_WEEKDAYS and is_all_tv(desc):
                titles = sorted(extract_titles(desc))
                print(f"  {d}  stale TV-only event, deleting.")
                retry(delete_event, cal_svc, ev["id"], d, dry_run)
                n_deletes += 1
                delete_log.append({"date": d, "titles": titles})
            else:
                print(f"  {d}  not in plan but not stale TV — leaving.")

    stats = {
        "updates": n_updates, "creates": n_creates, "deletes": n_deletes,
        "update_log": update_log, "create_log": create_log, "delete_log": delete_log,
    }

    print("\n" + "-" * 60)
    total = n_updates + n_creates + n_deletes
    if total == 0:
        print(f"Disney sync -- no changes; {len(future_by_date)} events already in sync")
    else:
        print(f"Disney sync -- {n_updates} update(s), {n_creates} create(s), {n_deletes} delete(s)")
    print("-" * 60)

    if not dry_run:
        post_slack_summary(stats, today)

    # ── Output JSON for Confluence ────────────────────────────────────────────
    if output_json:
        print("\nFetching historical calendar events for Confluence...")
        yesterday = today - timedelta(days=1)
        past_events = retry(list_calendar_events, cal_svc, history_start, yesterday)
        past_events_sorted = sorted(
            [e for e in past_events if event_date(e)],
            key=lambda e: event_date(e),
            reverse=True,
        )
        print(f"  {len(past_events_sorted)} past Disney Launches events fetched.")

        output = {
            "today": str(today),
            "upcoming_events": [
                {"date": e["date"], "description": e["description"], "count": e["count"]}
                for e in parser_output
            ],
            "past_events": [
                {
                    "date": str(event_date(e)),
                    "description": e.get("description", ""),
                }
                for e in past_events_sorted
            ],
        }
        with open(output_json, "w") as fh:
            json.dump(output, fh, indent=2)
        print(f"  Written to {output_json}")

    return stats


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Disney calendar sync")
    parser.add_argument("--dry-run",     action="store_true")
    parser.add_argument("--output-json", default=None)
    args = parser.parse_args()
    run_sync(dry_run=args.dry_run, output_json=args.output_json)
