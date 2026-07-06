#!/usr/bin/env python3
"""
wb_sync.py — Warner Bros. Calendar Sync
========================================
Reads the "New Releases" tab of the WB Release Schedule Google Sheet and
keeps the "New Releases – Webstores Clients" Google Calendar in sync.

Usage:
    python wb_sync.py [--dry-run] [--output-json PATH] [--verbose]

Authentication:
    Set GOOGLE_OAUTH_CREDENTIALS=/path/to/credentials.json
    Set GOOGLE_OAUTH_TOKEN=/path/to/token.json
"""

import argparse
import json
import os
import sys
import time
import urllib.request
from collections import defaultdict
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

try:
    from google.oauth2 import service_account
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    sys.exit("Missing dependencies. Run:  pip install -r requirements.txt")

# ── Configuration ─────────────────────────────────────────────────────────────
SPREADSHEET_ID   = os.getenv("WB_SHEET_ID",    "1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ")
SHEET_TAB_NAME   = os.getenv("WB_SHEET_TAB",   "New Releases")
CALENDAR_ID      = os.getenv("WB_CALENDAR_ID",
    "c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com")
CALENDAR_TZ      = "Europe/London"
SHEET_PUBLIC_URL = "https://docs.google.com/spreadsheets/d/1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ/edit"
SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL", "")
EVENT_SUMMARY    = "Warner Bros. Launches"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/calendar",
]

# Column indices (0-based). Fallbacks used only if the header row can't be matched by name.
COL_TITLE        = 2   # "Upcoming Releases"
COL_TYPE         = 3   # Film / TV / Bundle
COL_RELEASE_TYPE = 4   # Pre-Order / Premium / Premium Reprice / Standard / 4K Release
COL_EST          = 5   # EST Launch Date
COL_VOD          = 7   # VOD Launch Date

# Header label -> fallback constant, used to re-resolve columns by name each run
# so a reordered/inserted sheet column doesn't silently point at the wrong data.
HEADER_COLUMNS = {
    "upcoming releases": "COL_TITLE",
    "type": "COL_TYPE",
    "release type": "COL_RELEASE_TYPE",
    "est launch date": "COL_EST",
    "vod launch date": "COL_VOD",
}

RELEASE_TYPE_ORDER = {
    "Pre-Order": 1, "Premium": 2, "Premium Reprice": 3, "Standard": 4, "4K Release": 5,
}
AVAILABILITY_ORDER = {"EST": 1, "VOD": 2, "EST/VOD": 3}


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
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(oauth_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as fh:
            fh.write(creds.to_json())
    return creds


# ── Sheet helpers ─────────────────────────────────────────────────────────────

def get_cell(row, idx):
    try:
        return str(row[idx]).strip()
    except IndexError:
        return ""


def _skip_val(v):
    if v is None:
        return True
    return str(v).strip().upper() in {"", "N/A", "NA", "FALSE", "TRUE", "TBC"}


def parse_date(v):
    if _skip_val(v):
        return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%-d/%-m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    parts = s.split("/")
    if len(parts) == 3:
        try:
            return date(int(parts[2]), int(parts[1]), int(parts[0]))
        except ValueError:
            pass
    return None


def fetch_sheet_rows(sheets_svc):
    result = (
        sheets_svc.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_TAB_NAME}'",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
        .execute()
    )
    return result.get("values", [])


# ── Invite date logic ─────────────────────────────────────────────────────────

def get_invite_date(release_date):
    """Map a release date to the earliest Tue/Wed/Fri on or after it."""
    wd = release_date.weekday()  # 0=Mon 1=Tue 2=Wed 3=Thu 4=Fri 5=Sat 6=Sun
    if wd == 1:   return release_date            # Tue → Tue
    if wd == 2:   return release_date            # Wed → Wed
    if wd == 4:   return release_date            # Fri → Fri
    if wd == 0:   return release_date + timedelta(days=1)   # Mon → Tue
    if wd == 3:   return release_date + timedelta(days=1)   # Thu → Fri
    if wd == 5:   return release_date + timedelta(days=3)   # Sat → Tue
    if wd == 6:   return release_date + timedelta(days=2)   # Sun → Tue
    return release_date  # unreachable


# ── Sheet parsing ─────────────────────────────────────────────────────────────

def parse_wb_releases(rows, window_start, window_end, verbose=False):
    """
    Returns:
        plan: dict[invite_date] -> list of entry dicts
        skip_log: list of (row_num, title, reason) for auditing
    """
    plan = defaultdict(list)
    skip_log = []
    n_total = 0

    # Find the header row (look for "Upcoming Releases" in col 2) and resolve
    # column indices by matching header labels, rather than trusting fixed
    # indices — an inserted/reordered sheet column would otherwise silently
    # point every lookup at the wrong data.
    data_start = 0
    header_row = None
    for i, row in enumerate(rows[:10]):
        cell = get_cell(row, COL_TITLE).lower()
        if "upcoming" in cell or "release" in cell or "title" in cell:
            data_start = i + 1
            header_row = row
            break

    col_title, col_type, col_release_type, col_est, col_vod = (
        COL_TITLE, COL_TYPE, COL_RELEASE_TYPE, COL_EST, COL_VOD,
    )
    if header_row is not None:
        resolved = {}
        for idx, cell in enumerate(header_row):
            key = str(cell).strip().lower()
            if key in HEADER_COLUMNS:
                resolved[HEADER_COLUMNS[key]] = idx
        col_title        = resolved.get("COL_TITLE", COL_TITLE)
        col_type         = resolved.get("COL_TYPE", COL_TYPE)
        col_release_type = resolved.get("COL_RELEASE_TYPE", COL_RELEASE_TYPE)
        col_est          = resolved.get("COL_EST", COL_EST)
        col_vod          = resolved.get("COL_VOD", COL_VOD)

    data_rows = rows[data_start:]

    for i, row in enumerate(data_rows):
        row_num = data_start + i + 1  # 1-based for display
        title = get_cell(row, col_title)

        if not title:
            if verbose:
                skip_log.append((row_num, "(blank)", "blank title"))
            continue

        n_total += 1
        content_type = get_cell(row, col_type)
        release_type = get_cell(row, col_release_type)

        # Blank release type + TV → Standard
        if not release_type and content_type.lower() == "tv":
            release_type = "Standard"

        if not release_type:
            skip_log.append((row_num, title, f"blank release type (Type={content_type!r})"))
            continue

        est_date = parse_date(get_cell(row, col_est))
        vod_date = parse_date(get_cell(row, col_vod))

        if est_date is None and vod_date is None:
            skip_log.append((row_num, title, "no valid EST or VOD date"))
            continue

        # Availability
        if est_date and vod_date:
            availability = "EST/VOD"
            release_date = min(est_date, vod_date)
        elif est_date:
            availability = "EST"
            release_date = est_date
        else:
            availability = "VOD"
            release_date = vod_date

        invite_date = get_invite_date(release_date)

        if not (window_start <= invite_date <= window_end):
            if verbose:
                skip_log.append((row_num, title,
                    f"invite date {invite_date} outside window [{window_start}, {window_end}]"))
            continue

        plan[invite_date].append({
            "title": title,
            "release_type": release_type,
            "availability": availability,
            "release_date": release_date,
            "invite_date": invite_date,
        })

    return plan, skip_log, n_total


# ── Description builder ───────────────────────────────────────────────────────

def _section_key(entry):
    rt = RELEASE_TYPE_ORDER.get(entry["release_type"], 99)
    av = AVAILABILITY_ORDER.get(entry["availability"], 99)
    return (rt, av)


def _section_header(release_type, availability):
    if release_type == "Premium":
        return "Premium:"
    return f"{release_type} {availability}:"


def build_description(entries):
    """Build plain-text event description for a list of entries on the same invite date."""
    # Group by (release_type, availability)
    groups = defaultdict(list)
    for e in entries:
        groups[(e["release_type"], e["availability"])].append(e)

    # Sort groups by release_type then availability order
    sorted_keys = sorted(groups.keys(),
        key=lambda k: (RELEASE_TYPE_ORDER.get(k[0], 99), AVAILABILITY_ORDER.get(k[1], 99)))

    sections = []
    for key in sorted_keys:
        rt, av = key
        header = _section_header(rt, av)
        group_entries = sorted(groups[key], key=lambda e: (e["release_date"], e["title"]))
        lines = [header]
        for e in group_entries:
            lines.append(f"• {e['title']}")
        sections.append("\n".join(lines))

    return "\n\n".join(sections)


# ── Calendar helpers ──────────────────────────────────────────────────────────

def dt_london(d, hour):
    tz = ZoneInfo(CALENDAR_TZ)
    return datetime(d.year, d.month, d.day, hour, 0, 0, tzinfo=tz).isoformat()


def list_calendar_events(cal_svc, window_start, window_end):
    all_events = []
    page_token = None
    time_min = datetime(window_start.year, window_start.month, window_start.day,
                        tzinfo=ZoneInfo(CALENDAR_TZ)).isoformat()
    time_max = datetime(window_end.year, window_end.month, window_end.day, 23, 59, 59,
                        tzinfo=ZoneInfo(CALENDAR_TZ)).isoformat()
    while True:
        resp = (
            cal_svc.events()
            .list(
                calendarId=CALENDAR_ID,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                orderBy="startTime",
                maxResults=2500,
                pageToken=page_token,
            )
            .execute()
        )
        all_events.extend(resp.get("items", []))
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
    result = (
        cal_svc.events()
        .insert(calendarId=CALENDAR_ID, body=body, sendNotifications=False)
        .execute()
    )
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
    result = (
        cal_svc.events()
        .update(calendarId=CALENDAR_ID, eventId=event_id, body=body, sendNotifications=False)
        .execute()
    )
    print(f"  UPDATED  {EVENT_SUMMARY!r} on {d} -> {result.get('htmlLink', '')}")
    return result


def delete_event(cal_svc, event_id, d, dry_run):
    if dry_run:
        print(f"  [DRY-RUN] DELETE  {EVENT_SUMMARY!r} on {d}")
        return
    cal_svc.events().delete(calendarId=CALENDAR_ID, eventId=event_id,
                             sendNotifications=False).execute()
    print(f"  DELETED  {EVENT_SUMMARY!r} on {d}")


def retry(fn, *args, retries=2, delay=30, **kwargs):
    for attempt in range(retries):
        try:
            return fn(*args, **kwargs)
        except HttpError as e:
            if attempt < retries - 1:
                print(f"  HTTP error {e.status_code}, retrying in {delay}s...")
                time.sleep(delay)
            else:
                raise


# ── Slack ─────────────────────────────────────────────────────────────────────

def _fmt_titles(titles):
    return ", ".join(f"`{t}`" for t in titles)


def post_slack_summary(stats, today):
    webhook_url = SLACK_WEBHOOK_URL
    if not webhook_url:
        return
    total = stats["updates"] + stats["creates"] + stats["deletes"]
    if total == 0:
        text = f"*Warner Bros. Calendar Sync* — {today}\nNo changes needed, calendar already in sync. :white_check_mark:"
    else:
        lines = [f"*Warner Bros. Calendar Sync* — {today} :calendar:"]
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


def _extract_titles_from_description(desc):
    """Return set of title strings from a plain-text WB event description."""
    if not desc:
        return set()
    titles = set()
    for line in desc.splitlines():
        line = line.strip()
        if line.startswith("•"):
            titles.add(line.lstrip("•").strip())
    return titles


# ── Main sync ─────────────────────────────────────────────────────────────────

def run_sync(dry_run=False, lookback_days=1, lookahead_days=28,
             output_json=None, verbose=False):
    today        = date.today()
    window_start = today - timedelta(days=lookback_days)
    window_end   = today + timedelta(days=lookahead_days)

    print(f"Warner Bros. Calendar Sync -- {today}  (window {window_start} -> {window_end})")
    print(f"  Dry-run: {dry_run}\n")

    creds      = get_credentials()
    sheets_svc = build("sheets",   "v4", credentials=creds)
    cal_svc    = build("calendar", "v3", credentials=creds)

    print("Fetching sheet...")
    rows = retry(fetch_sheet_rows, sheets_svc)
    print(f"  {len(rows)} rows fetched.")

    plan, skip_log, n_total = parse_wb_releases(rows, window_start, window_end, verbose=verbose)

    print(f"  {n_total} titled rows found; {len(skip_log)} skipped.")
    if skip_log:
        print(f"  Skipped rows:")
        for row_num, title, reason in skip_log:
            print(f"    row {row_num}: {title!r} — {reason}")
    print(f"  Invite dates in window: {sorted(plan.keys())}")

    print("\nFetching calendar events...")
    all_events = retry(list_calendar_events, cal_svc, window_start, window_end)
    print(f"  {len(all_events)} events fetched.")

    wb_events_by_date = {}
    for ev in all_events:
        d = event_date(ev)
        if d and ev.get("summary", "") == EVENT_SUMMARY:
            wb_events_by_date[d] = ev

    n_updates = 0
    n_creates = 0
    n_deletes = 0
    update_log = []
    create_log = []
    delete_log = []

    print("\nDiff...")
    for d in sorted(plan.keys()):
        entries     = plan[d]
        new_desc    = build_description(entries)
        new_titles  = {e["title"] for e in entries}

        if d in wb_events_by_date:
            existing     = wb_events_by_date[d]
            current_desc = existing.get("description", "") or ""
            if new_desc.strip() == current_desc.strip():
                print(f"  {d}  {EVENT_SUMMARY} -- already in sync, skip.")
            else:
                old_titles = _extract_titles_from_description(current_desc)
                added      = sorted(new_titles - old_titles)
                removed    = sorted(old_titles - new_titles)
                retry(update_event, cal_svc, existing["id"], d, new_desc, dry_run)
                n_updates += 1
                update_log.append({"date": d, "added": added, "removed": removed})
        else:
            retry(create_event, cal_svc, d, new_desc, dry_run)
            n_creates += 1
            create_log.append({"date": d, "titles": sorted(new_titles)})

    print("\nDeletion check...")
    for d, ev in sorted(wb_events_by_date.items()):
        if d not in plan:
            all_titles = sorted(_extract_titles_from_description(ev.get("description", "")))
            print(f"  {d}  {EVENT_SUMMARY} -- no longer in sheet, deleting.")
            retry(delete_event, cal_svc, ev["id"], d, dry_run)
            n_deletes += 1
            delete_log.append({"date": d, "titles": all_titles or [EVENT_SUMMARY]})

    stats = {
        "updates": n_updates, "creates": n_creates, "deletes": n_deletes,
        "update_log": update_log, "create_log": create_log, "delete_log": delete_log,
    }

    total_changes = n_updates + n_creates + n_deletes
    print("\n" + "-" * 60)
    if total_changes == 0:
        print(f"WB weekly sync -- no changes; {len(wb_events_by_date)} events already in sync")
    else:
        print(f"WB weekly sync -- {n_updates} update(s), {n_creates} create(s), {n_deletes} delete(s)")
    print(f"Sheet: {SHEET_PUBLIC_URL}")
    print("-" * 60)

    if not dry_run:
        post_slack_summary(stats, today)

    if output_json:
        releases = {}
        for d, entries in plan.items():
            releases[str(d)] = [
                {
                    "title": e["title"],
                    "release_type": e["release_type"],
                    "availability": e["availability"],
                    "release_date": str(e["release_date"]),
                }
                for e in entries
            ]
        output = {
            "today": str(today),
            "window_start": str(window_start),
            "window_end": str(window_end),
            "releases": releases,
        }
        with open(output_json, "w") as fh:
            json.dump(output, fh, indent=2)
        print(f"  Release data written to {output_json}")

    return stats


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Warner Bros. calendar sync")
    parser.add_argument("--dry-run",       action="store_true")
    parser.add_argument("--lookback-days", type=int, default=1)
    parser.add_argument("--lookahead-days",type=int, default=28)
    parser.add_argument("--output-json",   default=None)
    parser.add_argument("--verbose",       action="store_true",
                        help="Log every skipped row including out-of-window ones")
    args = parser.parse_args()

    run_sync(
        dry_run=args.dry_run,
        lookback_days=args.lookback_days,
        lookahead_days=args.lookahead_days,
        output_json=args.output_json,
        verbose=args.verbose,
    )
