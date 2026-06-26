#!/usr/bin/env python3
"""
nbcu_sync.py — NBCU Film Release Calendar Sync
================================================
Reads the "New Releases" tab of the NBCU Film Release Schedule Google Sheet
and keeps the "New Releases – Webstores Clients" Google Calendar in sync.

Usage:
    python nbcu_sync.py [--dry-run] [--lookback-days 14] [--lookahead-days 120]

Authentication:
    Service-account (recommended for remote/cron use):
        Set GOOGLE_APPLICATION_CREDENTIALS=/path/to/service_account.json
    OAuth (interactive, first-run only):
        Set GOOGLE_OAUTH_CREDENTIALS=/path/to/credentials.json
        A browser window will open on first run; token is cached in token.json.
"""

import argparse
import html as html_module
import json
import os
import re
import sys
import time
import urllib.request
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

# ── Google API imports ────────────────────────────────────────────────────────
try:
    from google.oauth2 import service_account
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:
    sys.exit(
        "Missing dependencies. Run:  pip install -r requirements.txt"
    )

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION  (override with env vars or edit here)
# ─────────────────────────────────────────────────────────────────────────────
SPREADSHEET_ID   = os.getenv("NBCU_SHEET_ID",    "1xgVyk1VOSiLeJjt7BKE9hqlO_Cp3_RYYZzPrAO6_6JQ")
SHEET_TAB_NAME   = os.getenv("NBCU_SHEET_TAB",   "New Releases")
CALENDAR_ID      = os.getenv("NBCU_CALENDAR_ID", "c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com")
CALENDAR_TZ      = "Europe/London"
SHEET_PUBLIC_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1xgVyk1VOSiLeJjt7BKE9hqlO_Cp3_RYYZzPrAO6_6JQ"
    "/edit?gid=1497414736#gid=1497414736"
)
SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL", "")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/calendar",
]

COL = {
    "VID":               0,   # A
    "TITLE":             1,   # B
    "TYPE":              2,   # C  NR / 4K / Library (or Film)
    "IS_4K_ONLY":        3,   # D
    "TIER":              4,   # E
    "ADDED_TO_CAL":      5,   # F  IGNORED
    "POEST_US":          6,   # G
    "POEST_CA":          7,   # H
    "PREMIUM_US":        8,   # I
    "PREMIUM_US_AVAIL":  9,   # J  ignored
    "PREMIUM_CA":       10,   # K
    "PREMIUM_CA_AVAIL": 11,   # L  ignored
    "REPRICE_US":       12,   # M
    "REPRICE_CA":       13,   # N
    "EST_US":           14,   # O
    "EST_US_AVAIL":     15,   # P  ignored
    "EST_CA":           16,   # Q
    "EST_CA_AVAIL":     17,   # R  ignored
    "VOD_US":           18,   # S
    "VOD_US_AVAIL":     19,   # T  ignored
    "VOD_CA":           20,   # U
    "VOD_CA_AVAIL":     21,   # V  ignored
}

DATE_COLS = [
    (COL["POEST_US"],   "Pre-Order",        "US"),
    (COL["POEST_CA"],   "Pre-Order",        "CA"),
    (COL["PREMIUM_US"], "Premium",          "US"),
    (COL["PREMIUM_CA"], "Premium",          "CA"),
    (COL["REPRICE_US"], "Premium Reprice",  "US"),
    (COL["REPRICE_CA"], "Premium Reprice",  "CA"),
    (COL["EST_US"],     "EST",              "US"),
    (COL["EST_CA"],     "EST",              "CA"),
    (COL["VOD_US"],     "VOD",              "US"),
    (COL["VOD_CA"],     "VOD",              "CA"),
]

TYPE_PRECEDENCE = {
    "Pre-Order":       1,
    "Premium":         2,
    "Premium Reprice": 3,
    "EST":             4,
    "VOD":             5,
    "EST/VOD":         4,
}


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


def _skip_val(v):
    if v is None:
        return True
    s = str(v).strip().lower()
    return s in {"", "n/a", "na", "tbc"}


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


def dt_london(d, hour):
    tz = ZoneInfo(CALENDAR_TZ)
    return datetime(d.year, d.month, d.day, hour, 0, 0, tzinfo=tz).isoformat()


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


def get_cell(row, idx):
    try:
        return str(row[idx]).strip()
    except IndexError:
        return ""


def is_hidden_row(row):
    return not get_cell(row, COL["VID"]) and not get_cell(row, COL["TITLE"])


def parse_new_releases(rows, window_start, window_end):
    data_rows = rows[2:]
    plan_cat1 = {}
    plan_cat2 = []

    for row in data_rows:
        if is_hidden_row(row):
            continue

        vid   = get_cell(row, COL["VID"])
        title = get_cell(row, COL["TITLE"])
        rtype = get_cell(row, COL["TYPE"]).strip()

        if rtype.lower() == "library" and not title:
            identifier = vid
            if not identifier:
                continue
            if "(now disabled)" in identifier.lower():
                continue
            seen_dates = set()
            for col_idx, _, _ in DATE_COLS:
                d = parse_date(get_cell(row, col_idx))
                if d and window_start <= d <= window_end and d not in seen_dates:
                    seen_dates.add(d)
                    plan_cat2.append({"identifier": identifier, "date": d})
            continue

        if not title or title.lower() == "tbc":
            continue

        is_4k = rtype.lower() == "4k"
        prefix = "4K " if is_4k else ""

        for col_idx, type_label, region in DATE_COLS:
            d = parse_date(get_cell(row, col_idx))
            if d is None or not (window_start <= d <= window_end):
                continue
            entry = {"title": title, "type": prefix + type_label, "region": region}
            plan_cat1.setdefault(d, []).append(entry)

    return plan_cat1, plan_cat2


def build_description(entries, manual_prefix=""):
    from collections import defaultdict
    title_type_regions = defaultdict(lambda: defaultdict(set))
    for e in entries:
        title_type_regions[e["title"]][e["type"]].add(e["region"])

    def merged_types(type_regions):
        est_r = type_regions.get("EST", type_regions.get("4K EST", None))
        vod_r = type_regions.get("VOD", type_regions.get("4K VOD", None))
        pfx = "4K " if "4K EST" in type_regions or "4K VOD" in type_regions else ""
        result = dict(type_regions)
        if est_r is not None and vod_r is not None and est_r == vod_r:
            result.pop("EST", None)
            result.pop("VOD", None)
            result.pop("4K EST", None)
            result.pop("4K VOD", None)
            result[pfx + "EST/VOD"] = est_r
        return result

    def region_groups(type_regions):
        tr = merged_types(type_regions)
        rs_types = defaultdict(list)
        for t, regions in tr.items():
            rs_types[frozenset(regions)].append(t)
        groups = []
        for rs, types in rs_types.items():
            types_sorted = sorted(types, key=lambda t: TYPE_PRECEDENCE.get(t.replace("4K ", ""), 99))
            groups.append((rs, types_sorted))
        groups.sort(key=lambda g: TYPE_PRECEDENCE.get(g[1][0].replace("4K ", ""), 99))
        return groups

    titles = sorted(title_type_regions.keys())
    n = len(titles)
    html = f"<p><b>NBCU releases landing today ({n}):</b></p><ul>"

    for title in titles:
        html += f"<li><b>{_esc(title)}</b><ul>"
        groups = region_groups(title_type_regions[title])
        for region_set, types in groups:
            region_list = sorted(region_set, key=lambda r: (0 if r == "US" else 1))
            region_str = ", ".join(region_list)
            html += f"<li>Regions: {region_str}<ul>"
            for t in types:
                html += f"<li>Type: {_esc(t)}</li>"
            html += "</ul></li>"
        html += "</ul></li>"

    html += "</ul>"
    if manual_prefix:
        return manual_prefix + html
    return html


def _esc(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def catalog_description():
    return (
        f'<a href="{SHEET_PUBLIC_URL}" target="_blank">'
        f"NBCU Film Release Schedule &amp; Title List</a>"
    )


def catalog_summary(identifier):
    s = re.sub(r"catalog launches", "Catalog Titles Launching", identifier, flags=re.IGNORECASE)
    return f"NBCU {s}"


_LINK_PREFIX_RE = re.compile(r'^(<p><a\s[^>]*>.*?</a></p>)+', re.IGNORECASE | re.DOTALL)

def extract_manual_prefix(description):
    if not description:
        return ""
    m = _LINK_PREFIX_RE.match(description)
    return m.group(0) if m else ""


def extract_titles_from_description(desc):
    """Return the set of title strings from an existing NBCU Launches event description."""
    if not desc:
        return set()
    return {html_module.unescape(t) for t in re.findall(r'<li><b>(.*?)</b>', desc)}


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


def create_event(cal_svc, summary, d, description, dry_run):
    body = {
        "summary": summary,
        "description": description,
        "start": {"dateTime": dt_london(d, 9), "timeZone": CALENDAR_TZ},
        "end":   {"dateTime": dt_london(d, 10), "timeZone": CALENDAR_TZ},
        "reminders": {"useDefault": False, "overrides": []},
    }
    if dry_run:
        print(f"  [DRY-RUN] CREATE  {summary!r} on {d}")
        return None
    result = (
        cal_svc.events()
        .insert(calendarId=CALENDAR_ID, body=body, sendNotifications=False)
        .execute()
    )
    print(f"  CREATED  {summary!r} on {d} -> {result.get('htmlLink','')}")
    return result


def update_event(cal_svc, event_id, summary, d, description, dry_run):
    body = {
        "summary": summary,
        "description": description,
        "start": {"dateTime": dt_london(d, 9), "timeZone": CALENDAR_TZ},
        "end":   {"dateTime": dt_london(d, 10), "timeZone": CALENDAR_TZ},
        "reminders": {"useDefault": False, "overrides": []},
    }
    if dry_run:
        print(f"  [DRY-RUN] UPDATE  {summary!r} on {d}")
        return None
    result = (
        cal_svc.events()
        .update(calendarId=CALENDAR_ID, eventId=event_id, body=body, sendNotifications=False)
        .execute()
    )
    print(f"  UPDATED  {summary!r} on {d} -> {result.get('htmlLink','')}")
    return result


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


def _fmt_titles(titles):
    return ", ".join(f"`{t}`" for t in titles)


def post_slack_summary(stats, today):
    webhook_url = SLACK_WEBHOOK_URL
    if not webhook_url:
        return
    total = stats["updates"] + stats["creates"] + stats["deletes"]
    if total == 0:
        text = f"*NBCU Calendar Sync* — {today}\nNo changes needed, calendar already in sync. :white_check_mark:"
    else:
        lines = [f"*NBCU Calendar Sync* — {today} :calendar:"]
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
        webhook_url,
        data=payload,
        headers={"Content-Type": "application/json"},
    )
    try:
        urllib.request.urlopen(req, timeout=10)
        print("  Slack notification sent.")
    except Exception as e:
        print(f"  Slack notification failed: {e}")


def run_sync(dry_run=False, lookback_days=14, lookahead_days=120, output_json=None):
    today        = date.today()
    window_start = today - timedelta(days=lookback_days)
    window_end   = today + timedelta(days=lookahead_days)

    print(f"NBCU Calendar Sync -- {today}  (window {window_start} -> {window_end})")
    print(f"  Dry-run: {dry_run}\n")

    creds      = get_credentials()
    sheets_svc = build("sheets",   "v4", credentials=creds)
    cal_svc    = build("calendar", "v3", credentials=creds)

    print("Fetching sheet...")
    rows = retry(fetch_sheet_rows, sheets_svc)
    print(f"  {len(rows)} rows fetched (including headers).")

    plan_cat1, plan_cat2 = parse_new_releases(rows, window_start, window_end)
    print(f"  Cat-1 dates in window: {sorted(plan_cat1.keys())}")
    print(f"  Cat-2 entries in window: {[(e['identifier'], str(e['date'])) for e in plan_cat2]}")

    print("\nFetching calendar events...")
    all_events = retry(list_calendar_events, cal_svc, window_start, window_end)
    print(f"  {len(all_events)} events fetched.")

    nbcu_launches_by_date = {}
    catalog_events_by_date = {}

    for ev in all_events:
        summary = ev.get("summary", "")
        d = event_date(ev)
        if d is None:
            continue
        if summary == "NBCU Launches":
            nbcu_launches_by_date[d] = ev
        elif re.match(r"^NBCU \d+ x Catalog Titles Launching$", summary):
            catalog_events_by_date[d] = ev

    n_updates = 0
    n_creates = 0
    n_deletes = 0
    affected_links = []
    update_log = []
    create_log = []
    delete_log = []

    print("\nCategory-1 diff (NBCU Launches)...")
    for d in sorted(plan_cat1.keys()):
        entries    = plan_cat1[d]
        new_desc   = build_description(entries)
        new_titles = {e["title"] for e in entries}

        if d in nbcu_launches_by_date:
            existing = nbcu_launches_by_date[d]
            current_desc = existing.get("description", "") or ""
            manual_prefix = extract_manual_prefix(current_desc)
            final_desc = (manual_prefix + new_desc) if manual_prefix else new_desc
            if final_desc.strip() == current_desc.strip():
                print(f"  {d}  NBCU Launches -- already in sync, skip.")
            else:
                old_titles = extract_titles_from_description(current_desc)
                added   = sorted(new_titles - old_titles)
                removed = sorted(old_titles - new_titles)
                result = retry(update_event, cal_svc,
                               existing["id"], "NBCU Launches", d, final_desc, dry_run)
                n_updates += 1
                update_log.append({"date": d, "added": added, "removed": removed})
                if result:
                    affected_links.append(result.get("htmlLink", ""))
        else:
            result = retry(create_event, cal_svc, "NBCU Launches", d, new_desc, dry_run)
            n_creates += 1
            create_log.append({"date": d, "titles": sorted(new_titles)})
            if result:
                affected_links.append(result.get("htmlLink", ""))

    n_cat2_creates = 0
    print("\nCategory-2 diff (Catalog launches)...")
    seen_cat2 = set()
    for entry in plan_cat2:
        d   = entry["date"]
        key = (entry["identifier"], d)
        if key in seen_cat2:
            continue
        seen_cat2.add(key)
        ev_summary = catalog_summary(entry["identifier"])
        desc       = catalog_description()
        if d in catalog_events_by_date:
            print(f"  {d}  {ev_summary!r} -- already exists, skip.")
        else:
            result = retry(create_event, cal_svc, ev_summary, d, desc, dry_run)
            n_creates += 1
            n_cat2_creates += 1
            create_log.append({"date": d, "titles": [ev_summary]})
            if result:
                affected_links.append(result.get("htmlLink", ""))

    print("\nDeletion check...")
    for d, ev in sorted(nbcu_launches_by_date.items()):
        if d not in plan_cat1:
            print(f"  {d}  NBCU Launches -- no longer in sheet, deleting.")
            if not dry_run:
                cal_svc.events().delete(calendarId=CALENDAR_ID, eventId=ev["id"],
                                        sendNotifications=False).execute()
            n_deletes += 1
            delete_log.append({"date": d, "titles": ["NBCU Launches"]})

    plan_cat2_dates = {e["date"] for e in plan_cat2}
    for d, ev in sorted(catalog_events_by_date.items()):
        if d not in plan_cat2_dates:
            ev_summary = ev.get("summary", "Catalog event")
            print(f"  {d}  {ev_summary!r} -- no longer in sheet, deleting.")
            if not dry_run:
                cal_svc.events().delete(calendarId=CALENDAR_ID, eventId=ev["id"],
                                        sendNotifications=False).execute()
            n_deletes += 1
            delete_log.append({"date": d, "titles": [ev_summary]})

    stats = {"updates": n_updates, "creates": n_creates, "deletes": n_deletes,
             "update_log": update_log, "create_log": create_log, "delete_log": delete_log}

    total_changes = n_updates + n_creates + n_deletes
    print("\n" + "-" * 60)
    if total_changes == 0:
        existing_nbcu = len(nbcu_launches_by_date) + len(catalog_events_by_date)
        print(f"NBCU weekly sync -- no changes; {existing_nbcu} events already in sync")
    else:
        print(
            f"NBCU weekly sync -- {n_updates} update(s), "
            f"{n_creates} create(s) (including {n_cat2_creates} catalog launch(es)), "
            f"{n_deletes} delete(s)"
        )
        if affected_links:
            print("Affected events:")
            for link in affected_links[:3]:
                print(f"  {link}")
    print(f"Sheet: {SHEET_PUBLIC_URL}")
    print("-" * 60)

    if not dry_run:
        post_slack_summary(stats, today)

    if output_json:
        releases = {}
        for d, entries in plan_cat1.items():
            releases[str(d)] = [
                {"title": e["title"], "type": e["type"], "region": e["region"]}
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
    parser = argparse.ArgumentParser(description="NBCU calendar sync")
    parser.add_argument("--dry-run",        action="store_true",
                        help="Print planned changes without writing to calendar")
    parser.add_argument("--lookback-days",  type=int, default=14)
    parser.add_argument("--lookahead-days", type=int, default=120)
    parser.add_argument("--output-json",    default=None,
                        help="Write release data as JSON to this path (for Confluence updates)")
    args = parser.parse_args()

    run_sync(
        dry_run=args.dry_run,
        lookback_days=args.lookback_days,
        lookahead_days=args.lookahead_days,
        output_json=args.output_json,
    )
