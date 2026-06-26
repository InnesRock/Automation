# NBCU Calendar Sync — Claude Code Task

## What this does
Reads the "New Releases" tab of the NBCU Film Release Schedule Google Sheet and
keeps the "New Releases – Webstores Clients" Google Calendar in sync.  
Run weekly (e.g. every Friday morning).

## How to run

​```bash
python nbcu_sync.py
​```

For a preview without writing anything:

​```bash
python nbcu_sync.py --dry-run
​```

## Environment

| Variable | Required | Description |
|---|---|---|
| `GOOGLE_APPLICATION_CREDENTIALS` | Yes (service account) | Path to service_account.json |
| `GOOGLE_OAUTH_CREDENTIALS` | Yes (OAuth fallback) | Path to credentials.json |
| `NBCU_SHEET_ID` | No | Override spreadsheet ID |
| `NBCU_CALENDAR_ID` | No | Override calendar ID |

## Source of truth

- Sheet: https://docs.google.com/spreadsheets/d/1xgVyk1VOSiLeJjt7BKE9hqlO_Cp3_RYYZzPrAO6_6JQ/edit?gid=1497414736
- Calendar ID: `c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com`

## Rules (baked into nbcu_sync.py)
- Window: today − 14 days → today + 120 days
- Category 1: Film/NR/4K/Library rows with a real title → "NBCU Launches" events (one per date, all titles bulleted)
- Category 2: Library rows with empty title + "x catalog launches" identifier → separate "NBCU N x Catalog Titles Launching" events
- Never delete existing NBCU events — only create or update
- Preserve any manual `<p><a href=...>` link prefix in existing event descriptions
- EST + VOD on same date/regions → collapsed to "EST/VOD"
- 4K prefix applied when column C == "4K"
- Events are silent (no notifications)
