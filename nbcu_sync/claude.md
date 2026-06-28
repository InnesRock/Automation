# NBCU Calendar Sync — Claude Code Task

## What this does
Reads the NBCU Film Release Schedule Google Sheet and keeps the
"New Releases – Webstores Clients" Google Calendar in sync.
Runs weekly every Tuesday at 01:00 UK via Claude Code routine.

## How to run

```bash
python nbcu_sync.py
```

For a preview without writing anything:

```bash
python nbcu_sync.py --dry-run
```

With Confluence output (used by the routine to update Confluence pages):

```bash
python nbcu_sync.py --output-json /tmp/nbcu_sync_output.json
```

## Environment

| Variable | Required | Description |
|---|---|---|
| `GOOGLE_OAUTH_CREDENTIALS` | Yes | Path to credentials.json (OAuth Desktop app) |
| `GOOGLE_OAUTH_TOKEN` | Yes | Path to token.json (OAuth token with refresh_token) |
| `SLACK_WEBHOOK_URL` | No | Slack incoming webhook URL for change notifications |
| `NBCU_SHEET_ID` | No | Override spreadsheet ID |
| `NBCU_CALENDAR_ID` | No | Override calendar ID |

## Source of truth

- Sheet: https://docs.google.com/spreadsheets/d/1xgVyk1VOSiLeJjt7BKE9hqlO_Cp3_RYYZzPrAO6_6JQ/edit
- Calendar ID: `c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com`
- Upcoming Confluence page ID: `3102081068`
- Historical Confluence page ID: `3101720593`

## Rules (baked into nbcu_sync.py)
- Window: today − 14 days → today + 120 days
- Category 1: Film/NR/4K/Library rows with a real title → "NBCU Launches" events (one per date, all titles bulleted)
- Category 2: Library rows with empty title + "x catalog launches" identifier → separate "NBCU N x Catalog Titles Launching" events
- Deletion: events in window no longer present in sheet are removed
- EST + VOD on same date/regions → collapsed to "EST/VOD"
- 4K prefix applied when column C == "4K"
- Events are silent (no notifications)
- Slack posts a changelog after each run (added/removed/created/deleted titles per date)
