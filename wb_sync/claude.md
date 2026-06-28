# Warner Bros. Calendar Sync — Claude Code Task

## What this does
Reads the "New Releases" tab of the WB Release Schedule Google Sheet and keeps the
"New Releases – Webstores Clients" Google Calendar in sync.
Runs weekly every Saturday at 08:00 UK via Claude Code routine.

## How to run

```bash
python wb_sync.py
```

For a preview without writing anything:

```bash
python wb_sync.py --dry-run
```

With verbose row audit (logs every skipped row including out-of-window):

```bash
python wb_sync.py --dry-run --verbose
```

With Confluence output (used by the routine to update Confluence pages):

```bash
python wb_sync.py --output-json /tmp/wb_sync_output.json
```

## Environment

| Variable | Required | Description |
|---|---|---|
| `GOOGLE_OAUTH_CREDENTIALS` | Yes | Path to credentials.json (OAuth Desktop app) |
| `GOOGLE_OAUTH_TOKEN` | Yes | Path to token.json (OAuth token with refresh_token) |
| `SLACK_WEBHOOK_URL` | No | Slack incoming webhook URL for change notifications |
| `WB_SHEET_ID` | No | Override spreadsheet ID |
| `WB_SHEET_TAB` | No | Override sheet tab name (default: "New Releases") |
| `WB_CALENDAR_ID` | No | Override calendar ID |

## Source of truth

- Sheet: https://docs.google.com/spreadsheets/d/1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ/edit
- Calendar ID: `c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com`
- Upcoming Confluence page ID: `3101524000`
- Historical Confluence page ID: `3102015518`

## Rules (baked into wb_sync.py)
- Window: today − 1 day → today + 28 days
- Columns (0-based): Title=2, Type=3, Release Type=4, EST=7, VOD=9
- Blank release type + Type=="TV" → treated as "Standard"
- Skip rows: blank title, no valid EST/VOD date, blank release type (non-TV)
- Invite date mapping: Sat→Tue(+3), Sun→Tue(+2), Mon→Tue(+1), Tue→Tue, Wed→Wed, Thu→Fri(+1), Fri→Fri
- Release type order: Pre-Order, Premium, Premium Reprice, Standard, 4K Release
- Availability order: EST, VOD, EST/VOD
- Premium section header omits availability (just "Premium:")
- Deletion: WB Launches events in window with no plan entries are removed
- Events are silent (no notifications)
- Slack posts a changelog after each run (added/removed/created/deleted titles per date)
