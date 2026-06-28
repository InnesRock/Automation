# Disney Calendar Sync — Claude Code Task

## What this does
Downloads the Disney Studio Title List XLSX and parser script from Google Drive,
runs the parser, and keeps the "New Releases – Webstores Clients" Google Calendar
in sync with Disney releases.
Runs weekly every Wednesday at 08:00 UK via Claude Code routine.

## How to run

```bash
python disney_sync.py
```

For a preview without writing anything:

```bash
python disney_sync.py --dry-run
```

With Confluence output (used by the routine to update Confluence pages):

```bash
python disney_sync.py --output-json /tmp/disney_sync_output.json
```

## Environment

| Variable | Required | Description |
|---|---|---|
| `GOOGLE_OAUTH_CREDENTIALS` | Yes | Path to credentials.json (OAuth Desktop app) |
| `GOOGLE_OAUTH_TOKEN` | Yes | Path to token.json — MUST include drive.readonly scope |
| `SLACK_WEBHOOK_URL` | No | Slack incoming webhook URL for change notifications |
| `DISNEY_XLSX_ID` | No | Override Drive file ID for the XLSX spreadsheet |
| `DISNEY_SCRIPT_ID` | No | Override Drive file ID for disney_plan.py |
| `DISNEY_CALENDAR_ID` | No | Override calendar ID |

## ⚠️ Re-authentication required
Disney sync needs `drive.readonly` scope (to download the XLSX from Drive).
The NBCU/WB token only has `spreadsheets.readonly` + `calendar`.
Delete token.json and re-authenticate once to get a token with all three scopes.
Update `GOOGLE_OAUTH_TOKEN_JSON` in all three routines with the new token.

## Source of truth

- Spreadsheet Drive ID: `1iVxfN6or2RObpSwOJsMy0MNE_bJt2ixt4KdnZZ6oQao`
- Parser script Drive ID: `1Ev-VzOnh5hmlZa-JPfatyaXxZMhwsR-m`
- Calendar ID: `c_ca57d25f2d93e30e42baa4e389da3b8fd871b3bb0f6acdf8b3330fdb7cc35d57@group.calendar.google.com`
- Upcoming Confluence page ID: `3102081040`
- Historical Confluence page ID: `3100737567`

## Rules (baked into disney_sync.py)
- Future window: today → today + 120 days
- Historical window: today − 730 days → yesterday (for Confluence Historical page)
- Parser script is downloaded fresh from Drive on each run (picks up manual fixes)
- Stale TV cleanup: delete Disney Launches events on Mon/Thu/Sat/Sun if all titles end
  with (TV) and the date is not in the parser output
- Events are silent (no notifications)
- Slack posts a changelog after each run (added/removed/created/deleted titles per date)
- Confluence Historical page is rebuilt from calendar events, not the parser output
