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
- Columns (0-based): MVPD check=0, Title=2, Type=3 (Film / Film Bundle / TV / TV Boxset), Release Type=4, Launch Date=5
  - The sheet used to have separate EST/VOD date columns; it's now a single "Launch
    Date" column (Warner Bros. always launches Film titles EST/VOD simultaneously,
    and TV/Bundle titles EST-only, so there's nothing left to disambiguate).
    Availability (EST/VOD) is no longer tracked or shown anywhere downstream.
  - "Film Bundle" and "TV Boxset" each get their own header (Film Bundles / TV
    Boxsets) -- they are NOT merged into the plain Film/TV headers.
- Blank release type + Type in ("TV", "TV Boxset") → treated as "Standard"
- Skip rows: blank title, no valid launch date, blank release type (non-TV/TV Boxset)
- Invite date mapping: Sat→Tue(+3), Sun→Tue(+2), Mon→Tue(+1), Tue→Tue, Wed→Wed, Thu→Fri(+1), Fri→Fri
- Release type order: Pre-Order, Premium, Premium Reprice, Standard, 4K Release, EST-only, VOD
  - EST-only/VOD are ordinary Release Type values (column E), unrelated to the old
    EST/VOD Availability concept that was removed -- they render exactly like any
    other Release Type (bulleted, colon-suffixed, underlined, with titles nested
    underneath), just sorted after 4K Release since there's no stronger signal on
    where else they'd belong.
- Type order: Film, Film Bundle, TV, TV Boxset -- headers "Films", "Film Bundles", "TV", "TV Boxsets"
- Calendar description layout — grouped Type -> Release Type -> Title, using Google
  Calendar's native bulleted-list HTML (`<ul>`/`<li>`, not manual "•" characters) so
  it renders with real disc/circle nested bullets, matching the NBCU calendar's style:
  ```
  Warner Bros. releases landing today (5):   (bold summary line, count = all releases that date)
                                              (blank line)
  Films                                      (bold header, no bullet; omitted if no titles of that type that date)
  • Premium:                                 (Release Type, bulleted, colon-suffixed, underlined)
     o Title A                               (Title, nested bullet, no styling)
        ▪ MVPD Check required                (italic; only if that row's MVPD check column == "Yes")
  • Standard:
     o Title B
  Film Bundles
  ...
  TV
  ...
  TV Boxsets
  ...
  ```
  Built as real HTML: `<b>...(N):</b><br><br><b>Films</b><ul><li><u>Premium:</u><ul><li>Title A<ul><li><i>MVPD Check required</i></li></ul></li></ul></li>...</ul>`.
  Titles/release types are HTML-escaped when building this. Type headers are bold
  only (no underline) -- only the Release Type line is underlined. The MVPD sub-
  bullet's italics are Calendar-only (Confluence renders it as plain unstyled text,
  same as any other bullet); everything else in this layout (native bullets,
  colon+underline, the MVPD sub-bullet's existence/wording/position) is shared
  between Calendar and Confluence, with Confluence using its own native nested
  `<ul>`/`<li>` bullets (Confluence storage format) rather than raw HTML, and no
  summary count line (that stays Calendar-only). Confluence Type header uses
  `<p><strong>...</strong></p>`.
- MVPD check column ("Yes"/"No"/"tbc"/blank, case-insensitive match on "yes") adds a
  nested "MVPD Check required" sub-bullet directly under that title (italic on
  Calendar, plain on Confluence) -- nothing else changes for that entry.
- MVPD Check invites: any row with "MVPD Check required" also gets its own separate
  calendar invite, one per flagged row (so a title flagged on both its Premium and
  Premium Reprice rows gets two invites, on their respective dates):
  - Title: `Warner Bros. MVPD Check - [TITLE]`
  - Time: 8:30am-9:00am `America/New_York` (the rest of the sync uses Europe/London),
    on that row's invite date
  - Description (`<br>`-joined, same as the main invite's HTML description field):
    ```
    [TITLE]
    Platforms: DirecTV, Verizon, Xfinity
    Type: [RELEASE TYPE]
    Expected SRPs: [BUY], [RENT]

    Campaign description: Tracking Compliance and Merchandising for "[TITLE]", across
    the title's Premium and Premium Reprice release windows.

    Segment descriptions:
    Compliance for the [RELEASE TYPE] of "[TITLE]" on [DATE].
    Merchandising support for the [RELEASE TYPE] of "[TITLE]" on [DATE].
    ```
    `[RELEASE TYPE]` is that row's Release Type value (Premium / Premium Reprice);
    `[DATE]` is that row's own invite date, formatted "4th August 2026" (same value
    both times). `[BUY]`/`[RENT]` come from `MVPD_PRICING`: Premium = $24.99
    PEST / $19.99 PVOD, Premium Reprice = $19.99 PEST / $9.99 PVOD. A flagged row
    with any other Release Type gets a warning logged and blank SRPs (pricing is
    only defined for Premium/Premium Reprice).
  - Sync lifecycle: created once per (date, title) key if missing; deleted if that
    key drops out of the plan (row removed from sheet, or MVPD flag flips to "No")
    on a later run. **Never updated** once created -- a manual edit to the SRP
    pricing (or anything else in the description) is left alone on subsequent runs,
    by design.
- Deletion: WB Launches events in window with no plan entries are removed
- Events are silent (no notifications)
- Slack posts a changelog after each run (added/removed/created/deleted titles per date)
- `--output-json` entries: `{title, type, release_type, release_date, mvpd_check}` (bool; no `availability` field)
