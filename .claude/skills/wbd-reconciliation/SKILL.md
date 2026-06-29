---
name: wbd-reconciliation
description: >
  Runs the weekly WBD Digital Calendar vs Release Schedule reconciliation. Use this skill
  immediately whenever the user uploads a file with "Digital Calendar" or "WBD" in the name,
  or asks to "check the DC", "run the reconciliation", "compare against the RS", "DC check",
  or "reconcile the digital calendar". The skill fetches fresh RS data from Google Drive
  at the start of every run — never uses cached data — then runs the Python comparison script
  and reports mismatches. Do not ask clarifying questions; start the workflow immediately.
---

# WBD Digital Calendar Reconciliation

Reconciles a weekly WBD Digital Calendar (DC) xlsx file against the live Warner Bros.
Release Schedule (RS) from Google Sheets, flagging date mismatches and missing titles.

---

## Step 0 — Locate the DC file

**If the user has not uploaded a file**, search the WBD DC Inbox Drive folder for the
most recent Digital Calendar attachment saved there by the Gmail→Drive script:

```
Tool: mcp__Google_Drive__search_files
folderId: 1c9w7H0TeVEDg3hH9rw0dI0d8_vcU7tG5
```

Take the file with the latest `modifiedTime` whose name contains "Digital Calendar"
or "WBD". Download it:

```
Tool: mcp__Google_Drive__download_file_content
fileId: <id from above>
```

Save the binary content to the uploads folder so the reconciliation script can read it
as an xlsx. If no file is found in Drive, ask the user to upload the DC file manually.

---

## Step 1 — Fetch fresh RS data from Google Drive

**Do this first, before anything else, on every single run. Never reuse a cached `rs_newreleases.csv`.**

Use the Google Drive connector to read the "New Releases" sheet from the RS Google Sheet:

- **Spreadsheet ID:** `1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ`
- **Sheet tab:** "New Releases" (gid=114630171)

Fetch all rows from that sheet. Then write ALL rows verbatim to the outputs folder as
`rs_newreleases.csv`. Include a header row:

```
go_live,vendor_id,title,type,release_type,col5,col6,est_launch,col8,vod_launch
```

**Critical:** write every row returned — do not cherry-pick or manually select a subset.
Missing even one row will cause false "MISSING FROM RS" flags.
The comparison script filters to relevant release types automatically.

The CSV must preserve the RS column order exactly:

| Index | Field |
|-------|-------|
| 0 | go_live flag |
| 1 | vendor_id |
| 2 | title |
| 3 | type (Film / TV / Bundle / 4K) |
| 4 | release_type |
| 5 | col5 |
| 6 | col6 |
| 7 | EST Launch Date (DD/MM/YYYY) |
| 8 | col8 |
| 9 | VOD Launch Date (DD/MM/YYYY or blank) |

---

## Step 2 — Ensure the reconciliation script exists

Check that `wbd_reconcile.py` is present in the outputs folder at the bash-accessible path
(`/sessions/.../mnt/outputs/wbd_reconcile.py`). If it is missing, copy the bundled version:

```bash
cp <skill_dir>/scripts/wbd_reconcile.py /sessions/.../mnt/outputs/wbd_reconcile.py
```

---

## Step 3 — Run the reconciliation

```bash
python3 /sessions/.../mnt/outputs/wbd_reconcile.py \
  "/sessions/.../mnt/uploads/<DC_filename>.xlsx" \
  "/sessions/.../mnt/outputs/rs_newreleases.csv"
```

Use the actual session paths from the environment.

---

## Step 4 — Report results

- **No issues:** "✅ All X entries pass — no mismatches found."
- **Issues found:** List each one with the title, what mismatched, and DC vs RS values.

Keep it concise — no need to list every passing title unless the user asks.

---

## Reconciliation rules

| DC field | Checked against | RS release_type |
|----------|-----------------|-----------------|
| PEST, PVOD | RS EST Launch Date, VOD Launch Date | `Premium` only (not "Premium Reprice") |
| EST, VOD | RS EST Launch Date, VOD Launch Date | `Standard` only |

Additional flags:
- PEST ≠ PVOD in DC → flag internal DC mismatch.
- DC has PEST/PVOD but RS has no Premium row → "NO PREMIUM ROW IN RS".
- DC has EST/VOD but RS has no Standard row → "NO STANDARD ROW IN RS".
- DC title not found in RS at all → "MISSING FROM RS".
- Multiple Standard rows: prefer the one whose EST matches DC; otherwise use highest row number.
- Title matching: case-insensitive, strips articles (the/a/an), punctuation, year suffixes
  like "(2026)", and annotations like "(Ani)".

---

## DC column layout (0-indexed, xlsx)

| Index | Field |
|-------|-------|
| 0 | Title |
| 2 | PEST (MM/DD/YYYY text) |
| 4 | PVOD (MM/DD/YYYY text) |
| 6 | EST (MM/DD/YYYY text) |
| 8 | VOD (MM/DD/YYYY text) |

---

## Project notes

- RS Google Sheet: `https://docs.google.com/spreadsheets/d/1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ/`
- Sheet tab: "New Releases" (gid=114630171)
- The DC is uploaded fresh each week by the user.
- `wbd_reconcile.py` is bundled in `scripts/` within this skill directory.
