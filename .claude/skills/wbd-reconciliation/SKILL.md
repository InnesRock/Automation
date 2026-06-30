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
or "WBD". Download it as XLSX:

```
Tool: mcp__Google_Drive__download_file_content
fileId: <id from above>
exportMimeType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
```

The tool returns a base64-encoded binary. Decode it and write the bytes to the uploads
folder as a `.xlsx` file so the reconciliation script can open it.

**If the user provides a Google Sheets URL for the DC**, extract the file ID from the URL
(the segment between `/d/` and `/edit`) and download it the same way using
`download_file_content` with the XLSX export MIME type.

If no file is found in Drive and no URL is provided, ask the user to upload the DC file
manually.

---

## Step 1 — Fetch fresh RS data from Google Drive as XLSX

**Do this first, before anything else, on every single run.**

**⚠️ Do NOT use `read_file_content` for the RS sheet.** That tool truncates large
spreadsheets and only returns data through the end of 2023 — missing all 2024–2026
rows and causing false "MISSING FROM RS" flags for every recent title.

Instead, download the full RS spreadsheet as XLSX:

```
Tool: mcp__Google_Drive__download_file_content
fileId: 1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ
exportMimeType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
```

The tool returns a base64-encoded binary. Decode it and write the bytes to the outputs
folder as `rs_newreleases.xlsx`:

```bash
python3 -c "
import base64, sys
data = '''<paste base64 string here>'''
with open('/path/to/outputs/rs_newreleases.xlsx', 'wb') as f:
    f.write(base64.b64decode(data))
"
```

The reconciliation script reads the **"New Releases (2026-)"** sheet tab directly from the
XLSX using openpyxl — no CSV conversion needed. This provides complete data with no truncation.

**RS "New Releases" tab column layout (0-indexed):**

| Index | Field |
|-------|-------|
| 0 | MVPD check (go_live flag) |
| 1 | Vendor ID |
| 2 | Upcoming Releases (title) |
| 3 | Type (Film / TV / Bundle / 4K) |
| 4 | Release type — **TRUE = Premium, FALSE = Standard** |
| 5 | Pre-Order (Home Ent date) |
| 6 | 4K Release flag |
| 7 | EST Launch Date (date object) |
| 8 | EST Launch Covered |
| 9 | VOD Launch Date (date object or blank) |

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
  "/sessions/.../mnt/outputs/rs_newreleases.xlsx"
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
| PEST, PVOD | RS EST Launch Date, VOD Launch Date | `Premium` only (Release type = TRUE) |
| EST, VOD | RS EST Launch Date, VOD Launch Date | `Standard` only (Release type = FALSE) |

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
- Sheet tab: "New Releases (2026-)" (gid=114630171)
- The DC is provided by the user each week (uploaded XLSX, Drive file, or Google Sheets URL).
- `wbd_reconcile.py` is bundled in `scripts/` within this skill directory.
- **Never use `read_file_content` for the RS** — it truncates the sheet at ~2023 data.
  Always use `download_file_content` with `exportMimeType: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
