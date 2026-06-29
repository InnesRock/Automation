---
name: wbd-reconciliation
description: >
  Runs the weekly WBD Digital Calendar vs Release Schedule reconciliation. Use this skill
  immediately whenever the user uploads a file with "Digital Calendar" or "WBD" in the name,
  or asks to "check the DC", "run the reconciliation", "compare against the RS", "DC check",
  or "reconcile the digital calendar". The skill writes fresh RS data from sync_sources at
  the start of every run — never uses cached data — then runs the Python comparison script
  and reports mismatches. Do not ask clarifying questions; start the workflow immediately.
---

# WBD Digital Calendar Reconciliation

Reconciles a weekly WBD Digital Calendar (DC) xlsx file against the live Warner Bros.
Release Schedule (RS) from Google Sheets, flagging date mismatches and missing titles.

---

## Step 1 — Write fresh RS data from sync_sources

**Do this first, before anything else, on every single run. Never reuse a cached `rs_newreleases.csv`.**

The RS Google Sheet is connected as a sync source and its content is injected into the
session context at startup. Find that data in the context — it will be a block of CSV-like
rows from the "New Releases" sheet.

Write ALL of those rows verbatim to the outputs folder as `rs_newreleases.csv` using the
Write tool (Mac-side path). Include a header row:

```
go_live,vendor_id,title,type,release_type,col5,col6,est_launch,col8,vod_launch
```

**Critical:** write every row you can see in the sync_sources data — do not cherry-pick or
manually select a subset. Missing even one row will cause false "MISSING FROM RS" flags.
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
