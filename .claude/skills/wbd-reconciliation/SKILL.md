---
name: wbd-reconciliation
description: >
  Runs the weekly WBD Digital Calendar vs Release Schedule reconciliation. Use this skill
  immediately whenever the user uploads a file with "Digital Calendar" or "WBD" in the name,
  or asks to "check the DC", "run the reconciliation", "compare against the RS", "DC check",
  or "reconcile the digital calendar". The skill fetches fresh RS data live from Google Drive
  at the start of every run — never uses cached data — then runs the Python comparison script
  and reports mismatches. Do not ask clarifying questions; start the workflow immediately.
---

# WBD Digital Calendar Reconciliation

Reconciles a weekly WBD Digital Calendar (DC) xlsx file against the live Warner Bros.
Release Schedule (RS) from Google Sheets, flagging date mismatches and missing titles.

---

## Step 1 — Fetch fresh RS data from Google Drive

**Do this first, before anything else, on every single run. Never reuse a cached RS file.**

The RS is the "New Releases" tab of the Warner Bros. Release Schedule & Title List
Google Sheet:

- File ID: `1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ`
- Sheet tab: `New Releases`

Download it fresh via the Google Drive connector, exported as xlsx, and save it to the
outputs folder as `rs.xlsx`. The comparison script reads the `New Releases` sheet directly,
so no CSV conversion is needed.

Use the Google Drive `download_file_content` tool with:

- `fileId`: `1ParrlYViu0ii8lP1xNe_TYrvR2obLK0ljyCvB9LGqLQ`
- `exportMimeType`: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`

The connector returns the file as base64. Decode it to `rs.xlsx`. The base64 may be wrapped
in a JSON envelope (e.g. an outer list whose first item has a `text` field containing JSON
with a `content` field). Decode defensively — walk in to find the base64 string, then write
the bytes:

```python
import json, base64, re

raw_tool_output = ...  # the raw string/JSON returned by download_file_content

def find_b64(obj):
    """Return the first plausible base64 xlsx payload found in a nested structure."""
    if isinstance(obj, str):
        s = obj.strip()
        # Try to parse nested JSON first
        if s[:1] in '[{':
            try:
                return find_b64(json.loads(s))
            except Exception:
                pass
        # A raw base64 xlsx starts with 'UEsD' (PK zip header)
        if s.startswith('UEsD'):
            return s
        return None
    if isinstance(obj, dict):
        for k in ('content', 'text', 'data'):
            if k in obj:
                hit = find_b64(obj[k])
                if hit:
                    return hit
        for v in obj.values():
            hit = find_b64(v)
            if hit:
                return hit
    if isinstance(obj, list):
        for v in obj:
            hit = find_b64(v)
            if hit:
                return hit
    return None

b64 = find_b64(raw_tool_output)
with open('/mnt/user-data/outputs/rs.xlsx', 'wb') as f:
    f.write(base64.b64decode(b64))
```

**Do NOT use `read_file_content` for the RS.** It returns a flattened natural-language
rendering that does not preserve the exact column layout the script relies on. Always use
the xlsx export.

If the Drive connector is unavailable, stop and tell the user — do not fall back to a stale
local copy.

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
  "/sessions/.../mnt/outputs/rs.xlsx"
```

Use the actual session paths from the environment. The RS argument may be the `rs.xlsx`
(preferred) or a `.csv` in the same column layout (see below) for backward compatibility.

---

## Step 4 — Report results

- **No issues:** "✅ All X entries pass — no mismatches found."
- **Issues found:** List each one with the title, what mismatched, and DC vs RS values.

Keep it concise — no need to list every passing title unless the user asks.

Because the RS now carries a **single Launch Date** per release-type row (see below), a VOD
mismatch where the EST date matches usually means the DC's VOD window differs from the RS
date, OR the RS simply isn't tracking a separate VOD window. Flag it for a human to confirm
which side is authoritative rather than assuming the DC is wrong.

---

## Reconciliation rules

| DC field | Checked against | RS release_type |
|----------|-----------------|-----------------|
| PEST, PVOD | RS Premium Launch Date | `Premium` only (not "Premium Reprice") |
| EST, VOD | RS Standard Launch Date | `Standard` only |

The RS holds one Launch Date per release-type row. Both DC premium dates (PEST, PVOD) are
checked against the single Premium Launch Date; both DC standard dates (EST, VOD) against the
single Standard Launch Date.

Additional flags:
- PEST ≠ PVOD in DC → flag internal DC mismatch.
- DC has PEST/PVOD but RS has no Premium row → "NO PREMIUM ROW IN RS".
- DC has EST/VOD but RS has no Standard row → "NO STANDARD ROW IN RS".
- DC title not found in RS at all → "MISSING FROM RS".
- Multiple Standard rows: prefer the one whose Launch Date matches DC EST; otherwise use the
  highest row number.
- Title matching: case-insensitive, strips articles (the/a/an), punctuation, year suffixes
  like "(2026)", and annotations like "(Ani)".

---

## RS "New Releases" column layout (0-indexed)

| Index | Field |
|-------|-------|
| 0 | MVPD check (go_live flag) |
| 1 | Vendor ID |
| 2 | Upcoming Releases (title) |
| 3 | Type (Film / TV / Film Bundle / TV Boxset ...) |
| 4 | Release type (Premium / Premium Reprice / Standard / Pre-Order / 4K Release) |
| 5 | **Launch Date** (single date per row) |
| 6 | Launch Covered |
| 7 | Avail on all platforms? |
| 8 | Added to Title List |
| 9 | Day of Week |

> **Note:** Earlier versions of this sheet had separate EST (col 7) and VOD (col 9) date
> columns. It no longer does — there is now one "Launch Date" at col 5. The script reads the
> date from col 5 and uses it as both the EST and VOD reference. If the sheet layout changes
> again, update `RS_COL_LAUNCH` (and the other `RS_COL_*` constants) at the top of
> `wbd_reconcile.py`.

If a `.csv` is passed instead of the xlsx, it must follow this same column order (Launch Date
at index 5).

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
