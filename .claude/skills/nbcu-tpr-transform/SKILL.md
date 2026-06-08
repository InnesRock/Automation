---
name: nbcu-tpr-transform
description: >
  Fetches the latest NBCU TPR (Temporary Price Reduction) pricing file from
  Google Drive, transforms it into the Looper import format, and saves a
  formatted XLSX output. Use this skill whenever processing NBCU TPR pricing
  files, running the NBCU pricing import pipeline, or when asked to transform
  an NBCU spreadsheet from the "NBCU TPR Inbox" Drive folder. Also triggers
  on "NBCU TPR", "NBCU pricing promotions", "run the NBCU transform", or any
  task involving the looper_format XLSX for NBCU pricing data.
---

# NBCU TPR Transform

Fetch the latest NBCU TPR file from Google Drive, transform it into the
Looper import format, and save the output XLSX.

## Connectors required

- **Google Drive MCP** (`mcp__fb1640e6__*`) — for file search and content
- **Gmail MCP** (`mcp__e02e688d__*`) — fallback source discovery

## Step 1 — Find the source file

Search the NBCU TPR Inbox Drive folder for the most recent file:

```
Tool: mcp__fb1640e6__search_files
folderId: 1NhHxH8fw_E30QD_4AwD825fVQg2cG6a6
```

Take the file with the latest `modifiedTime`. Note its `id` and `name`.

**Fallback (if no file found):** Search Gmail with:
```
from:HEDWMSTR.SUPPORT@nbcuni.com subject:"Looper TPR" has:attachment newer_than:3d
```
Then locate the corresponding file in Drive by name.

## Step 2 — Fetch both files as text

> **Critical:** Use `read_file_content`, NOT `download_file_content`. The
> binary download truncates large files and produces a corrupt ZIP. The text
> representation from `read_file_content` is what the transform script parses.

```
Tool: mcp__fb1640e6__read_file_content
fileId: <source file ID from Step 1>
```

```
Tool: mcp__fb1640e6__read_file_content
fileId: 1OiBQF5Vza0sohXDZ7Cf0otRr6JOrKImF   ← reference workbook (fixed)
```

The reference workbook is "NBCU Price Promotions ChatGPT.xlsx". It contains:
- `looper_format` sheet — example output rows (ignore for transformation)
- `title list` sheet — Vendor Identifier + Title pairs used for MPM validation

## Step 3 — Write text to temp files

Extract the `fileContent` string from each tool result and write it to disk
so the transform script can read it:

```
Write /tmp/nbcu_src.txt   ← fileContent from source file
Write /tmp/nbcu_ref.txt   ← fileContent from reference file
```

## Step 4 — Run the transform script

The bundled script is at `scripts/transform.py` (relative to this SKILL.md).
Determine today's date for the output filename, then run:

```bash
pip install openpyxl --break-system-packages -q
python <skill_dir>/scripts/transform.py \
    /tmp/nbcu_src.txt \
    /tmp/nbcu_ref.txt \
    "<outputs_dir>/NBCU Price Promotions ChatGPT <D><MON><YYYY>.xlsx"
```

Where `<D><MON><YYYY>` is today's date with no leading zero on the day,
uppercase 3-letter month, and 4-digit year — e.g. `5JUN2026`, `14AUG2026`.

The script prints a JSON summary to stdout:
```json
{
  "output_rows": 1981,
  "unknown_mpms": ["05603"],
  "unknown_platforms": [],
  "skipped_rows": 0
}
```

## Step 5 — Post to Slack

Run:
pip install slack-sdk --break-system-packages -q
python <skill_dir>/scripts/post_to_slack.py "<output_path>" "✅ NBCU TPR Transform complete — <output_rows> rows." "<unknown_mpms_csv from summary, or empty string if none>"

## Platform rename map (for reference)

The transform script handles these internally, but they are documented here
for transparency. Matching is case-insensitive on the PARTNER field:

| Source PARTNER  | Output Platform Name |
|-----------------|----------------------|
| amazon          | Amazon               |
| amazon ca       | Amazon               |
| vudu            | Fandango at Home     |
| google play us  | YouTube              |
| google play ca  | YouTube              |
| cineplex        | CosmoGO              |
| itunes us       | Apple TV             |
| itunes canada   | Apple TV             |

Any PARTNER not in this map is passed through unchanged and flagged as unknown.

## Output format

One row per non-zero price point. Columns (in order):

| # | Column Name        | Notes                              |
|---|--------------------|------------------------------------|
| 1 | Platform Name      | Renamed partner                    |
| 2 | Territory ISO Code | US or CA                           |
| 3 | Promo Start Date   | ISO 8601 (YYYY-MM-DD)              |
| 4 | Promo End Date     | ISO 8601 (YYYY-MM-DD)              |
| 5 | Format             | SD, HD, or 4K                      |
| 6 | Title              | Product name, unquoted             |
| 7 | Retail Price       | "X.XX" string, text-formatted cell |
| 8 | Promoted Price     | Always `-` (literal hyphen)        |
| 9 | Vendor Identifier  | MPM / Product ID                   |
|10 | MPM                | Same as Vendor Identifier, text-formatted cell |

Sheet name: `looper_format`
