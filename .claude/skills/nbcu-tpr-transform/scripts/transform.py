#!/usr/bin/env python3
"""
NBCU TPR Transform Script
=========================
Transforms the text representation of an NBCU TPR source spreadsheet into
the Looper import format XLSX.

Usage
-----
    python transform.py <src_text_path> <ref_text_path> <output_path>

Arguments
---------
src_text_path   Path to a text file containing the fileContent string from
                read_file_content on the NBCU source XLSX.
ref_text_path   Path to a text file containing the fileContent string from
                read_file_content on the reference workbook
                (NBCU Price Promotions ChatGPT.xlsx, Drive ID:
                1OiBQF5Vza0sohXDZ7Cf0otRr6JOrKImF).
output_path     Where to write the output .xlsx file.

Output
------
Prints a JSON summary to stdout:
    {"output_rows": N, "unknown_mpms": [...], "unknown_platforms": [...], "skipped_rows": N}

Exits with code 0 on success, non-zero on failure.
"""

import sys
import re
import json
import openpyxl
from datetime import datetime


# ---------------------------------------------------------------------------
# Platform rename map — case-insensitive match on the PARTNER column
# ---------------------------------------------------------------------------
RENAMES = {
    'amazon':         'Amazon',
    'amazon ca':      'Amazon',
    'vudu':           'Fandango at Home',
    'google play us': 'YouTube',
    'google play ca': 'YouTube',
    'cineplex':       'CosmoGO',
    'itunes us':      'Apple TV',
    'itunes canada':  'Apple TV',
}

# ---------------------------------------------------------------------------
# Regex for a single data row in the source text representation.
#
# The text representation from read_file_content joins Excel rows with spaces
# and separates fields with commas (using standard CSV quoting for fields
# that contain commas).  Each row has 12 fields:
#   PARTNER, START DATE, END DATE, TERRITORY, PRODUCT ID, PRODUCT NAME,
#   US SD SRP, US HD SRP, US UHD SRP, CA SD SRP, CA HD SRP, CA UHD SRP
# ---------------------------------------------------------------------------
ROW_PATTERN = re.compile(
    r'([^,]+),'                       # 1  PARTNER
    r'(\d{1,2}/\d{1,2}/\d{4}),'      # 2  START DATE  (M/D/YYYY)
    r'(\d{1,2}/\d{1,2}/\d{4}),'      # 3  END DATE    (M/D/YYYY)
    r'(US|CA),'                        # 4  TERRITORY
    r'([^,]+),'                        # 5  PRODUCT ID (MPM / Vendor ID)
    r'((?:"[^"]*"|[^,]+)),'           # 6  PRODUCT NAME (optionally quoted)
    r'([\d.]+),([\d.]+),([\d.]+),'    # 7-9  US SD / HD / UHD SRP
    r'([\d.]+),([\d.]+),([\d.]+)'     # 10-12 CA SD / HD / UHD SRP
)

# Output column headers — must match the looper_format sheet exactly
HEADERS = [
    'Platform Name',
    'Territory ISO Code',
    'Promo Start Date',
    'Promo End Date',
    'Format',
    'Title',
    'Retail Price',
    'Promoted Price',
    'Vendor Identifier',
    'MPM',
]

# Tokens that look like MPMs in the text but are actually header/label words
_MPM_SKIP = {'Vendor', 'Title', 'USD', 'CAD', 'SRP', 'TPR', 'EST', 'UHD',
             'SD', 'HD', 'US', 'CA', 'START', 'END', 'DATE', 'PARTNER',
             'PRODUCT', 'NAME', 'ID', 'TERRITORY', 'FORMAT'}


def _parse_date(date_str: str) -> str:
    """Convert M/D/YYYY → YYYY-MM-DD."""
    return datetime.strptime(date_str.strip(), '%m/%d/%Y').strftime('%Y-%m-%d')


def _load_ref_mpms(ref_text: str) -> set:
    """
    Extract the set of valid MPMs from the 'title list' section of the
    reference workbook text.  Returns an empty set if the section is not
    found (non-fatal — unknown MPM check is skipped).
    """
    marker = 'title list '
    idx = ref_text.find(marker)
    if idx == -1:
        return set()
    title_section = ref_text[idx + len(marker):]
    # MPMs appear as the first comma-separated token on each row, e.g.:
    #   " ABC123,Some Title Name"
    mpm_pat = re.compile(r'(?:^| )([A-Z0-9]{3,8}),')
    return {
        m.group(1)
        for m in mpm_pat.finditer(title_section)
        if m.group(1) not in _MPM_SKIP
    }


def transform(src_text: str, ref_text: str, output_path: str) -> dict:
    """
    Core transformation.  Returns a summary dict.
    """
    # Locate the start of the data rows by finding the last column header
    # in the source header row.  Everything after this marker is data.
    HEADER_MARKER = 'CA EST UHD SRP TPR'
    idx = src_text.find(HEADER_MARKER)
    if idx == -1:
        sys.exit(
            "ERROR: Could not find header marker 'CA EST UHD SRP TPR' in "
            "source text.  Check that the correct source file was loaded."
        )
    data_text = src_text[idx + len(HEADER_MARKER):]

    ref_mpms = _load_ref_mpms(ref_text)

    rows = []
    unknown_mpms: dict = {}  # mpm -> title
    unknown_platforms: set = set()
    skipped_rows = 0

    for m in ROW_PATTERN.finditer(data_text):
        partner     = m.group(1).strip()
        start_date  = _parse_date(m.group(2))
        end_date    = _parse_date(m.group(3))
        territory   = m.group(4).strip()
        mpm         = m.group(5).strip()
        title       = m.group(6).strip().strip('"')
        prices      = [m.group(i) for i in range(7, 13)]   # 6 prices
        formats     = ['SD', 'HD', '4K', 'SD', 'HD', '4K']

        # Resolve platform name
        partner_key = partner.lower()
        if partner_key in RENAMES:
            platform = RENAMES[partner_key]
        else:
            platform = partner
            unknown_platforms.add(partner)

        # Flag MPMs not in the reference list
        if ref_mpms and mpm not in ref_mpms:
            unknown_mpms[mpm] = title

        # Emit one output row per non-zero price
        row_emitted = False
        for price_str, fmt in zip(prices, formats):
            if float(price_str) == 0.0:
                continue
            row_emitted = True
            rows.append([
                platform,                   # Platform Name
                territory,                  # Territory ISO Code
                start_date,                 # Promo Start Date
                end_date,                   # Promo End Date
                fmt,                        # Format
                title,                      # Title
                f'{float(price_str):.2f}',  # Retail Price (text string)
                '-',                        # Promoted Price
                mpm,                        # Vendor Identifier
                mpm,                        # MPM
            ])

        if not row_emitted:
            skipped_rows += 1

    # ------------------------------------------------------------------
    # Write XLSX
    # ------------------------------------------------------------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'looper_format'
    ws.append(HEADERS)

    for row in rows:
        ws.append(row)

    # Apply text number format to Retail Price (col 7) and MPM (col 10)
    # so Excel doesn't reinterpret them as numbers.
    for row_cells in ws.iter_rows(min_row=2):
        row_cells[6].number_format = '@'   # Retail Price  (0-indexed → col 7)
        row_cells[9].number_format = '@'   # MPM           (0-indexed → col 10)

    wb.save(output_path)

    # Write unknown MPMs CSV if any
    csv_path = None
    if unknown_mpms:
        import csv
        csv_path = output_path.replace('.xlsx', '_unknown_mpms.csv')
        with open(csv_path, 'w', newline='') as cf:
            writer = csv.writer(cf)
            writer.writerow(['MPM', 'Title'])
            for mpm, title in sorted(unknown_mpms.items()):
                writer.writerow([mpm, title])

    return {
        'output_rows':       len(rows),
        'unknown_mpms':      sorted(unknown_mpms.keys()),
        'unknown_mpms_csv':  csv_path,
        'unknown_platforms': sorted(unknown_platforms),
        'skipped_rows':      skipped_rows,
    }


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == '__main__':
    if len(sys.argv) != 4:
        print(
            f'Usage: {sys.argv[0]} <src_text_path> <ref_text_path> <output_path>',
            file=sys.stderr,
        )
        sys.exit(1)

    src_text_path, ref_text_path, output_path = sys.argv[1], sys.argv[2], sys.argv[3]

    with open(src_text_path, 'r', encoding='utf-8') as f:
        src_text = f.read()
    with open(ref_text_path, 'r', encoding='utf-8') as f:
        ref_text = f.read()

    summary = transform(src_text, ref_text, output_path)
    print(json.dumps(summary, indent=2))
