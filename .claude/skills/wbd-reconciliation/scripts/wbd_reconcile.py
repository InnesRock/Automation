"""
WBD Digital Calendar vs Release Schedule Reconciliation
========================================================
Usage:
    python3 wbd_reconcile.py <digital_calendar.xlsx> <release_schedule.xlsx|csv>

The release_schedule can be either:
  - An XLSX file exported from the RS Google Sheet (preferred — complete data)
  - A CSV file in the legacy format (kept for backwards compatibility)

RS Google Sheet column order (0-indexed, "New Releases" tab):
    0: MVPD check (go_live flag)
    1: Vendor ID
    2: Upcoming Releases (title)
    3: Type (Film/TV/Bundle/4K)
    4: Release type  — TRUE = Premium, FALSE = Standard
                       (legacy text values "Premium"/"Standard" also accepted)
    5: Pre-Order (Home Ent date)
    6: 4K Release
    7: EST Launch Date  (date object from XLSX, or DD/MM/YYYY string from CSV)
    8: EST Launch Covered
    9: VOD Launch Date  (date object from XLSX, DD/MM/YYYY string, or blank)

DC column order (0-indexed, xlsx):
    0: title
    2: PEST  (MM/DD/YYYY text)
    4: PVOD  (MM/DD/YYYY text)
    6: EST   (MM/DD/YYYY text)
    8: VOD   (MM/DD/YYYY text)

Rules:
    - PEST/PVOD in DC must match each other
    - PEST/PVOD in DC -> RS Premium EST/VOD
    - EST/VOD  in DC -> RS Standard EST/VOD
    - Only "Premium" rows (TRUE) match PEST/PVOD
    - Only "Standard" rows (FALSE) match EST/VOD
    - Flag DC entries missing from RS entirely
"""

import sys
import csv
import openpyxl
from datetime import datetime, date, timedelta
import re
from collections import defaultdict

# Only reconcile entries where at least one date is within this window
LOOKBACK_DAYS = 28


# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

def parse_dc_date(v):
    """DC dates are MM/DD/YYYY text strings (or real date/datetime objects)."""
    if isinstance(v, str) and v.strip():
        try:
            return datetime.strptime(v.strip(), '%m/%d/%Y').date()
        except ValueError:
            return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    return None


def parse_rs_date(v):
    """RS dates: date/datetime objects from XLSX, or DD/MM/YYYY strings from CSV."""
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s or s.upper() == 'N/A':
        return None
    try:
        return datetime.strptime(s, '%d/%m/%Y').date()
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Release type mapping
# ---------------------------------------------------------------------------

def map_release_type(raw):
    """
    Maps the RS Release type column to 'Premium' or 'Standard'.
    The column now stores Python booleans (True/False) or the strings
    'TRUE'/'FALSE' from a Google Sheets checkbox. Legacy text values
    'Premium' and 'Standard' are also accepted for backwards compatibility.
    """
    if raw is True:
        return 'Premium'
    if raw is False:
        return 'Standard'
    if isinstance(raw, str):
        s = raw.strip().upper()
        if s == 'TRUE':
            return 'Premium'
        if s == 'FALSE':
            return 'Standard'
        if raw.strip() == 'Premium':
            return 'Premium'
        if raw.strip() == 'Standard':
            return 'Standard'
    return None


# ---------------------------------------------------------------------------
# Title normalisation
# ---------------------------------------------------------------------------

def normalize(s):
    s = str(s).lower().strip()
    s = re.sub(r'[^a-z0-9 ]', '', s)
    s = re.sub(r'\b(the|a|an)\b', '', s)
    return re.sub(r'\s+', ' ', s).strip()


def normalize_loose(s):
    """Also strips year suffixes like (2026) and annotations like (Ani)."""
    s = str(s).lower().strip()
    s = re.sub(r'\(\d{4}\)', '', s)
    s = re.sub(r'\([a-z]+\)', '', s)   # e.g. (Ani)
    s = re.sub(r'[^a-z0-9 ]', '', s)
    s = re.sub(r'\b(the|a|an)\b', '', s)
    return re.sub(r'\s+', ' ', s).strip()


# ---------------------------------------------------------------------------
# Load RS from XLSX (preferred — reads "New Releases" tab directly, no truncation)
# ---------------------------------------------------------------------------

def load_rs_xlsx(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    # Find the "New Releases (2026-)" sheet (or any tab whose name starts with "New Releases")
    ws = None
    for name in wb.sheetnames:
        if name.lower().startswith('new release'):
            ws = wb[name]
            break
    if ws is None:
        ws = wb.active
        print(f"Warning: 'New Releases' sheet not found — using '{ws.title}'", file=sys.stderr)

    rs = defaultdict(lambda: defaultdict(list))
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        if not row or len(row) < 10:
            continue
        title = row[2]
        if not title or not str(title).strip():
            continue
        title = str(title).strip()
        rtype = map_release_type(row[4])
        if rtype not in ('Premium', 'Standard'):
            continue
        est = parse_rs_date(row[7])
        vod = parse_rs_date(row[9])
        entry = {'row': i + 2, 'est': est, 'vod': vod, 'title': title}
        for key in set([normalize(title), normalize_loose(title)]):
            rs[key][rtype].append(entry)

    return rs


# ---------------------------------------------------------------------------
# Load RS from CSV (legacy fallback)
# ---------------------------------------------------------------------------

def load_rs_csv(csv_path):
    rs = defaultdict(lambda: defaultdict(list))

    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for i, row in enumerate(reader):
            if len(row) < 10:
                continue
            title = row[2].strip()
            rtype = map_release_type(row[4].strip())
            if not title or rtype not in ('Premium', 'Standard'):
                continue
            est = parse_rs_date(row[7])
            vod = parse_rs_date(row[9])
            entry = {'row': i + 1, 'est': est, 'vod': vod, 'title': title}
            for key in set([normalize(title), normalize_loose(title)]):
                rs[key][rtype].append(entry)

    return rs


def load_rs(path):
    if path.lower().endswith(('.xlsx', '.xls')):
        return load_rs_xlsx(path)
    return load_rs_csv(path)


# ---------------------------------------------------------------------------
# Load DC from XLSX
# ---------------------------------------------------------------------------

def load_dc(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    cutoff = date.today() - timedelta(days=LOOKBACK_DAYS)
    entries = []
    skipped = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        title = row[0]
        if not title or not str(title).strip():
            continue
        pest = parse_dc_date(row[2])
        pvod = parse_dc_date(row[4])
        est  = parse_dc_date(row[6])
        vod  = parse_dc_date(row[8])
        dates = [d for d in [pest, pvod, est, vod] if d is not None]
        if not dates:
            continue
        # Skip entries where every date is older than the lookback window
        if max(dates) < cutoff:
            skipped += 1
            continue
        entries.append({
            'title': str(title).strip(),
            'pest': pest, 'pvod': pvod,
            'est': est,  'vod': vod,
        })
    if skipped:
        print(f"Skipped {skipped} DC entries with all dates older than {LOOKBACK_DAYS} days (before {cutoff})")
    return entries


# ---------------------------------------------------------------------------
# Reconcile
# ---------------------------------------------------------------------------

def reconcile(dc_entries, rs):
    issues = []
    passed = []

    for dc in dc_entries:
        title = dc['title']
        nk = normalize(title)
        nk_loose = normalize_loose(title)

        rs_data = rs.get(nk) or rs.get(nk_loose)

        if rs_data is None:
            issues.append(f"MISSING FROM RS — {title}")
            continue

        entry_ok = True

        # 1. PEST / PVOD internal consistency
        if dc['pest'] and dc['pvod'] and dc['pest'] != dc['pvod']:
            issues.append(
                f"PEST/PVOD MISMATCH IN DC — {title} | "
                f"PEST={dc['pest']}, PVOD={dc['pvod']}"
            )
            entry_ok = False

        # 2. Premium check (PEST/PVOD vs RS Premium EST/VOD)
        if dc['pest'] or dc['pvod']:
            premium_rows = rs_data.get('Premium', [])
            if not premium_rows:
                issues.append(f"NO PREMIUM ROW IN RS — {title}")
                entry_ok = False
            else:
                pr = premium_rows[0]
                if dc['pest'] and pr['est'] and dc['pest'] != pr['est']:
                    issues.append(
                        f"PEST ≠ RS Premium EST — {title} | "
                        f"DC={dc['pest']}, RS={pr['est']} (row {pr['row']})"
                    )
                    entry_ok = False
                if dc['pvod'] and pr['vod'] and dc['pvod'] != pr['vod']:
                    issues.append(
                        f"PVOD ≠ RS Premium VOD — {title} | "
                        f"DC={dc['pvod']}, RS={pr['vod']} (row {pr['row']})"
                    )
                    entry_ok = False
                if dc['pvod'] and not pr['vod']:
                    issues.append(
                        f"PVOD ≠ RS Premium VOD — {title} | "
                        f"DC={dc['pvod']}, RS VOD=blank (row {pr['row']})"
                    )
                    entry_ok = False

        # 3. Standard check (EST/VOD vs RS Standard EST/VOD)
        if dc['est'] or dc['vod']:
            standard_rows = rs_data.get('Standard', [])
            if not standard_rows:
                issues.append(f"NO STANDARD ROW IN RS — {title}")
                entry_ok = False
            else:
                if len(standard_rows) > 1:
                    exact = [r for r in standard_rows if r['est'] == dc['est']]
                    standard_rows = exact if exact else [
                        max(standard_rows, key=lambda r: r['row'])
                    ]
                sr = standard_rows[0]
                if dc['est'] and sr['est'] and dc['est'] != sr['est']:
                    issues.append(
                        f"EST ≠ RS Standard EST — {title} | "
                        f"DC={dc['est']}, RS={sr['est']} (row {sr['row']})"
                    )
                    entry_ok = False
                if dc['vod'] and sr['vod'] and dc['vod'] != sr['vod']:
                    issues.append(
                        f"VOD ≠ RS Standard VOD — {title} | "
                        f"DC={dc['vod']}, RS={sr['vod']} (row {sr['row']})"
                    )
                    entry_ok = False
                if dc['vod'] and not sr['vod']:
                    issues.append(
                        f"VOD ≠ RS Standard VOD — {title} | "
                        f"DC VOD={dc['vod']}, RS VOD=blank (row {sr['row']})"
                    )
                    entry_ok = False

        if entry_ok:
            passed.append(title)

    return passed, issues


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) != 3:
        print("Usage: python3 wbd_reconcile.py <digital_calendar.xlsx> <release_schedule.xlsx|csv>")
        sys.exit(1)

    dc_path = sys.argv[1]
    rs_path = sys.argv[2]

    print(f"Loading DC:  {dc_path}")
    print(f"Loading RS:  {rs_path}")

    dc_entries = load_dc(dc_path)
    rs = load_rs(rs_path)

    premium_count = sum(len(v.get('Premium', [])) for v in rs.values())
    standard_count = sum(len(v.get('Standard', [])) for v in rs.values())
    print(f"\nDC entries: {len(dc_entries)}")
    print(f"RS Premium rows loaded:  {premium_count // max(1, 1)}")
    print(f"RS Standard rows loaded: {standard_count // max(1, 1)}")

    passed, issues = reconcile(dc_entries, rs)

    print(f"\n{'='*60}")
    print(f"PASS ({len(passed)})")
    print(f"{'='*60}")
    for t in passed:
        print(f"  ✅  {t}")

    print(f"\n{'='*60}")
    print(f"ISSUES ({len(issues)})")
    print(f"{'='*60}")
    if issues:
        for i, iss in enumerate(issues, 1):
            print(f"  {i}. ❌  {iss}")
    else:
        print("  ✅  No issues found!")


if __name__ == '__main__':
    main()
