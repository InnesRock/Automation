"""
WBD Digital Calendar vs Release Schedule Reconciliation
========================================================
Usage:
    python3 wbd_reconcile.py <digital_calendar.xlsx> <release_schedule.csv>

The release_schedule.csv is exported from the "New Releases" tab of the RS
Google Sheet. At session start Claude writes this from the sync_sources data.

RS CSV column order (0-indexed):
    0: go_live_flag
    1: vendor_id (optional)
    2: title
    3: type (Film/TV/Bundle/4K)
    4: release_type  (Premium | Premium Reprice | Standard | Pre-Order | 4K Release)
    5: col5
    6: col6
    7: EST Launch Date  (DD/MM/YYYY)
    8: col8
    9: VOD Launch Date  (DD/MM/YYYY or blank)
    ...

DC column order (0-indexed):
    0: title
    2: PEST  (MM/DD/YYYY text)
    4: PVOD  (MM/DD/YYYY text)
    6: EST   (MM/DD/YYYY text)
    8: VOD   (MM/DD/YYYY text)

Rules:
    - PEST/PVOD in DC must match each other
    - PEST/PVOD in DC -> RS Premium EST/VOD
    - EST/VOD  in DC -> RS Standard EST/VOD
    - Only "Premium" rows (not "Premium Reprice") match PEST/PVOD
    - Only "Standard" rows match EST/VOD
    - Flag DC entries missing from RS entirely
"""

import sys
import csv
import openpyxl
from datetime import datetime, date
import re
from collections import defaultdict


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
    """RS CSV dates are DD/MM/YYYY."""
    if not v or not str(v).strip():
        return None
    try:
        return datetime.strptime(str(v).strip(), '%d/%m/%Y').date()
    except ValueError:
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
# Load RS from CSV
# ---------------------------------------------------------------------------

def load_rs(csv_path):
    """
    Returns dict: norm_title -> {release_type -> list of {row, est, vod, title}}
    Only includes rows with release_type in: Premium, Standard
    (Pre-Order and Premium Reprice are ignored for date-matching purposes)
    """
    rs = defaultdict(lambda: defaultdict(list))

    with open(csv_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        for i, row in enumerate(reader):
            if len(row) < 10:
                continue
            title = row[2].strip()
            rtype = row[4].strip()
            if not title or rtype not in ('Premium', 'Standard'):
                continue
            est = parse_rs_date(row[7])
            vod = parse_rs_date(row[9])
            entry = {'row': i + 1, 'est': est, 'vod': vod, 'title': title}
            for key in set([normalize(title), normalize_loose(title)]):
                rs[key][rtype].append(entry)

    return rs


# ---------------------------------------------------------------------------
# Load DC from XLSX
# ---------------------------------------------------------------------------

def load_dc(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active
    entries = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        title = row[0]
        if not title or not str(title).strip():
            continue
        pest = parse_dc_date(row[2])
        pvod = parse_dc_date(row[4])
        est  = parse_dc_date(row[6])
        vod  = parse_dc_date(row[8])
        if not any([pest, pvod, est, vod]):
            continue
        entries.append({
            'title': str(title).strip(),
            'pest': pest, 'pvod': pvod,
            'est': est,  'vod': vod,
        })
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

        # Find RS match
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
                # If multiple Standard rows, prefer the one whose EST matches DC
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
        print("Usage: python3 wbd_reconcile.py <digital_calendar.xlsx> <release_schedule.csv>")
        sys.exit(1)

    dc_path = sys.argv[1]
    rs_path = sys.argv[2]

    print(f"Loading DC:  {dc_path}")
    print(f"Loading RS:  {rs_path}")

    dc_entries = load_dc(dc_path)
    rs = load_rs(rs_path)

    print(f"\nDC entries: {len(dc_entries)}")
    print(f"RS Premium+Standard rows loaded: {sum(len(v['Premium']) + len(v['Standard']) for v in rs.values()) // 2}")

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
