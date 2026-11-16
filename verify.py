"""Compare the Python output with the V5.1 workbook built from the same folder.

Usage: python3 -I verify.py NEW.xlsx V51.xlsm
Formulas in NEW.xlsx are recalculated with LibreOffice first.
"""
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import load_workbook


def recalc(path):
    """Recalculate with LibreOffice. LibreOffice treats =TOI!C6 (empty source)
    as empty text in comparisons, Excel treats it as 0. N() restores Excel's
    behaviour for this check only; the delivered workbook is not changed."""
    wb = load_workbook(path)
    if "ADT" in wb.sheetnames:
        for row in wb["ADT"].iter_rows(min_col=5, max_col=6):
            for c in row:
                if isinstance(c.value, str) and re.fullmatch(r"=TOI![CD]\d+", c.value):
                    c.value = "=N(" + c.value[1:] + ")"
    tmpx = Path(tempfile.mkdtemp()) / Path(path).name
    wb.save(tmpx)
    path = tmpx
    out = Path(tempfile.mkdtemp())
    subprocess.run(["soffice", "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                   check=True, capture_output=True, timeout=180)
    return out / Path(path).name


def norm(v):
    if v in (None, ""):
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return round(float(v), 6)
    if isinstance(v, str):
        s = v.replace("​", "").strip()
        try:
            return round(float(s), 6)
        except ValueError:
            return s
    return v


PAIRS = [  # (new sheet, V5.1 sheet, ignore cells)
    ("Summary MTR", "Summary MTR"),
    ("TOI", "TOI"),
    ("FULL TOI", "FULL TOI"),
    ("MTR_PR", "MTR_PR"), ("MTR_SR", "MTR_SR"), ("MTR_TOS", "MTR_TOS"), ("MTR_WHT", "MTR_WHT"),
    ("Tax Paid Summary", "Tax Paid Summary"),
    ("Company Info", "TPS_CompanyInfo"),
    ("PR", "PR"), ("SR", "SR"), ("TOS", "TOS"), ("WHT", "WHT"),
    ("Audit Dashboard", "Audit_Dashboard"),
]


def compare(new_ws, old_ws, name):
    diffs, cells = [], 0
    rows = max(new_ws.max_row, old_ws.max_row)
    cols = max(new_ws.max_column, old_ws.max_column)
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            a, b = norm(new_ws.cell(r, c).value), norm(old_ws.cell(r, c).value)
            if a is None and b is None:
                continue
            cells += 1
            if name in ("PR", "SR", "TOS", "WHT") and (r, c) == (1, 1):
                continue  # "Consolidated on: <timestamp>"
            if name == "Audit Dashboard" and r == 8:
                continue  # report date
            if a != b:
                diffs.append((new_ws.cell(r, c).coordinate, a, b))
    return cells, diffs


def main(new, old):
    new_calc = recalc(new)
    n = load_workbook(new_calc, data_only=True)
    o = load_workbook(old, data_only=True)
    total_diff = 0
    for a, b in PAIRS:
        if a not in n.sheetnames or b not in o.sheetnames:
            print(f"{a:18s} MISSING (new={a in n.sheetnames}, old={b in o.sheetnames})")
            continue
        cells, diffs = compare(n[a], o[b], a)
        total_diff += len(diffs)
        print(f"{a:18s} {cells:6d} cells compared, {len(diffs):4d} differences")
        for d in diffs[:6]:
            print("      ", d)
    print("TOTAL differences:", total_diff)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
