"""Tax Audit Engine: Python port of the V5.1 Excel/VBA reports.

consolidate.py is used unchanged. Everything the V5.1 VBA modules built
(Summary MTR, TOI, FULL TOI, MTR detail, Company Info, Tax Paid Summary,
Audit Dashboard) is rebuilt here and written to ONE Excel workbook with live
formulas, so the auditor can still review and edit it in Excel.
"""
import re
import sys
from copy import copy
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import consolidate as cons  # noqa: E402  (original V5.1 script, unchanged)

FONT = "Khmer OS Siemreap"
THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
YELLOW = PatternFill("solid", fgColor="FFFFCC")
BLUE = PatternFill("solid", fgColor="BDD7EE")
LBLUE = PatternFill("solid", fgColor="DCE6F1")
SUM_COLS = {"PR": (7, 15), "SR": (8, 19), "TOS": (10, 14), "WHT": (10, 11)}
HEADER_ROWS = 2
ADT_OFFSET = 9  # Module 4 pastes ADT at row 10


class Log(list):
    def add(self, where, msg):
        self.append((datetime.now().strftime("%Y-%m-%d %H:%M:%S"), where, msg))


def ci_sheet(wb, name):
    for ws in wb.worksheets:
        if ws.title.casefold() == name.casefold():
            return ws
    return None


def first_xlsx(folder):
    files = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    return files[0] if files else None


def find_dir(root, name):
    for d in root.iterdir():
        if d.is_dir() and d.name.upper() == name.upper():
            return d
    return None


def fmt_all(ws):
    fonts = {}
    for row in ws.iter_rows():
        for c in row:
            f = c.font
            key = (f.bold, f.italic, f.color)
            if key not in fonts:
                fonts[key] = Font(name=FONT, size=10, bold=f.bold, italic=f.italic, color=f.color)
            c.font = fonts[key]


def last_row_col(ws, col=1):
    last = 0
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, col).value not in (None, ""):
            last = r
            break
    return last


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ---------------------------------------------------------------- Module 5
def build_summary_mtr(wb, log):
    ws = wb.create_sheet("Summary MTR")
    dest = 1
    for name in ("SR", "PR", "TOS", "WHT"):
        src = ci_sheet(wb, name)
        if src is None:
            log.add("Summary MTR", f"Sheet {name} not found, skipped")
            continue
        yr = next((r for r in range(1, src.max_row + 1) if src.cell(r, 1).value == "Year"), None)
        if yr is None:
            log.add("Summary MTR", f"{name}: no 'Year' table found")
            continue
        total_cols = src.max_column
        s, e = SUM_COLS[name]
        ws.cell(dest, 1, f"Folder: {name}").font = Font(bold=True)
        for r in range(HEADER_ROWS):
            for c in range(1, total_cols + 1):
                a, b = src.cell(2 + r, c), ws.cell(dest + 1 + r, c)
                b.value = a.value
                b._style = copy(a._style)
                b.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        for m in src.merged_cells.ranges:
            if 2 <= m.min_row and m.max_row <= 1 + HEADER_ROWS:
                ws.merge_cells(start_row=m.min_row - 2 + dest + 1, start_column=m.min_col,
                               end_row=m.max_row - 2 + dest + 1, end_column=m.max_col)
        for c in range(1, total_cols + 1):
            ws.column_dimensions[L(c)].width = 13.5
        block_start, row, y = dest + 1, dest + 1 + HEADER_ROWS, yr + 1
        while src.cell(y, 1).value not in (None, ""):
            ws.cell(row, 1, src.cell(y, 1).value)
            for c in range(s, e + 1):
                cell = ws.cell(row, c, src.cell(y, c).value)
                cell.number_format = "#,##0"
            row, y = row + 1, y + 1
        for r in range(block_start, row):
            for c in range(1, e + 1):
                ws.cell(r, c).border = BOX
        dest = row + 2
    fmt_all(ws)
    return ws


def build_toi(wb, toi_wb, log):
    ws = wb.create_sheet("TOI")
    cur, blocks = 1, []
    for n in range(4, 8):
        sh = ci_sheet(toi_wb, f"step{n}")
        if sh is None:
            log.add("TOI", f"step{n} missing in TOI file")
            continue
        last = last_row_col(sh, 1)
        if last < 5:
            continue
        first = cur
        for r in range(5, last + 1):
            for c in range(1, 5):
                v = sh.cell(r, c).value
                ws.cell(cur + r - 5, c, None if v == "" else v)
        end = first + (last - 4) - 1
        ws.cell(first, 5, "Compare")
        ws.cell(first, 6, "% Change")
        ws.cell(first, 7, "Remark")
        blocks.append((first, end))
        cur = end + 1
        for c in range(1, 8):
            ws.cell(cur, c).fill = YELLOW
        cur += 1
    for first, end in blocks:
        for r in range(first + 1, end + 1):
            ws.cell(r, 5, f'=IF(AND(C{r}="",D{r}=""),"",IF(C{r}="",0,C{r})-IF(D{r}="",0,D{r}))')
            ws.cell(r, 6, f'=IF(OR(E{r}="",D{r}=0),"",E{r}/D{r})')
            ws.cell(r, 7, f'=IF(E{r}="","",IF(E{r}>0,"Increase",IF(E{r}<0,"Decrease","")))')
    fmt_all(ws)
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=7):
        for c in row:
            c.border = BOX
            if c.column >= 2:
                c.alignment = Alignment(horizontal="center")
            if 3 <= c.column <= 5:
                c.number_format = "#,##0"
            if c.column == 6:
                c.number_format = "0.0%"
    for col, w in zip("ABCDEFG", (60, 10, 20, 20, 18, 11, 11)):
        ws.column_dimensions[col].width = w
    return ws


def build_full_toi(wb, toi_wb, log):
    ws = wb.create_sheet("FULL TOI")
    cur = 1
    widths = {}
    for n in range(1, 18):
        sh = ci_sheet(toi_wb, f"step{n}")
        if sh is None:
            continue
        if not any(c.value not in (None, "") for row in sh.iter_rows() for c in row):
            continue
        rows, cols = sh.max_row, sh.max_column
        first = cur
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                v = sh.cell(r, c).value
                if v not in (None, ""):
                    ws.cell(cur + r - 1, c, v)
        for c in range(5, cols + 1):
            w = sh.column_dimensions[L(c)].width
            if w:
                widths[c] = w
        ws.cell(first, 1).font = Font(bold=True)
        for c in range(1, cols + 1):
            ws.cell(first, c).font = Font(bold=True)
        for r in range(first, first + rows):
            for c in range(1, cols + 1):
                ws.cell(r, c).border = BOX
        cur += rows
        for c in range(1, cols + 1):
            ws.cell(cur, c).fill = YELLOW
        cur += 1
        if first > 1:
            from openpyxl.worksheet.pagebreak import Break
            ws.row_breaks.append(Break(id=first - 1))
    fmt_all(ws)
    for row in ws.iter_rows():
        for c in row:
            c.number_format = "#,##0"
            c.alignment = Alignment(horizontal="left" if c.column == 1 else "center")
    for c, w in widths.items():
        ws.column_dimensions[L(c)].width = w
    for col, w in (("A", 57), ("B", 20), ("C", 30), ("D", 30)):
        ws.column_dimensions[col].width = w
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    return ws


# ---------------------------------------------------------------- Module 6
def build_mtr_detail(wb, root, log):
    idx_ws = ci_sheet(wb, cons.MTR_INDEX)
    index = {}
    if idx_ws is not None:
        for r in range(2, idx_ws.max_row + 1):
            t, row = idx_ws.cell(r, 1).value, idx_ws.cell(r, 2).value
            if t and num(row):
                index.setdefault(t, []).append(int(row))
        del wb[idx_ws.title]
    for tax in ("PR", "SR", "TOS", "WHT"):
        src = ci_sheet(wb, tax)
        if src is None or tax not in index:
            log.add(f"MTR_{tax}", "Consolidation sheet or file subtotal rows not found")
            continue
        s, e = SUM_COLS[tax]
        n = e - s + 1
        ws = wb.create_sheet(f"MTR_{tax}")
        ws.cell(1, 1, f"{tax} - FILE LIST").font = Font(bold=True, size=12)
        ws.cell(2, 1, "N")
        ws.cell(2, 2, "File")
        for c in range(s, e + 1):
            ws.cell(2, c - s + 3, src.cell(2, c).value)
            ws.cell(3, c - s + 3, src.cell(3, c).value)
        for r in (2, 3):
            for c in range(1, n + 3):
                x = ws.cell(r, c)
                x.font, x.fill, x.border = Font(bold=True), BLUE, BOX
                x.alignment = Alignment(horizontal="center")
        by_year = {}
        for r in index[tax]:
            y = str(src.cell(r, 2).value).strip() if src.cell(r, 2).value not in (None, "") else "Unknown"
            by_year.setdefault(y, []).append(r)
        out, k, sub_rows = 4, 0, {}
        for y in sorted(by_year, key=str.casefold):
            ws.cell(out, 1, f"YEAR: {y}").font = Font(bold=True)
            out += 1
            start = out
            for r in by_year[y]:
                k += 1
                fname = str(src.cell(r, 1).value)
                ws.cell(out, 1, k)
                cell = ws.cell(out, 2, fname)
                path = root / tax / y / fname
                if y != "Unknown" and path.exists():
                    cell.hyperlink = str(path)
                for c in range(n):
                    ws.cell(out, 3 + c, src.cell(r, s + c).value)
                out += 1
            ws.cell(out, 1, f"Subtotal {y}")
            for c in range(3, n + 3):
                ws.cell(out, c, f"=SUM({L(c)}{start}:{L(c)}{out - 1})")
            for c in range(1, n + 3):
                ws.cell(out, c).font, ws.cell(out, c).fill = Font(bold=True), LBLUE
            sub_rows[y] = out
            out += 2
        ws.cell(out, 1, "GRAND TOTAL")
        for c in range(3, n + 3):
            ws.cell(out, c, "=" + "+".join(f"{L(c)}{sub_rows[y]}" for y in sorted(by_year, key=str.casefold)))
        for c in range(1, n + 3):
            x = ws.cell(out, c)
            x.font, x.fill, x.border = Font(bold=True), BLUE, BOX
        fmt_all(ws)
        for r in range(4, out + 1):
            for c in range(3, n + 3):
                ws.cell(r, c).number_format = "#,##0.00"
        ws.column_dimensions["A"].width = 14
        ws.column_dimensions["B"].width = 48
        for c in range(3, n + 3):
            ws.column_dimensions[L(c)].width = 18


# ---------------------------------------------------------------- Module 3
def company_info(wb, info_wb, log):
    ws = wb.create_sheet("Company Info")
    s1 = ci_sheet(info_wb, "Sheet1")
    if s1 is None:
        log.add("Company Info", "Sheet1 not found in Company Info file")
        return ws
    ws["A1"] = "TAX PAID SUMMARY"
    for out, src in enumerate((1, 2, 5, 6, 3), start=3):
        ws.cell(out, 1, s1.cell(src, 1).value)
        ws.cell(out, 2, s1.cell(src, 2).value)
    ws.merge_cells("A1:E1")
    ws["A1"].font = Font(bold=True, size=14)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    for r in range(3, 8):
        ws.cell(r, 1).font = Font(bold=True)
    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 60
    return ws


def info_year(sheet):
    """Require the year recorded by the Info template instead of guessing it."""
    value = sheet["B3"].value
    if num(value) and value == int(value) and 1900 <= value <= 2100:
        return int(value)
    match = re.fullmatch(r"(?:Year\s*)?(19\d{2}|20\d{2}|2100)", str(value).strip(), re.I)
    if match:
        return int(match.group(1))
    raise ValueError(f"Info {sheet.title}!B3 must contain its year (for example 2026).")


def tax_paid_summary(wb, info_wb, months_prior, log):
    if not 1 <= months_prior <= 12:
        raise ValueError("Months of newest year must be between 1 and 12.")
    s1, s2, s3 = (ci_sheet(info_wb, n) for n in ("Sheet1", "Sheet2", "Sheet3"))
    if s1 is None or s2 is None or s3 is None:
        log.add("Tax Paid Summary", "Sheet1/Sheet2/Sheet3 missing in Company Info file")
        return None
    year2, year3 = info_year(s2), info_year(s3)
    if year2 == year3:
        raise ValueError("Info Sheet2 and Sheet3 must describe different years.")
    if year2 > year3:
        s2, s3, year2, year3 = s3, s2, year3, year2
    y_a, y_b = year2, year3
    end_b = 3 + months_prior - 1
    period = f"{y_b} / {months_prior} months"
    ws = wb.create_sheet("_Tax Paid Summary")
    ws["A1"] = "TAX PAID SUMMARY"
    info_row = 3
    for src in (1, 2, 5, 6, 3):
        ws.cell(info_row, 1, s1.cell(src, 1).value)
        ws.cell(info_row, 2, s1.cell(src, 2).value)
        info_row += 1
    hdr = info_row + 1
    for c, v in enumerate(("N", "Tax item", f"Year {y_a}", period, "Sub Total"), start=1):
        ws.cell(hdr, c, v)
    # Keep first-match semantics for duplicate item names, while indexing once.
    newest_items = {}
    for values in s3.iter_rows(min_row=2, max_row=last_row_col(s3, 1), max_col=14, values_only=True):
        name = str(values[1] or "").strip()
        newest_items.setdefault(name, sum(v for v in values[2:end_b] if num(v)))
    out, n = hdr + 1, 1
    for values in s2.iter_rows(min_row=2, max_row=last_row_col(s2, 1), max_col=14, values_only=True):
        name = str(values[1] or "").strip()
        if not name or name.upper() == "TYPE OF TAX" or "???" in name:
            continue
        if len(name) <= 4 and name.isascii() and name.isdigit() and 1900 <= int(name) <= 2100:
            continue
        v_a = sum(v for v in values[2:14] if num(v))
        v_b = newest_items.get(name, 0)
        if v_a or v_b:
            for c, v in enumerate((n, name, v_a, v_b, v_a + v_b), start=1):
                ws.cell(out, c, v)
            n, out = n + 1, out + 1
    first, last = hdr + 1, out - 1
    ws.cell(out, 2, "Total")
    for c in "CDE":
        ws[f"{c}{out}"] = f"=SUM({c}{first}:{c}{last})"
    fmt_all(ws)
    for c in range(1, 6):
        for r in (hdr, out):
            x = ws.cell(r, c)
            x.font, x.fill = Font(name=FONT, size=10, bold=True), BLUE
            x.alignment = Alignment(horizontal="center", vertical="center")
        for r in range(hdr, out + 1):
            ws.cell(r, c).border = BOX
            if c >= 3:
                ws.cell(r, c).number_format = "#,##0"
    ws.merge_cells("A1:E1")
    ws["A1"].font = Font(name=FONT, size=14, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    for r in range(3, info_row):
        ws.cell(r, 1).font = Font(name=FONT, size=10, bold=True)
    for col, w in zip("ABCDE", (30, 48, 16, 16, 16)):
        ws.column_dimensions[col].width = w
    old = ci_sheet(wb, "Tax Paid Summary")
    if old is not None:
        position = wb.index(old)
        del wb[old.title]
        wb.move_sheet(ws, offset=position - wb.index(ws))
    ws.title = "Tax Paid Summary"
    log.add("Tax Paid Summary", f"{y_a}: all 12 months; {y_b}: first {months_prior} months")
    return ws


# ---------------------------------------------------------------- Module 4
def copy_cell(src, dst):
    dst.value = src.value
    if src.has_style:
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.number_format = copy(src.alignment), src.number_format


def build_adt_and_dashboard(wb, toi_ws, company_ws, log, link_all=False):
    tpl = load_workbook(HERE / "rules" / "ADT_rules_template.xlsx")["ADT"]
    adt = wb.create_sheet("ADT")
    for row in tpl.iter_rows():
        for c in row:
            copy_cell(c, adt.cell(c.row, c.column))
    for k, d in tpl.column_dimensions.items():
        adt.column_dimensions[k].width = d.width
    for k, d in tpl.row_dimensions.items():
        if d.height:
            adt.row_dimensions[k].height = d.height
    for m in tpl.merged_cells.ranges:
        adt.merge_cells(str(m))
    # Link every item to the TOI sheet by its code (A0..A52, B0.., C.., D..).
    toi_row = {}
    for r in range(1, toi_ws.max_row + 1):
        code = toi_ws.cell(r, 2).value
        if isinstance(code, str) and re.fullmatch(r"[ABCD]\d+", code.strip()):
            toi_row.setdefault(code.strip(), r)
    fixed = mismatched = blank = linked = 0
    for r in range(1, adt.max_row + 1):
        code = adt.cell(r, 4).value
        if not (isinstance(code, str) and re.fullmatch(r"[ABCD]\d+", code.strip())):
            continue
        t = toi_row.get(code.strip())
        if t is None:
            log.add("ADT", f"Code {code} (row {r}) not found in TOI")
            continue
        for col, tcol in ((5, "C"), (6, "D")):
            want = f"=TOI!{tcol}{t}"
            cur = adt.cell(r, col).value
            if cur == want:
                continue
            if cur is None:
                blank += 1
                if not link_all:
                    continue  # V5.1 leaves these rows unlinked; keep that unless --link-all
                linked += 1
            elif isinstance(cur, (int, float)):
                fixed += 1
                log.add("ADT", f"{L(col)}{r} ({code}): hard-coded {cur:,} replaced by {want}")
            elif isinstance(cur, str) and cur.startswith("="):
                mismatched += 1
                log.add("ADT", f"{L(col)}{r} ({code}): formula {cur} pointed to a different TOI row, now {want}")
            adt.cell(r, col).value = want
    log.add("ADT", f"{fixed} stale hard-coded numbers replaced by TOI links, {mismatched} mis-pointed links corrected")
    if blank and not link_all:
        log.add("ADT", f"{blank} Current/Prior cells have no TOI link in V5.1 (current assets, liabilities and most income lines show blank). Run with --link-all to fill them.")
    if linked:
        log.add("ADT", f"{linked} previously blank Current/Prior cells linked to TOI (--link-all)")
    dash = wb.create_sheet("Audit Dashboard")
    dash["A1"] = "TAX AUDIT DASHBOARD"
    dash.merge_cells("A1:F1")
    dash["A1"].font = Font(name=FONT, size=14, bold=True)
    dash["A1"].alignment = Alignment(horizontal="center")
    for r in range(3, 8):
        dash.cell(r, 1, f"='Company Info'!A{r}")
        dash.cell(r, 2, f"='Company Info'!B{r}")
    dash["A8"], dash["B8"] = "Report Date:", "=TODAY()"
    dash["B8"].number_format = "yyyy-mm-dd"
    for row in tpl.iter_rows():
        for c in row:
            if c.value is None:
                continue
            d = dash.cell(c.row + ADT_OFFSET, c.column)
            d.value = f'=IF(ADT!{c.coordinate}="","",ADT!{c.coordinate})'
            if c.has_style:
                d.font, d.fill, d.border = copy(c.font), copy(c.fill), copy(c.border)
                d.alignment, d.number_format = copy(c.alignment), c.number_format
    for k, d in tpl.column_dimensions.items():
        dash.column_dimensions[k].width = d.width
    dash.column_dimensions["B"].width = max(dash.column_dimensions["B"].width or 0, 18)
    return dash


# ---------------------------------------------------------------- finance folders
FINANCE_FOLDERS = {
    "GL": ("General ledger", {"GL", "GENERAL LEDGER", "LEDGER"}),
    "TB": ("Trial balance", {"TB", "TRIAL BALANCE"}),
    "BANK": ("Bank", {"BANK", "BANKS", "BANK STATEMENT", "BANK STATEMENTS"}),
    "SALES": ("Sales", {"SALES", "SALE", "REVENUE"}),
    "AR": ("A/R", {"AR", "A-R", "A_R", "A/R", "RECEIVABLE", "RECEIVABLES"}),
    "EXPENSE": ("Expense", {"EXPENSE", "EXPENSES", "PURCHASE LEDGER"}),
}
FINANCE_NAMES = set().union(*(v[1] for v in FINANCE_FOLDERS.values()))
FIN_EXT = {".xlsx", ".xlsm", ".xls", ".csv", ".pdf"}


def finance_folder(root, key):
    for d in Path(root).iterdir():
        if d.is_dir() and d.name.upper() in FINANCE_FOLDERS[key][1] and not d.name.startswith("."):
            return d
    return None


def check_finance(company_folder):
    """Scan the finance sub-folders (GL, TB, Bank, Sales, A/R, Expense).
    Returns (summary rows, file rows). Reads only headers/row counts, never changes files."""
    summary, files = [], []
    for key, (label, _names) in FINANCE_FOLDERS.items():
        folder = finance_folder(company_folder, key)
        if folder is None:
            summary.append((label, "", 0, "", "Optional" if key in ("AR", "EXPENSE") else "Folder not found"))
            continue
        items = sorted(p for p in folder.rglob("*") if p.is_file() and not p.name.startswith((".", "~$"))
                       and p.suffix.lower() in FIN_EXT)
        exts = sorted({p.suffix.lower() for p in items})
        status = "Files found; validate contents" if items else "No files"
        if any(p.suffix.lower() == ".pdf" for p in items):
            status = "Not ready: convert PDF and review"
        summary.append((label, folder.name, len(items), ", ".join(exts), status))
        for p in items:
            sheets, rows, note = "", "", ""
            try:
                if p.suffix.lower() in (".xlsx", ".xlsm"):
                    wb = load_workbook(p, read_only=True, data_only=True)
                    sheets = ", ".join(wb.sheetnames[:4])
                    rows = wb.worksheets[0].max_row
                    if "Conversion Review" in wb.sheetnames:
                        note = "REVIEW: PDF extraction; check balances and map columns"
                    wb.close()
                elif p.suffix.lower() == ".csv":
                    with open(p, "r", encoding="utf-8-sig", errors="replace") as f:
                        rows = sum(1 for _ in f)
                elif p.suffix.lower() == ".pdf":
                    note = "NOT READY: convert PDF to Excel; review and map columns"
                elif p.suffix.lower() == ".xls":
                    note = "Old .xls: save as .xlsx"
            except Exception as exc:
                note = "Cannot open: " + str(exc).splitlines()[0]
            files.append((label, str(p.relative_to(folder)), sheets, rows, note or "Readable"))
    return summary, files


# ---------------------------------------------------------------- source check
def check_sources(company_folder):
    """Validate every monthly file with the same reader consolidate.py uses.
    Returns rows of (tax, year, file name, data rows, result)."""
    root = Path(company_folder)
    out = []
    for tax, layout in cons.LAYOUTS.items():
        folder = find_dir(root, tax)
        if folder is None:
            out.append((tax, "", "(folder not found)", 0, "MISSING"))
            continue
        found = False
        for year, path in cons.files_in(folder):
            found = True
            try:
                data = cons.read_source(path, layout)
                out.append((tax, year, path.name, len(data.rows), "READY"))
            except Exception as exc:
                out.append((tax, year, path.name, 0, "ERROR: " + str(exc).splitlines()[0]))
        if not found:
            out.append((tax, "", "(no monthly .xlsx files)", 0, "EMPTY"))
    for name in ("TOI", "INFO"):
        folder = find_dir(root, name)
        f = first_xlsx(folder) if folder else None
        out.append((name, "", f.name if f else "(not found)", 0, "READY" if f else "MISSING"))
    return out


# ---------------------------------------------------------------- sections
# Each button in the Control Center builds one section. Sheets a section needs
# internally (for example the dashboard needs TOI and Company Info) are kept.
SECTIONS = {
    "consolidate": ("Consolidated data (PR, SR, TOS, WHT)", {"PR", "SR", "TOS", "WHT"}),
    "summary": ("Summary MTR", {"Summary MTR"}),
    "tois": ("TOIs Overview", {"TOI", "FULL TOI"}),
    "toi": ("TOI compare", {"TOI"}),
    "fulltoi": ("FULL TOI", {"FULL TOI"}),
    "mtr": ("MTR detail and summary", {"MTR_PR", "MTR_SR", "MTR_TOS", "MTR_WHT", "Summary MTR"}),
    "mtr_pr": ("MTR detail: PR", {"MTR_PR"}),
    "mtr_sr": ("MTR detail: SR", {"MTR_SR"}),
    "mtr_tos": ("MTR detail: TOS", {"MTR_TOS"}),
    "mtr_wht": ("MTR detail: WHT", {"MTR_WHT"}),
    "tps": ("Tax Paid Summary", {"Tax Paid Summary"}),
    "dashboard": ("Audit Dashboard", {"Audit Dashboard", "ADT", "TOI", "Company Info"}),
}


def compact_consolidation(wb, file_rows):
    """Re-layout tax sheets after summary/detail have consumed file subtotal rows."""
    for tax, subtotal_rows in file_rows.items():
        if not subtotal_rows:
            continue
        src = wb[tax]
        width, last = src.max_column, max(subtotal_rows)
        year_header = last + 3
        if src.cell(year_header, 1).value != "Year":
            raise ValueError(f"{tax}: expected year totals after monthly file subtotals")
        years = [r for r in range(year_header + 1, src.max_row + 1)
                 if src.cell(r, 1).value not in (None, "")]
        target = wb.create_sheet(f"_compact_{tax}")
        for col, dim in src.column_dimensions.items():
            target.column_dimensions[col] = copy(dim)
        def copy_row(source_row, destination_row):
            for c in range(1, width + 1):
                source = src.cell(source_row, c)
                dest = target.cell(destination_row, c, source.value)
                dest._style = source._style
        copy_row(1, 1)
        target["A1"] = "Grand Total"
        for index, r in enumerate(years, 2):
            copy_row(r, index)
            target.cell(index, 1, f"Year {src.cell(r, 1).value}")
        header_start = len(years) + 2
        copy_row(2, header_start)
        copy_row(3, header_start + 1)
        shift = header_start - 2
        for merged in src.merged_cells.ranges:
            if 2 <= merged.min_row <= merged.max_row <= 3:
                target.merge_cells(start_row=merged.min_row+shift, end_row=merged.max_row+shift,
                                   start_column=merged.min_col, end_column=merged.max_col)
        out = header_start + 2
        excluded = set(subtotal_rows)
        for r in range(4, last + 1):
            if r not in excluded:
                copy_row(r, out)
                out += 1
        target.freeze_panes = f"A{header_start + 2}"
        for r in range(1, len(years) + 2):
            for cell in target[r]:
                cell.fill, cell.font = LBLUE, Font(name=FONT, size=10, bold=True)
        position = wb.index(src)
        del wb[tax]
        target.title = tax
        wb.move_sheet(target, offset=position - wb.index(target))


# ---------------------------------------------------------------- driver
def build_workbook(company_folder, months_prior=3, link_all=False):
    """Build every report in memory. Returns (workbook, log). Nothing is written to the company folder."""
    import tempfile
    root = Path(company_folder)
    log = Log()
    tdir = Path(tempfile.mkdtemp(prefix="taxaudit_"))
    tmp = tdir / "consolidated.xlsx"
    original_hidden = cons.hidden
    # Finance folders (GL, Bank, ...) are not tax data: keep them out of the tax consolidation.
    cons.hidden = lambda p: original_hidden(p) or (p.parent == root and p.name.upper() in FINANCE_NAMES)
    try:
        with open(tdir / "consolidated.log", "w", encoding="utf-8") as lf:
            cons.consolidate(root, tmp, lf, excel_job=True)
    finally:
        cons.hidden = original_hidden
    wb = load_workbook(tmp)
    import shutil
    shutil.rmtree(tdir, ignore_errors=True)
    for ws in list(wb.worksheets):
        if ws.title.upper() not in SUM_COLS and ws.title != cons.MTR_INDEX:
            if ws.max_row <= 4 and str(ws["A4"].value).startswith("No data found"):
                del wb[ws.title]
                log.add("Consolidate", f"Folder {ws.title} has no year/month files, skipped")

    toi_dir, info_dir = find_dir(root, "TOI"), find_dir(root, "INFO")
    toi_file = first_xlsx(toi_dir) if toi_dir else None
    info_file = first_xlsx(info_dir) if info_dir else None
    log.add("Inputs", f"TOI file: {toi_file.name if toi_file else 'NOT FOUND'}")
    log.add("Inputs", f"Company Info file: {info_file.name if info_file else 'NOT FOUND'}")

    build_summary_mtr(wb, log)
    toi = full = None
    if toi_file:
        toi_wb = load_workbook(toi_file, data_only=True)
        toi = build_toi(wb, toi_wb, log)
        full = build_full_toi(wb, toi_wb, log)
    file_rows = {}
    if cons.MTR_INDEX in wb:
        for tax, row in wb[cons.MTR_INDEX].iter_rows(min_row=2, max_col=2, values_only=True):
            file_rows.setdefault(tax, []).append(int(row))
    build_mtr_detail(wb, root, log)
    company = None
    if info_file:
        info_wb = load_workbook(info_file, data_only=True)
        company = company_info(wb, info_wb, log)
        tax_paid_summary(wb, info_wb, months_prior, log)
    if toi is not None and company is not None:
        build_adt_and_dashboard(wb, toi, company, log, link_all)
    else:
        log.add("Audit Dashboard", "Skipped: needs both TOI file and Company Info file")

    compact_consolidation(wb, file_rows)

    lg = wb.create_sheet("Log")
    lg.append(("DateTime", "Where", "Message"))
    for rec in log:
        lg.append(rec)
    lg.column_dimensions["C"].width = 110
    order = ["Audit Dashboard", "Tax Paid Summary", "TOI", "FULL TOI", "Summary MTR",
             "MTR_PR", "MTR_SR", "MTR_TOS", "MTR_WHT", "PR", "SR", "TOS", "WHT", "ADT", "Company Info", "Log"]
    wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + \
                 [s for s in wb._sheets if s.title not in order]
    return wb, log


def build(company_folder, output, months_prior=3, link_all=False, sections=None):
    """Build and save to an Excel file. sections=None saves every report."""
    wb, log = build_workbook(company_folder, months_prior, link_all)
    if sections:
        keep = {"Log"}
        for key in sections:
            keep |= SECTIONS[key][1]
        for name in list(wb.sheetnames):
            if name not in keep:
                del wb[name]
        log.add("Build", "Sections: " + ", ".join(SECTIONS[k][0] for k in sections))
        wb["Log"].append((datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Build", "Sections: " + ", ".join(SECTIONS[k][0] for k in sections)))
    wb.active = 0
    wb.save(output)
    return log


if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Tax Audit Engine: build all reports")
    p.add_argument("company_folder", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--months", type=int, default=3, help="months of the newest year; older year always includes 12 months (default 3)")
    p.add_argument("--link-all", action="store_true", help="also link ADT rows that V5.1 leaves blank")
    a = p.parse_args()
    for rec in build(a.company_folder, a.output, a.months, a.link_all):
        print(" | ".join(rec[1:]))
