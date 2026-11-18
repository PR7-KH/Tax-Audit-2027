"""Source inspect: one pass that checks every source folder and validates every file.

inspect(company_folder) returns folder rows and file rows, each marked
    ✓ ready   ⚠ usable with a note   ✗ not ready   — not provided (optional)
plus the finance tables it already read (so GL / Sales / Bank open instantly) and the
PDFs that still need converting. Nothing here touches the GUI or changes files, except
standardize_extraction(), which writes a new clean Excel next to a converted PDF.
"""
import re
from collections import Counter, defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook

import engine
import finance_view as fv
from input_tools import conversion_records, converted_output, visible_files

try:
    from toi_pdf import is_toi_output
except Exception:  # pragma: no cover - toi_pdf needs optional packages on some machines
    def is_toi_output(path):
        return False

READY, NOTE, BAD, NONE = "✓", "⚠", "✗", "—"
COLUMNS = ["", "Source", "Year / scope", "Files", "Format", "Remark"]
FILE_COLUMNS = ["", "Source", "Year / scope", "File", "Format", "Remark"]
TAG = {READY: "total", NOTE: "warning", BAD: "error", NONE: "even"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
DATA_EXT = {".xlsx", ".xlsm", ".csv"}
NEEDS_DATE = {"GL", "BANK", "SALES"}
TOTAL_ROW = re.compile(r"^\s*(grand\s+)?(sub\s*)?total\b|^\s*សរុប", re.I)


def _fmt(paths):
    labels = {".xlsx": "Excel", ".xlsm": "Excel", ".xls": "Old Excel", ".pdf": "PDF", ".csv": "CSV", ".docx": "Word"}
    c = Counter(labels.get(p.suffix.lower(), "Other") for p in paths)
    return ", ".join(f"{n} {k}" for k, n in sorted(c.items())) or "—"


def _month_span(months):
    months = sorted(set(months))
    if not months:
        return ""
    if months == list(range(months[0], months[-1] + 1)):
        return MONTHS[months[0] - 1] + ("–" + MONTHS[months[-1] - 1] if len(months) > 1 else "")
    return ", ".join(MONTHS[m - 1] for m in months)


def file_month(name, year):
    """Month number from a monthly file name such as SALE_01_2025.xlsx or PR 2025-03.xlsx."""
    stem = Path(name).stem
    for pat in (rf"(?<!\d)(0?[1-9]|1[0-2])[ _\-.]+{year}(?!\d)", rf"(?<!\d){year}[ _\-.]+(0?[1-9]|1[0-2])(?!\d)"):
        m = re.search(pat, stem)
        if m:
            return int(m.group(1))
    return None


# ------------------------------------------------------------------ PDF -> standard Excel
def _number(text):
    t = str(text).strip()
    if re.fullmatch(r"\(?-?[\d,]+(\.\d+)?\)?", t) and re.search(r"\d", t):
        neg = t.startswith("(") and t.endswith(")")
        v = float(t.strip("()").replace(",", ""))
        v = -v if neg else v
        return int(v) if v == int(v) and abs(v) < 1e15 else v
    return t


def standardize_extraction(extraction, target=None, kind=""):
    """Turn a raw PDF extraction (sheets 'P1 Table 1', ...) into one clean table.
    Returns (ok, message). Writes target only when ok and target does not exist yet."""
    wb = load_workbook(extraction, read_only=True)
    try:
        names = [n for n in wb.sheetnames if re.fullmatch(r"P\d+ Table \d+", n)]
        if not names:
            return False, "the PDF gave text lines only, no table"
        tables = [[list(r) for r in wb[n].iter_rows(values_only=True)] for n in names]
    finally:
        wb.close()
    norm = lambda row: [re.sub(r"\s+", " ", str(v or "")).strip().casefold() for v in row]
    first = tables[0]
    h = fv._header_index(first)
    if h is None:
        return False, "no heading row was found in the PDF tables"
    header = [re.sub(r"\s+", " ", str(v or "")).strip() for v in first[h]]
    width = max(i + 1 for i, v in enumerate(header) if v)
    header = [v or f"Column {i + 1}" for i, v in enumerate(header[:width])]
    rows, skipped = [], 0
    for ti, table in enumerate(tables):
        start = h + 1 if ti == 0 else 0
        if ti and table and norm(table[0][:width]) == norm(header):
            start = 1
        body = table[start:]
        if body and len([v for v in body[0] if v not in (None, "")]) > width:
            skipped += 1
            continue
        for r in body:
            r = (list(r) + [None] * width)[:width]
            cells = [v for v in r if v not in (None, "")]
            if not cells or TOTAL_ROW.match(str(cells[0])):
                continue
            rows.append([None if v in (None, "") else _number(v) for v in r])
    if not rows:
        return False, "the PDF tables have a heading but no data rows"
    t = fv.Table(header, rows)
    if not t.amount:
        return False, "no amount column (Amount, Debit, Credit, Money In...) in the PDF tables"
    if kind.upper() in NEEDS_DATE and t.date_col is None:
        return False, "no Date column in the PDF tables"
    note = f"{len(rows):,} rows from {len(tables) - skipped} PDF tables" + (f"; {skipped} tables with other headings left out" if skipped else "")
    if target is not None:
        target = Path(target)
        if target.exists():
            return False, f"{target.name} already exists, so nothing was written"
        out = Workbook()
        ws = out.active
        ws.title = "Data"
        ws.append(header)
        for r in rows:
            ws.append(r)
        out.save(target)
    return True, note


# ------------------------------------------------------------------ checks
def check_data_file(path, kind):
    """(mark, remark, (path, columns, rows) or None) for one finance Excel / CSV file."""
    try:
        cols, rows = fv.read_rows(path)
    except Exception as exc:
        try:
            wb = load_workbook(path, read_only=True)
            raw = "Conversion Review" in wb.sheetnames
            wb.close()
        except Exception:
            raw = False
        if raw:
            return BAD, "Raw PDF extraction, not a standard table: keep one heading row (Date, Description, Amount...) and the data rows", None
        return BAD, f"Cannot read: {str(exc).splitlines()[0]}", None
    t = fv.Table(cols, rows)
    if not rows:
        return BAD, "Heading found but no data rows", None
    if not t.amount:
        return BAD, "No amount column: need a heading such as Amount, Debit / Credit, Money In / Out or Balance", None
    if kind in NEEDS_DATE and t.date_col is None:
        return BAD, "No Date column", None
    span = ""
    if t.months:
        years = sorted({m[:4] for m in t.months})
        span = f", {_month_span(int(m[5:]) for m in t.months if m[:4] == years[-1])} {years[-1]}" if len(years) == 1 else f", {t.months[0]} to {t.months[-1]}"
    return READY, f"{len(rows):,} rows{span}", (Path(path), cols, rows)


def _amount_total(part):
    t = fv.Table(part[1], part[2])
    return round(sum(v for r in t.rows for c in t.amount[:1] for v in [r[c]] if isinstance(v, (int, float)) and not isinstance(v, bool)), 2)


def inspect(company_folder, import_results=None):
    root = Path(company_folder)
    records = _safe(lambda: conversion_records(root), {})
    folder_rows, file_rows = [], []
    tables, pending, data_parts = {}, [], {}

    def frow(mark, source, scope, paths, remark, path):
        folder_rows.append(dict(mark=mark, values=[mark, source, scope, len(paths), _fmt(paths), remark], path=path))

    def filerow(mark, source, scope, path, remark):
        file_rows.append(dict(mark=mark, values=[mark, source, scope, path.name, _fmt([path]), remark], path=path))

    # ---- Download (organize + import log in one place)
    downloads = [d for d in root.iterdir() if d.is_dir() and d.name.casefold() == "download"]
    if downloads:
        waiting = visible_files(downloads[0])
        reasons = {str(r[0]): r[2] for r in (import_results or []) if len(r) > 2}
        for p in waiting:
            why = reasons.get(str(p.relative_to(root)), "")
            remark = ("Not moved: " + why.split(":", 1)[-1].strip() + ". Rename it to include PR, SR, TOS, WHT, GL, TB, Bank, Sales, AR, "
                      "Expense or TOI (tax files also need the year), or move it yourself") if why.startswith(("SKIPPED", "ERROR")) else \
                     "Waiting: click Organize Download"
            filerow(BAD if why.startswith(("SKIPPED", "ERROR")) else NOTE, "Download", "—", p, remark)
        moved = sum(1 for r in import_results or [] if len(r) > 2 and r[2] == "MOVED")
        if waiting:
            frow(NOTE, "Download", "—", waiting, f"{len(waiting)} file(s) still in Download" + (f"; {moved} moved this time" if moved else "")
                 + ". Click Organize Download; files whose name is not recognised stay here", downloads[0])
        else:
            frow(READY, "Download", "—", [], f"All files organized{f' ({moved} moved)' if moved else ''}", downloads[0])

    # ---- monthly tax returns
    checked = {(r[0], str(r[1]), r[2]): r for r in _safe(lambda: engine.check_sources(root), [])}
    for tax in engine.cons.LAYOUTS:
        directory = engine.find_dir(root, tax)
        if directory is None:
            frow(BAD, tax, "—", [], f"Missing folder: create {tax}/<year> and put the 12 monthly Excel returns in it", root)
            continue
        files = visible_files(directory)
        if not files:
            frow(BAD, tax, "—", [], "Folder is empty", directory)
            continue
        groups = defaultdict(list)
        for p in files:
            rel = p.relative_to(directory)
            groups[rel.parts[0] if len(rel.parts) > 1 else ""].append(p)
        years = sorted(y for y in groups if re.fullmatch(r"\d{4}", y))
        for year in sorted(groups, key=lambda y: (y == "", y)):
            paths = groups[year]
            marks, months, problems = [], [], []
            for p in paths:
                ext = p.suffix.lower()
                if not year:
                    m, r = BAD, f"Move into a year folder, for example {tax}/2025"
                elif ext == ".xlsx":
                    c = checked.get((tax, year, p.name))
                    if c is None:
                        m, r = BAD, "Not read (check the year folder name)"
                    elif c[4] == "READY" and c[3]:
                        m, r = READY, f"{c[3]:,} rows"
                        mm = file_month(p.name, year)
                        if mm:
                            months.append(mm)
                    elif c[4] == "READY":
                        m, r = BAD, "No transaction rows"
                    else:
                        m, r = BAD, str(c[4]).replace("ERROR: ", "Cannot read: ")
                elif ext == ".pdf":
                    twin = p.with_suffix(".xlsx")
                    m, r = (NOTE, f"PDF copy; {twin.name} is used") if twin.exists() else \
                           (BAD, "PDF: the monthly return must be the Excel (.xlsx) file from the tax system")
                elif ext == ".xls":
                    m, r = BAD, "Old Excel (.xls): open it and save as .xlsx"
                else:
                    m, r = BAD, "Unsupported format"
                marks.append(m)
                filerow(m, tax, year or "(no year folder)", p, r)
                if m == BAD:
                    problems.append(f"{p.name}: {r}")
            dup = [MONTHS[k - 1] for k, n in Counter(months).items() if n > 1]
            if problems:
                mark = BAD
                remark = f"{len(problems)} of {len(paths)} file(s) not ready. " + problems[0]
            else:
                full = len(set(months)) == 12
                newest = year == (years[-1] if years else year)
                if months and not full and not newest:
                    mark = NOTE
                    remark = f"{len(set(months))} months; missing {_month_span(set(range(1, 13)) - set(months))}"
                else:
                    mark = READY
                    remark = (f"{len(set(months))} months ({_month_span(months)})" if months else f"{sum(x == READY for x in marks)} file(s) ready")
                    if months and not full:
                        remark += ", newest year"
                if dup:
                    mark, remark = NOTE, remark + f"; two files for {', '.join(dup)}"
            frow(mark, tax, year or "(no year folder)", paths, remark, directory / year if year else directory)

    # ---- TOI and Info
    for name, need in (("TOI", "annual return (TOI) Excel"), ("Info", "company information Excel")):
        directory = engine.find_dir(root, name)
        if directory is None:
            frow(BAD, name, "All", [], f"Missing folder: add the {name} folder with the {need}", root)
            continue
        paths = [p for p in visible_files(directory) if not (name == "TOI" and is_toi_output(p))]
        xl = [p for p in paths if p.suffix.lower() in (".xlsx", ".xlsm")]
        ok = None
        for p in paths:
            if p in xl:
                try:
                    wb = load_workbook(p, read_only=True)
                    sheets = wb.sheetnames
                    wb.close()
                    if name == "TOI" and not any(s.lower().replace(" ", "") == "step5" for s in sheets):
                        m, r = NOTE, "Opens, but no Step5 sheet: TOI lines may not be read"
                    else:
                        m, r = READY, f"{len(sheets)} sheets"
                        ok = ok or p
                except Exception as exc:
                    m, r = BAD, f"Cannot open: {str(exc).splitlines()[0]}"
            elif p.suffix.lower() == ".pdf":
                m, r = NOTE, "PDF kept for reference; the system reads the Excel file"
            else:
                m, r = BAD, "Use the .xlsx file"
            filerow(m, name, "All", p, r)
        if ok:
            frow(READY, name, "All", paths, f"Uses {ok.name}" + (f" (+{len(xl) - 1} more Excel, only the first is used)" if len(xl) > 1 else ""), directory)
        else:
            frow(BAD, name, "All", paths, f"No usable {need}" if paths else "Folder is empty", directory)

    # ---- audited financial statements (optional): PDF with text, or Word
    try:
        import audit_report
        ar_dir, ar_files = audit_report.report_files(root)
    except Exception:
        ar_dir, ar_files = None, []
    if ar_dir is None:
        frow(NONE, "Audit Report", "All", [], "Not provided (optional): add an 'Audit Report' folder with the audited financial statements (PDF or Word)", root)
    else:
        good = []
        for p in ar_files:
            try:
                if p.suffix.lower() == ".pdf":
                    import pdfplumber
                    with pdfplumber.open(str(p)) as pdf:
                        chars = sum(len((pg.extract_text() or "").strip()) for pg in pdf.pages[:6])
                    m, r = (READY, f"Text PDF, {len(pdf.pages)} page(s)") if chars >= 300 else \
                           (BAD, "Scanned PDF (no text): ask for the text PDF or Word version, or OCR it")
                else:
                    m, r = READY, "Word document"
            except Exception as exc:
                m, r = BAD, "Cannot read: " + str(exc).splitlines()[0]
            filerow(m, "Audit Report", "All", p, r)
            if m == READY:
                good.append(p)
        others = [p for p in visible_files(ar_dir) if p not in ar_files]
        for p in others:
            filerow(NOTE, "Audit Report", "All", p, "Not used: put the report as PDF or Word (.docx)")
        if good:
            frow(READY, "Audit Report", "All", ar_files, f"Uses {good[0].name}" + (f" (+{len(ar_files) - 1} more: keep only the final report)" if len(ar_files) > 1 else ""), ar_dir)
        else:
            frow(BAD, "Audit Report", "All", ar_files, "No readable PDF / Word report" if ar_files else "Folder has no PDF or Word file", ar_dir)

    # ---- tax payment history (optional): GDT Payment History PDF / Excel, else Info Sheet2 / Sheet3
    try:
        import tax_paid
        tp_rows, tp_dir, _tp_notes = tax_paid.collect(root)
    except Exception:
        tp_rows, tp_dir = [], None
    years = sorted({r[0] for r in tp_rows})
    if tp_dir is None:
        frow(NOTE if not years else READY, "Tax Paid", "All", [],
             (f"Uses Info Sheet2 / Sheet3: {', '.join(map(str, years))}" if years else
              "Not provided (optional): add a 'Tax Paid' folder with the GDT Payment History (PDF or Excel)"), root)
    else:
        files = [p for p in visible_files(tp_dir)]
        for p in files:
            filerow(READY if p.suffix.lower() in (".pdf", ".xlsx", ".xlsm") else NOTE, "Tax Paid", "All", p,
                    "Payment History" if p.suffix.lower() in (".pdf", ".xlsx", ".xlsm") else "Not used: PDF or Excel only")
        frow(READY if years else BAD, "Tax Paid", "All", files,
             f"Years {', '.join(map(str, years))}" if years else "No Payment History table found in the files", tp_dir)

    # ---- financial folders
    for key, (label, _names) in engine.FINANCE_FOLDERS.items():
        directory = engine.finance_folder(root, key)
        if directory is None:
            frow(NONE, label, "All", [], f"Not provided (optional): add the {key} folder to use {label}", root)
            continue
        paths = [p for p in visible_files(directory) if p.suffix.lower() in engine.FIN_EXT]
        if not paths:
            frow(NONE, label, "All", [], "Folder is empty", directory)
            continue
        results, parts = {}, []
        for p in paths:
            if p.suffix.lower() in DATA_EXT:
                m, r, part = check_data_file(p, key)
                results[p] = (m, r)
                if part:
                    parts.append(part)
            elif p.suffix.lower() == ".xls":
                results[p] = (BAD, "Old Excel (.xls): open it and save as .xlsx")
        # duplicates: two ready files with the same rows and the same total
        seen = {}
        for part in parts:
            sig = (len(part[2]), _amount_total(part))
            if sig in seen:
                results[part[0]] = (NOTE, f"Same rows and total as {seen[sig].name}: keep one, or totals will double")
            else:
                seen[sig] = part[0]
        for p in paths:
            if p.suffix.lower() != ".pdf":
                continue
            twins = [q for q in paths if q.suffix.lower() in DATA_EXT and q.stem.casefold() == p.stem.casefold()]
            if twins and results.get(twins[0], (BAD,))[0] == READY:
                results[p] = (READY, f"Converted: {twins[0].name} is used")
                continue
            if twins:
                results[p] = (BAD, f"Converted to {twins[0].name}, but it is not usable: {results[twins[0]][1]}")
                continue
            out = _safe(lambda: converted_output(root, p, records), None)
            if out is not None:
                ok, why = _safe(lambda: standardize_extraction(out, None, key), (False, "the extraction could not be read"))
                results[p] = (BAD, (f"Converted, but not saved as a standard table yet: click Convert PDFs again" if ok else
                                    f"Converted, but not a standard table ({why}). Export it from the accounting system as Excel, "
                                    f"or save a clean Excel named {p.stem}.xlsx in this folder"))
                continue
            total = _safe(lambda: __import__("sales_reconciliation").read_pdf_total(p), None)
            copy_of = next((q for q in parts if total is not None and abs(_amount_total(q) - total) < 1), None)
            if copy_of:
                results[p] = (READY, f"PDF copy: its total agrees with {copy_of[0].name}, which is used")
            else:
                results[p] = (BAD, "PDF: click Convert PDFs to make a standard Excel")
                pending.append(p)
        for p in paths:
            m, r = results.get(p, (BAD, "Unsupported format"))
            filerow(m, label, "All", p, r)
        dup_free = [part for part in parts if results[part[0]][0] == READY]
        bad = [p for p in paths if results[p][0] == BAD]
        if dup_free:
            tables[key] = fv.table_from_parts(dup_free, key)
            rows = sum(len(part[2]) for part in dup_free)
            t = tables[key]
            span = f", {t.months[0]} to {t.months[-1]}" if t.months else ""
            text = f"{len(dup_free)} file(s) used, {rows:,} rows{span}"
            notes = [p for p in paths if results[p][0] == NOTE]
            if bad or notes:
                first = (bad or notes)[0]
                frow(NOTE, label, "All", paths, f"{text}. {len(bad) + len(notes)} file(s) need attention: {first.name}: {results[first][1]}", directory)
            else:
                frow(READY, label, "All", paths, text, directory)
        else:
            first = bad[0] if bad else paths[0]
            frow(BAD, label, "All", paths, f"Not ready: {first.name}: {results[first][1]}", directory)

    counts = Counter(r["mark"] for r in folder_rows)
    return dict(folder=str(root), rows=folder_rows, files=file_rows, counts=counts, tables=tables, pending_pdfs=pending)


def _safe(fn, default):
    try:
        return fn()
    except Exception:
        return default


def banner(result):
    c = result["counts"]
    text = f"{READY} {c.get(READY, 0)} ready    {NOTE} {c.get(NOTE, 0)} note    {BAD} {c.get(BAD, 0)} not ready    {NONE} {c.get(NONE, 0)} not provided"
    if c.get(BAD):
        text += "    —  Fix the ✗ rows (double-click a row to open it). Build all data still runs; reports that need a ✗ source are skipped."
    else:
        text += "    —  Ready: click Build all data."
    if result["pending_pdfs"]:
        text += f"    {len(result['pending_pdfs'])} PDF(s) can be converted: click Convert PDFs."
    return text


def display(result, per_file=False):
    rows = result["files"] if per_file else result["rows"]
    return [r["values"] for r in rows], [TAG[r["mark"]] for r in rows], [r["path"] for r in rows]
