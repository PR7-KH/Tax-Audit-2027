"""Tax Payment History: the GDT "Payment History" statement (PDF) or the same table in Excel, with the
monthly detail, a year summary and comparisons with the tax returns.

Sources, all used together (the newest / most complete copy of each year wins):
  * Tax Paid folder: GDT Payment History PDFs and Excel files with the same layout;
  * Info workbook: Sheet2 / Sheet3 (the existing company-info layout).
Layout: a year row (No, year, 12 monthly totals, total), then one row per tax type
(No, Khmer name, January..December, Total) with the English name on the next line / row.

build(company_folder) -> audit_analysis.Analysis.  Nothing here touches the GUI.
"""
import re
from collections import defaultdict
from pathlib import Path

import audit_analysis as aa

from engine import TAX_PAID_NAMES as NAMES
MONTHS = aa.MONTHS


def _money(v):
    if v in (None, "", "-", "–"):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace(",", "").strip()
    try:
        return float(t)
    except ValueError:
        return None


def _english(text):
    lines = [l.strip() for l in str(text or "").splitlines()]
    eng = [l for l in lines if re.search(r"[A-Za-z]", l) and not re.search(r"[ក-៿\x00-]", l)]
    return re.sub(r"\s+", " ", " ".join(eng)).strip()


def _khmer(text):
    lines = [l.strip() for l in str(text or "").splitlines() if re.search(r"[ក-៿]", l)]
    return " ".join(lines)


def kind_of(name):
    """Short group used for the comparisons."""
    n = name.lower()
    if "fringe" in n:
        return "FBT"
    if "salary" in n:
        return "TOS"
    if "withholding" in n:
        return "WHT"
    if "prepayment" in n:
        return "PREPAY"
    if "annual" in n or "corporate income" in n or n.strip() == "tax on income":
        return "CIT"
    if "value added" in n or n.strip() == "vat":
        return "VAT"
    return "OTHER"


def parse_rows(rows, source):
    """[(year, tax type EN, Khmer, [12 months], total, source)] from table rows (lists of cells)."""
    out, year, last = [], None, None
    for r in rows:
        r = list(r) + [None] * (15 - len(r))
        no, name = r[0], r[1]
        if isinstance(name, (int, float)) and 2000 <= float(name) <= 2100:
            year = int(name)
            last = None
            continue
        if isinstance(name, str) and re.fullmatch(r"\s*(20\d\d)\s*", name):
            year = int(name)
            last = None
            continue
        if year is None:
            continue
        values = [c for c in r[2:] if c is not None and str(c).strip() != ""]
        nums = [_money(c) for c in values]
        if isinstance(name, str) and len(values) >= 13 and str(no or "").strip().isdigit():
            months, total = nums[-13:-1], nums[-1]
            eng, kh = _english(name), _khmer(name)
            rec = [year, eng, kh, months, total, source]
            out.append(rec)
            last = rec if not eng else None
            continue
        if last is not None and isinstance(name, str) and (no in (None, "")) and _english(name):
            last[1] = _english(name)           # English name on the row after the Khmer one (Excel)
            last = None
    return [r for r in out if r[1] or r[2]]


def read_pdf(path):
    import pdfplumber
    rows = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            for t in page.extract_tables():
                rows.extend(t)
    return parse_rows(rows, Path(path).name)


def read_excel(path, sheets=None):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    try:
        for ws in wb.worksheets:
            if sheets and ws.title not in sheets:
                continue
            out.extend(parse_rows([list(r) for r in ws.iter_rows(values_only=True)], f"{Path(path).name} / {ws.title}"))
    finally:
        wb.close()
    return out


def collect(root):
    """All payment rows and notes; one copy per year (the one with more months filled)."""
    root = Path(root)
    notes, found = [], []
    folder = next((d for d in sorted(root.iterdir()) if d.is_dir() and d.name.upper() in NAMES), None) if root.is_dir() else None
    if folder is not None:
        for p in sorted(folder.rglob("*")):
            if not p.is_file() or p.name.startswith((".", "~$")):
                continue
            try:
                if p.suffix.lower() == ".pdf":
                    rows = read_pdf(p)
                elif p.suffix.lower() in (".xlsx", ".xlsm"):
                    rows = read_excel(p)
                else:
                    continue
            except Exception as exc:
                notes.append(f"{p.name}: {str(exc).splitlines()[0]}")
                continue
            if not rows:
                notes.append(f"{p.name}: no Payment History table found")
            found.extend(rows)
    info = aa.subfolder(root, {"INFO", "INFORMATION"})
    if info is not None:
        for p in sorted(info.glob("*.xlsx")):
            if p.name.startswith("~$"):
                continue
            try:
                found.extend(read_excel(p, sheets=None))
            except Exception as exc:
                notes.append(f"Info {p.name}: {str(exc).splitlines()[0]}")
            break
    by_year = defaultdict(lambda: defaultdict(list))
    for r in found:
        by_year[r[0]][r[5]].append(r)
    chosen = []
    for year, sources in sorted(by_year.items()):
        filled = lambda rows: sum(1 for r in rows for v in r[3] if v not in (None, 0.0))
        best = max(sources, key=lambda s: (filled(sources[s]), s))
        chosen.extend(sources[best])
        if len(sources) > 1:
            notes.append(f"{year}: {len(sources)} copies found; using {best} (most months filled)")
    return chosen, folder, notes


def _returns(root):
    """Declared amounts per (year, month) from the monthly returns, for the comparisons."""
    out = {"TOS": defaultdict(float), "WHT": defaultdict(float), "SR": defaultdict(float)}
    try:
        import tos_analysis as ta
        d = aa.subfolder(root, {"TOS"})
        if d is not None:
            for y, m, _p, _l, _c, recs in ta.read_tos(d)[0]:
                if isinstance(y, int):
                    out["TOS"][(y, m)] += sum(r["tax"] + r["fb_tax"] for r in recs)
    except Exception:
        pass
    try:
        import wht_analysis as w
        for ym, _p, _l, rows in w.read_files(aa.subfolder(root, {"WHT"}), "WHT")[0]:
            import consolidate as cons
            lay = cons.LAYOUTS["WHT"]
            out["WHT"][ym] += sum(aa.num(r[lay.first_col]) for r in rows if len(r) > lay.first_col and isinstance(r[0], (int, float)))
    except Exception:
        pass
    try:
        import consolidate as cons
        d = aa.subfolder(root, {"SR"})
        if d is not None:
            lay = cons.LAYOUTS["SR"]
            for y, p in cons.files_in(d):
                m = aa.file_month(p.name, y)
                if m and str(y).isdigit():
                    for r in cons.read_source(p, lay).rows:
                        out["SR"][(int(y), m)] += aa.num(r[lay.first_col - 1]) if len(r) >= lay.first_col else 0
    except Exception:
        pass
    return out


def build(company_folder, tol=1000.0):
    root = Path(company_folder)
    an = aa.Analysis("TAXPAID", "Tax Payment History")
    rows, folder, notes = collect(root)
    for n in notes:
        an.add_note(n)
    if not rows:
        an.add_note("No tax payment data: add the GDT Payment History (PDF or Excel) to a 'Tax Paid' folder, "
                    "or fill Info Sheet2 / Sheet3.")
        an.add_summary_only()
        return an
    years = sorted({r[0] for r in rows})
    an.period = f"{years[0]}–{years[-1]}" if len(years) > 1 else str(years[0])

    # ---- monthly detail (like the Excel)
    detail = []
    for y, eng, kh, months, total, src in sorted(rows, key=lambda r: r[0]):     # statement order within a year
        detail.append([y, eng or kh, kind_of(eng or kh)] + [m for m in months] + [total if total is not None else sum(v or 0 for v in months), kh, src])
    an_detail_cols = ["Year", "Tax type", "Group"] + MONTHS + ["Total", "Khmer name", "Source"]

    # ---- year summary: tax type x year
    names = []
    for r in detail:
        if r[1] not in names:
            names.append(r[1])
    tot = {(r[0], r[1]): r[15] or 0.0 for r in detail}
    summ = []
    for n in names:
        vals = [tot.get((y, n)) for y in years]
        change = None if len(years) < 2 or vals[-1] is None or vals[-2] is None else vals[-1] - vals[-2]
        pct = None if change is None or not vals[-2] else round(change / vals[-2] * 100, 1)
        summ.append([n, kind_of(n)] + vals + [change, pct])
    grand = ["TOTAL", ""] + [sum(tot.get((y, n)) or 0 for n in names) for y in years]
    g_change = grand[-1] - grand[-2] if len(years) > 1 else None
    summ.append(grand + [g_change, round(g_change / grand[-2] * 100, 1) if g_change is not None and grand[-2] else None])

    # ---- monthly totals by group
    groups = ("TOS", "FBT", "WHT", "PREPAY", "CIT", "VAT", "OTHER")
    month_rows = []
    paid = defaultdict(float)
    for r in detail:
        for i, v in enumerate(r[3:15], start=1):
            if v:
                paid[(r[0], i, r[2])] += v
    for y in years:
        for m in range(1, 13):
            vals = [paid.get((y, m, g), 0.0) for g in groups]
            if any(vals):
                month_rows.append([f"{y}-{m:02d}", sum(vals)] + vals)

    # ---- comparisons with the returns
    dec = _returns(root)
    cmp_rows, cmp_tags, findings = [], [], []

    def align(paid_key, declared):
        """Payment of month m is usually for the return of month m−1: use the alignment with the smaller difference."""
        same = sum(abs(paid.get((y, m, paid_key), 0) - declared.get((y, m), 0)) for (y, m) in declared)
        nxt = sum(abs(paid.get((*aa.next_month((y, m)), paid_key), 0) - declared.get((y, m), 0)) for (y, m) in declared)
        return 1 if nxt < same else 0

    for key, label, declared_key, factor in (("TOS", "Salary tax + FBT", "TOS", 1.0), ("WHT", "Withholding tax", "WHT", 1.0),
                                              ("PREPAY", "Prepayment of income tax (1 % of SR sales)", "SR", 0.01)):
        declared = {k: v * factor for k, v in dec[declared_key].items()}
        if not declared:
            continue
        shift = align(key, declared) if key != "TOS" else align("TOS", declared)
        for (y, m) in sorted(declared):
            if y not in years:
                continue
            py, pm = aa.next_month((y, m)) if shift else (y, m)
            p = paid.get((py, pm, key), 0.0) + (paid.get((py, pm, "FBT"), 0.0) if key == "TOS" else 0.0)
            d = declared[(y, m)]
            diff = p - d
            st = "Agrees" if abs(diff) <= max(tol, 0.01 * max(abs(p), abs(d))) else ("Not paid" if not p and d else "Differs")
            cmp_rows.append([label, f"{y}-{m:02d}", f"{py}-{pm:02d}", d, p, diff, st])
            cmp_tags.append({"Agrees": "total", "Not paid": "error", "Differs": "warning"}[st])
            if st != "Agrees":
                findings.append((label, f"{y}-{m:02d}", diff, st))
    try:
        import toi_analysis as ta
        t, path, _e = ta.load(root)
        ty = ta.year_of(path) if t is not None else None
        if t is not None and ty:
            cit_paid = sum(v for (y, m, g), v in paid.items() if g == "CIT" and y == ty + 1)
            if (ty + 1) in years or cit_paid:
                d = t.c("B47")
                diff = cit_paid - d
                st = "Agrees" if abs(diff) <= max(tol, 0.01 * max(abs(d), abs(cit_paid))) else "Differs"
                cmp_rows.append([f"Annual income tax for {ty} (TOI B47) paid in {ty + 1}", str(ty), str(ty + 1), d, cit_paid, diff, st])
                cmp_tags.append("total" if st == "Agrees" else "warning")
                if st != "Agrees":
                    findings.append(("Annual income tax", str(ty), diff, st))
    except Exception:
        pass

    # ---- summary
    S = [["Source", (folder.name + " folder" if folder else "Info workbook") + f": {len(years)} year(s) {', '.join(map(str, years))}", None, "", "section"]]
    for y in years:
        S.append([f"Total paid {y}", f"{sum(tot.get((y, n)) or 0 for n in names):,.0f} KHR",
                  sum(tot.get((y, n)) or 0 for n in names), "", "even"])
    with_tos = [r[0] for r in month_rows if r[2] > 0]
    zero_tos = [r[0] for r in month_rows if r[2] == 0 and r[1] > 0 and with_tos and with_tos[0] < r[0] < with_tos[-1]]
    if zero_tos:
        S.append(["Months with payments but no salary tax", ", ".join(zero_tos[:12]), None, "Check the TOS return and payment of those months.", "warning"])
    by_label = defaultdict(list)
    for label, mon, diff, st in findings:
        by_label[label].append((mon, diff, st))
    for label, items in by_label.items():
        S.append([f"{label}: paid ≠ declared", f"{len(items)} month(s): " + ", ".join(m for m, _d, _s in items[:8]),
                  sum(d for _m, d, _s in items), "Compare the payment receipts with the returns; late or missing payment means penalties.",
                  "error" if any(s == "Not paid" for _m, _d, s in items) else "warning"])
    if not findings and cmp_rows:
        S.append(["Payments vs returns", "All compared months agree", None, "", "total"])
    an.view("Summary", ["Area", "Result", "Amount (KHR)", "Note / action"], S, [r[4] for r in S], money=[2], total=False,
            desc="Tax paid per the GDT Payment History, by year, and compared with the monthly and annual returns.")
    an.view("By year", ["Tax type", "Group"] + [str(y) for y in years] + ["Change (last year)", "Change %"], summ,
            ["even"] * (len(summ) - 1) + ["total"], money=list(range(2, 3 + len(years))), total=False,
            desc="Total paid per tax type and year.")
    an.view("Monthly detail", an_detail_cols, detail, None, money=list(range(3, 16)), wide=True,
            desc="Every tax type and month, as in the Payment History statement.")
    an.view("By month", ["Month", "Total paid", "Salary tax", "Fringe benefit tax", "WHT", "Prepayment of income tax", "Annual income tax",
                         "VAT", "Other"], month_rows, None, money=list(range(1, 9)), month_col=0, desc="Payments per month, by group.")
    if cmp_rows:
        an.view("Paid vs declared", ["Tax", "Return month", "Payment month", "Declared", "Paid", "Paid − declared", "Status"],
                cmp_rows, cmp_tags, money=[3, 4, 5], month_col=1,
                desc="Payment month = return month or the next month, whichever matches better (taxes are paid the month after).")
    return an
