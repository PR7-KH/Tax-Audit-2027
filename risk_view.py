"""Tax risk finding report: one year of sales, cash-in and A/R, and expense, each compared
across every source that exists (monthly tax returns, annual return TOI, sales listing,
general ledger, trial balance, bank statements, A/R listing).

build_report(company_folder) -> {"year", "rows", "findings", "notes"}; each row is a dict with
level (0 section, 1 item, 2 detail), no, label, amount, base, diff, status, note. Nothing here
touches the GUI, so it is testable. Missing sources are shown as "Not provided", never guessed.
"""
import re
from collections import Counter, defaultdict
from pathlib import Path

import sales_reconciliation as sr
import finance_view as fv

SECTION_ORDER = ("Sales", "Cash", "Expense")


def _folder(root, names):
    return sr._subfolder(root, names)


def _try(notes, label, fn, default=None):
    try:
        return fn()
    except Exception as exc:  # a broken source must not hide the others
        notes.append(f"{label}: {str(exc).splitlines()[0] if str(exc) else type(exc).__name__}")
        return default


def read_tb(folder):
    """Trial balance: net sales (credit - debit on 4xxxx sales accounts) and net expense (debit -
    credit on 5xxxx and 6xxxx excluding income tax), from the last Debit / Credit column pair."""
    files = fv.table_files(folder)
    if not files:
        return None
    cols, rows = fv.read_rows(files[0])
    low = [c.lower() for c in cols]
    code = next((i for i, c in enumerate(low) if "code" in c or c in ("account no", "account no.", "acct")), None)
    name = next((i for i, c in enumerate(low) if "account name" in c or c in ("name", "description", "account")), None)
    debit = [i for i, c in enumerate(low) if "debit" in c]
    credit = [i for i, c in enumerate(low) if "credit" in c]
    if code is None or name is None or not debit or not credit:
        raise ValueError(f"{files[0].name}: need Account code, Account name, Debit and Credit headings")
    d, c = debit[-1], credit[-1]
    out = {"file": files[0].name, "sales": 0.0, "expense": 0.0}
    for r in rows:
        acct, nm = str(r[code] or ""), str(r[name] or "")
        dv, cv = sr._num(r[d]), sr._num(r[c])
        if sr.is_sales_account(acct, nm):
            out["sales"] += cv - dv
        elif acct[:1] in ("5", "6") and not re.search(r"income tax", nm, re.I):
            out["expense"] += dv - cv
    return out


def read_ar_listing(folder):
    files = fv.table_files(folder)
    if not files:
        return None
    cols, rows = fv.read_rows(files[0])
    for token in ("balance", "outstanding", "amount"):
        i = next((i for i, c in enumerate(cols) if token in c.lower()), None)
        if i is not None:
            return {"file": files[0].name, "column": cols[i], "total": sum(sr._num(r[i]) for r in rows if not isinstance(r[i], str) or sr._num(r[i]))}
    raise ValueError(f"{files[0].name}: no Balance / Outstanding / Amount heading")


def read_pr(folder, year):
    """Monthly purchase report (PR) totals of the invoice-value column, by invoice date."""
    import consolidate as cons
    layout = cons.LAYOUTS["PR"]
    total, monthly = 0.0, defaultdict(float)
    for yr, path in cons.files_in(folder):
        if str(yr) != str(year):
            continue
        for r in cons.read_source(path, layout).rows:
            d = sr._date(r[1])
            if d is not None and d.year == year:
                v = sr._num(r[layout["first_col"] - 1] if isinstance(layout, dict) else r[layout.first_col - 1])
                total += v
                monthly[d.month] += v
    return {"total": total, "monthly": dict(monthly)}


def expense_accounts(gl):
    """GL lines on cost of sales (5xxxx) and operating expense (6xxxx) accounts; income tax excluded."""
    cos, opex = 0.0, 0.0
    for r in gl:
        if re.search(r"income tax", r["acct"], re.I):
            continue
        net = r["debit"] - r["credit"]
        if r["code"].startswith("5"):
            cos += net
        elif r["code"].startswith("6"):
            opex += net
    return cos, opex


def match_gl_bank(gl_accounts, bank, acct, cur):
    """Find the GL cash/bank account that belongs to one bank statement account."""
    digits = re.sub(r"\D", "", acct)[-4:] if acct else ""
    if digits:
        hit = [g for g in gl_accounts if digits in g[1]]
        if len(hit) == 1:
            return hit[0]
    alias = {"cnb": "canadia"}.get(bank.lower(), bank.lower())
    cand = [g for g in gl_accounts if alias in g[1].lower()]
    narrowed = [g for g in cand if cur.lower() in g[1].lower()]
    if len(narrowed) == 1:
        return narrowed[0]
    return cand[0] if len(cand) == 1 else None


def build_report(company_folder, year=None, fx=4000.0, tol=1000.0):
    root = Path(company_folder)
    notes = []
    sales_dir = _folder(root, {"SALES", "SALE", "REVENUE"})
    gl_dir = _folder(root, {"GL", "GENERAL LEDGER", "LEDGER"})
    bank_dir = _folder(root, {"BANK", "BANKS", "BANK STATEMENT", "BANK STATEMENTS"})
    sr_dir, toi_dir, pr_dir = _folder(root, {"SR"}), _folder(root, {"TOI"}), _folder(root, {"PR"})
    tb_dir, ar_dir = _folder(root, {"TB", "TRIAL BALANCE"}), _folder(root, {"AR", "A/R", "RECEIVABLE", "RECEIVABLES"})

    listing_files = [p for p in fv.table_files(sales_dir)] if sales_dir else []
    listing_files = [p for p in listing_files if p.suffix.lower() == ".xlsx"]
    if not year and listing_files:
        counts = Counter(r["date"].year for r in (_try(notes, "Sales", lambda: sr.read_sales_listing(listing_files[0], None), []) or []))
        year = max(counts, key=lambda y: (counts[y], y)) if counts else None
    if not year:
        raise ValueError("Cannot tell the tax year: add a sales listing Excel file to the Sales folder")
    listing = _try(notes, "Sales listing", lambda: sr.read_sales_listing(listing_files[0], year)) if listing_files else None
    gl_files = [p for p in fv.table_files(gl_dir) if p.suffix.lower() == ".xlsx"] if gl_dir else []
    gl = _try(notes, "GL", lambda: sr.read_gl(gl_files[0], year)) if gl_files else None
    bank = _try(notes, "Bank", lambda: sr.read_bank(bank_dir, year)[0]) if bank_dir and fv.table_files(bank_dir) else None
    sr_rows = _try(notes, "SR", lambda: sr.read_sr(sr_dir, year)) if sr_dir else None
    toi = _try(notes, "TOI", lambda: sr.read_toi(toi_dir)) if toi_dir else None
    tb = _try(notes, "TB", lambda: read_tb(tb_dir)) if tb_dir else None
    ar_list = _try(notes, "A/R listing", lambda: read_ar_listing(ar_dir)) if ar_dir else None
    pr = _try(notes, "PR", lambda: read_pr(pr_dir, year)) if pr_dir else None
    pdf_total = None
    if sales_dir:
        pdfs = sorted(sales_dir.glob("*.pdf"))
        pdf_total = sr.read_pdf_total(pdfs[0]) if pdfs else None

    k = lambda v: f"{v:,.0f}"
    rows = []

    def add(level, no, label, amount=None, base=None, status="", note="", rel=0.0):
        diff = None if amount is None or base is None else amount - base
        if not status and diff is not None:
            status = "Agrees" if abs(diff) <= max(tol, rel * abs(base)) else "Differs"
        rows.append(dict(level=level, no=no, label=label, amount=amount, base=base, diff=diff, status=status, note=note))

    def missing(level, no, label, why):
        add(level, no, label, None, None, "Not provided", why)

    gl_sales = sum(r["credit"] - r["debit"] for r in gl if sr.is_sales_account(r["code"], r["acct"])) if gl else None
    sales_codes = ", ".join(sorted({c for c, _ in {(r["code"], r["acct"]) for r in gl or [] if sr.is_sales_account(r["code"], r["acct"])}}))

    # ---------------- 1. sales
    add(0, "1", f"Total sales {year}")
    if sr_rows is not None:
        sr_m = [sum(r["amount"] for r in sr_rows if r["date"].month == m) for m in range(1, 13)]
        gl_m = [sum(r["credit"] - r["debit"] for r in gl if r["date"].month == m and sr.is_sales_account(r["code"], r["acct"])) for m in range(1, 13)] if gl else None
        bad = [sr.MONTHS[m] for m in range(12) if gl_m and abs(sr_m[m] - gl_m[m]) > tol]
        add(1, "1.1", "Per monthly tax returns (SR, 12 months)", sum(sr_m), gl_sales,
            note=("Months differing from GL: " + ", ".join(bad)) if bad else "Every month agrees to GL." if gl_m else "")
    else:
        missing(1, "1.1", "Per monthly tax returns (SR, 12 months)", "No SR folder")
    if toi and toi.get("B0") is not None:
        add(1, "1.2", "Per annual return (TOI, operating revenue B0)", toi["B0"], gl_sales,
            note=f"Products B1 {k(toi.get('B1', 0))} + services B3 {k(toi.get('B3', 0))}")
    else:
        missing(1, "1.2", "Per annual return (TOI, operating revenue B0)", "No TOI file or no B0 line")
    if listing is not None:
        add(1, "1.3", "Per sales listing", sum(r["amount"] for r in listing), gl_sales,
            note=f"{len(listing):,} invoices" + (f"; PDF total {k(pdf_total)}" if pdf_total is not None else ""))
    else:
        missing(1, "1.3", "Per sales listing", "No sales listing Excel")
    if gl is not None:
        add(1, "1.4", "Per general ledger", gl_sales, gl_sales, "Base", f"Accounts {sales_codes}")
    else:
        missing(1, "1.4", "Per general ledger", "No GL Excel")
    if tb:
        add(1, "1.5", "Per trial balance", tb["sales"], gl_sales, note=tb["file"])
    else:
        missing(1, "1.5", "Per trial balance", "No TB folder / file" if not tb_dir else "TB file could not be read")

    # ---------------- 2. cash in and A/R
    add(0, "2", "Total cash in and A/R")
    if bank:
        groups = defaultdict(lambda: dict(native=0.0, khr=0.0, receipts=0.0))
        for b in bank:
            label = re.sub(r"\s+", " ", f"{b['bank']} {b.get('acct') or ''} ({b['cur']})").strip()
            g = groups[(b["bank"], b.get("acct") or "", b["cur"], label)]
            v = b["amount"] * (fx if b["cur"] == "USD" else 1)
            g["native"] += b["amount"]; g["khr"] += v
            if b["cls"] in sr.RECEIPT_CLASSES:
                g["receipts"] += v
        gl_bank = {}
        for r in gl or []:
            if r["code"].startswith("10") and re.search(r"bank|cash in", r["acct"], re.I):
                gl_bank.setdefault((r["code"], r["acct"]), 0.0)
                gl_bank[(r["code"], r["acct"])] += r["debit"]
        accts = [(c, n) for c, n in gl_bank]
        total_khr = sum(g["khr"] for g in groups.values())
        add(1, "2.1", f"Bank account cash-in per statements (USD at {fx:,.0f} KHR)", total_khr, None, "Info",
            "Bank credits include transfers and non-sales receipts; compare with GL debits per account")
        for n, ((bk, ac, cur, label), g) in enumerate(sorted(groups.items(), key=lambda x: x[0][3]), start=1):
            m = match_gl_bank(accts, bk, ac, cur)
            base = gl_bank[m] if m else None
            note = f"{cur} {g['native']:,.2f} credits; customer-type receipts {k(g['receipts'])} KHR-eq."
            note += f"; GL {m[0]} {m[1]} debits" if m else "; no matching GL bank account found"
            add(2, f"2.1.{n}", label, g["khr"], base, note=note + ("; 1% allowed for exchange-rate rounding" if cur == "USD" else ""), rel=0.01 if cur == "USD" else 0.0)
    else:
        missing(1, "2.1", "Bank account cash-in per statements", "No bank statement Excel files")
    if gl is not None:
        ar = [r for r in gl if re.search(r"receivable", r["acct"], re.I)]
        if ar:
            bal0 = ar[0]["bal"]
            opening = bal0 - (ar[0]["debit"] - ar[0]["credit"]) if bal0 is not None else None
            inv = sum(r["debit"] for r in ar)
            cash = sum(r["credit"] for r in ar if re.search(r"cash", r["split"], re.I) and not re.search(r"bank", r["split"], re.I))
            via_bank = sum(r["credit"] for r in ar) - cash
            closing = (opening or 0) + inv - cash - via_bank
            add(1, "2.2", "Accounts receivable per GL", closing, None, "Info", ar[0]["acct"])
            if opening is not None:
                add(2, "2.2.1", "Opening balance", opening)
            add(2, "2.2.2", "Add: invoiced (debits)", inv, gl_sales, note="Compared with GL sales")
            add(2, "2.2.3", "Less: collected via bank accounts", via_bank, None, "Info",
                f"{via_bank / (via_bank + cash):.0%} of collections" if via_bank + cash else "")
            add(2, "2.2.4", "Less: collected in cash on hand", cash, None, "Warning" if cash and cash / (via_bank + cash) > 0.2 else "Info",
                f"{cash / (via_bank + cash):.0%} of collections; no bank evidence" if via_bank + cash else "")
            add(2, "2.2.5", "Closing balance (opening + invoiced - collected)", closing, ar[-1]["bal"] if ar[-1]["bal"] is not None else None,
                note="Compared with the running balance on the last GL line")
        else:
            missing(1, "2.2", "Accounts receivable per GL", "No account with 'receivable' in its name")
        if ar_list:
            add(1, "2.3", "Accounts receivable per A/R listing", ar_list["total"], closing if ar else None,
                note=f"Sum of column '{ar_list['column']}' in {ar_list['file']}; check this is the closing balance column")
        else:
            missing(1, "2.3", "Accounts receivable per A/R listing", "No A/R folder / file")
    else:
        missing(1, "2.2", "Accounts receivable per GL", "No GL Excel")

    # ---------------- 3. expense
    add(0, "3", f"Total expense {year}")
    cos, opex = expense_accounts(gl) if gl is not None else (None, None)
    gl_exp = None if gl is None else cos + opex
    ap_credit = sum(r["credit"] for r in gl if re.search(r"payable", r["acct"], re.I) and not re.search(r"tax|nssf|salary", r["acct"], re.I)) if gl else None
    if pr is not None:
        add(1, "3.1", "Per monthly tax returns (PR purchase report, 12 months)", pr["total"], ap_credit,
            note="Different basis: purchases incl. inventory and assets. Compared with GL purchase invoices booked (Accounts Payable credits)")
    else:
        missing(1, "3.1", "Per monthly tax returns (PR purchase report, 12 months)", "No PR folder")
    toi_cos = sum(toi.get(c, 0) for c in ("B4", "B5", "B6")) if toi else None
    toi_opex = toi.get("B22") if toi else None
    if toi and toi_opex is not None:
        add(1, "3.2", "Per annual return (TOI: cost of sales + operating expenses)", toi_cos + toi_opex, gl_exp,
            note="B4+B5+B6 plus B22; income tax B47 excluded")
        add(2, "3.2.1", "Cost of sales (B4+B5+B6) vs GL 5xxxx", toi_cos, cos)
        add(2, "3.2.2", "Operating expenses (B22) vs GL 6xxxx", toi_opex, opex)
    else:
        missing(1, "3.2", "Per annual return (TOI: cost of sales + operating expenses)", "No TOI file or no B22 line")
    if gl is not None:
        add(1, "3.3", "Per general ledger (5xxxx cost of sales + 6xxxx operating expenses)", gl_exp, gl_exp, "Base",
            f"Cost of sales {k(cos)}; operating expenses {k(opex)}; income tax expense excluded")
    else:
        missing(1, "3.3", "Per general ledger", "No GL Excel")
    if tb:
        add(1, "3.4", "Per trial balance", tb["expense"], gl_exp, note=tb["file"])
    else:
        missing(1, "3.4", "Per trial balance", "No TB folder / file" if not tb_dir else "TB file could not be read")

    # ---------------- 4. risk findings
    findings, a_out = [], None
    if listing is not None and gl is not None and bank is not None:
        a_out = sr.analyse(listing, gl, bank, fx, tol, None, None, pdf_total, sr_rows, toi)
        findings = list(a_out["findings"])
    elif gl is not None or listing is not None:
        notes.append("Risk findings that need Sales listing, GL and Bank were skipped")
    add(0, "4", "Risk findings")
    ef = []
    if toi and toi_opex is not None and gl_exp is not None:
        for lab, t, g in (("cost of sales", toi_cos, cos), ("operating expenses", toi_opex, opex)):
            d = t - g
            if abs(d) > tol:
                ef.append(dict(sev="High" if abs(d) > 0.01 * max(abs(g), 1) else "Medium", area="Expense: TOI vs GL",
                               title=f"TOI {lab} differ from GL", amount=d,
                               evidence=f"TOI {k(t)} vs GL {k(g)}.", action="Reconcile the TOI line to GL accounts; look for adjustments or non-deductible items."))
    if gl is not None and not (tb or pr):
        ef.append(dict(sev="Info", area="Expense coverage", title="Expense tested against TOI and GL only", amount=None,
                       evidence="No PR purchase report and no trial balance were provided.", action="Add the PR folder and the TB to complete the expense comparison."))
    findings = sorted(findings + ef, key=lambda f: sr.SEV_ORDER[f["sev"]])
    for n, f in enumerate(findings, start=1):
        add(1, f"4.{n}", f"{f['area']}: {f['title']}", f["amount"], None, f["sev"], f"{f['action']}  |  {f['evidence']}")
    return {"year": year, "rows": rows, "findings": findings, "notes": notes, "analysis": a_out, "fx": fx, "tol": tol}


COLUMNS = ["No", "Item", "Amount (KHR)", "Compared with", "Difference", "Status", "Note"]


def display_rows(report):
    """(rows of text, tag per row) ready for the table view."""
    out, tags = [], []
    fmt = lambda v: "" if v is None else f"{v:,.0f}"
    for r in report["rows"]:
        out.append([r["no"], "    " * r["level"] + r["label"], fmt(r["amount"]), fmt(r["base"]),
                    ("" if r["diff"] is None else f"{r['diff']:+,.0f}"), r["status"], r["note"]])
        s = r["status"]
        tags.append("section" if r["level"] == 0 else "error" if s == "High" else "warning" if s in ("Differs", "Warning", "Medium", "Not provided")
                    else "total" if s in ("Base", "OK", "Agrees") else "even")
    return out, tags


def save_workbook(report, path):
    """Excel file: the sections table first, then (when Sales, GL and Bank exist) the monthly
    comparison, bank cash-in and GL collection sheets from the sales reconciliation."""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    a = report.get("analysis")
    if a is not None:
        sr.write_workbook(path, a, report["year"], report["fx"], report["tol"], a["sr_monthly"], a["annual_sales"], None, None)
        wb = load_workbook(path)
    else:
        wb = Workbook()
        wb.remove(wb.active)
    ws = wb.create_sheet("Tax risk sections", 0)
    ws.append(COLUMNS[:2] + ["Amount (KHR)", "Compared with", "Difference", "Status", "Note"])
    for c in ws[1]:
        c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="BDD7EE")
    fills = {"High": "F8CBAD", "Differs": "FFE699", "Medium": "FFE699", "Warning": "FFE699", "Not provided": "EDEDED",
             "Agrees": "C6E0B4", "OK": "C6E0B4", "Base": "C6E0B4"}
    for r in report["rows"]:
        ws.append([r["no"], "    " * r["level"] + r["label"], r["amount"], r["base"], r["diff"], r["status"], r["note"]])
        row = ws.max_row
        if r["level"] == 0:
            for c in ws[row]:
                c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="DCE9F3")
        elif r["status"] in fills:
            ws.cell(row, 6).fill = PatternFill("solid", fgColor=fills[r["status"]])
        for col in (3, 4, 5):
            ws.cell(row, col).number_format = "#,##0;(#,##0);-"
        ws.cell(row, 7).alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDEFG", (7, 66, 20, 20, 18, 14, 90)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "C2"
    wb.active = 0
    wb.save(path)
