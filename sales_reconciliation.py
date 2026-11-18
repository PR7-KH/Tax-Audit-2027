"""Total-sales reconciliation and tax-risk findings.

Compares one year of sales across independent sources:
  * sales listing (Excel, optionally PDF total)   * general ledger (Excel)
  * bank statements (Excel, one or more banks)    * SR monthly tax returns and the annual
    tax return (typed into the Inputs sheet or given in an inputs JSON file)

Usage:
    python3 sales_reconciliation.py <company folder> <output.xlsx> [--year 2025]
        [--fx 4000] [--tolerance 1000] [--inputs inputs.json]
Expected sub-folders: GL, Bank (or BANKS), Sales.  Each reader detects columns by their
heading text, so other layouts work as long as the headings are recognisable. Anything
it cannot recognise raises a clear error rather than guessing.
"""
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).resolve().parent))

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2, "Info": 3, "OK": 4}
RECEIPT_CLASSES = ("Foreign remittance", "In-house transfer received", "Cheque deposit")


# ------------------------------------------------------------------ helpers
def _num(v):
    if isinstance(v, bool):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.replace(",", "").replace("៛", "").strip()
        try:
            return float(t)
        except ValueError:
            return 0.0
    return 0.0


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str):
        m = re.match(r"\s*(\d{4})-(\d{2})-(\d{2})", v)
        if m:
            return date(int(m[1]), int(m[2]), int(m[3]))
        m = re.match(r"\s*(\d{2})-(\d{2})-(\d{4})", v)
        if m:
            return date(int(m[3]), int(m[2]), int(m[1]))
    return None


def _find_header(ws, wanted, scan=20):
    """Return (row_index, {key: column_index}) for the first row that holds every wanted key.
    wanted maps key -> list of lowercase tokens; a heading matches when it contains any token."""
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=scan, values_only=True), start=1):
        heads = [str(c).strip().lower() if c is not None else "" for c in row]
        found = {}
        for key, tokens in wanted.items():
            for ci, h in enumerate(heads):
                if h and any(t(h) if callable(t) else t in h for t in tokens):
                    found[key] = ci
                    break
        if all(k in found for k in wanted if not k.startswith("?")):
            return i, found
    return None, None


def _sheet(wb, token):
    for ws in wb.worksheets:
        if token in ws.title.lower():
            return ws
    return wb.worksheets[0]


def _subfolder(root, names):
    for d in Path(root).iterdir():
        if d.is_dir() and d.name.upper() in names and not d.name.startswith("."):
            return d
    return None


def _xlsx_in(folder):
    return sorted(p for p in Path(folder).rglob("*.xlsx")
                  if not any(part.startswith((".", "~$")) for part in p.relative_to(folder).parts))


# ------------------------------------------------------------------ readers
def read_sales_listing(path, year):
    wb = load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        hi, cols = _find_header(ws, {
            "date": ["date"], "invoice": ["invoice", "declaration"],
            "buyer": ["buyer", "customer"], "amount": ["total amount", "amount"],
            "?month": ["month"]})
        if hi:
            break
    else:
        raise ValueError(f"{Path(path).name}: cannot find Date / Invoice / Buyer / Total Amount headings")
    rows = []
    for r in ws.iter_rows(min_row=hi + 1, values_only=True):
        d = _date(r[cols["date"]])
        if d is None or (year and d.year != year):
            continue
        rows.append({"date": d, "invoice": str(r[cols["invoice"]] or "").strip(),
                     "buyer": str(r[cols["buyer"]] or "").strip(), "amount": _num(r[cols["amount"]])})
    wb.close()
    return rows


def read_pdf_total(pdf_path):
    """Best effort: the grand total printed on the last page of a sales-listing PDF."""
    try:
        import pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            text = pdf.pages[-1].extract_text() or ""
        m = re.findall(r"Total\s+([\d,]+(?:\.\d+)?)", text)
        return float(m[-1].replace(",", "")) if m else None
    except Exception:
        return None


def read_sr(folder, year):
    """Monthly SR (sale report) workbooks, read with the same reader the consolidation uses."""
    import consolidate as cons
    rows = []
    for yr, path in cons.files_in(folder):
        if str(yr) != str(year):
            continue
        for r in cons.read_source(path, cons.LAYOUTS["SR"]).rows:
            d = _date(r[1])
            if d is not None:
                rows.append({"date": d, "invoice": str(r[2] or "").strip(), "buyer": str(r[6] or "").strip(),
                             "amount": _num(r[7]), "file": path.name})
    return rows


def read_toi(folder):
    """Annual return: operating revenue (code B0), product sales (B1), goods (B2), services (B3)
    from Step 5, and the total of sales to related persons from Step 14 (None when not found)."""
    files = [p for p in _xlsx_in(folder)]
    if not files:
        return {}
    wb = load_workbook(files[0], read_only=True, data_only=True)
    out = {"file": files[0].name}
    for ws in wb.worksheets:
        t = ws.title.lower().replace(" ", "")
        if t == "step5":
            for r in ws.iter_rows(values_only=True):
                if len(r) > 2 and isinstance(r[1], str) and re.fullmatch(r"B\d+", r[1]):
                    out[r[1]] = _num(r[2])
        if t == "step14":
            inside, total = False, 0.0
            for r in ws.iter_rows(values_only=True):
                first = str(r[0] or "").strip()
                if first.startswith("ក."):
                    inside = True
                    continue
                if inside and first.startswith("សម្គាល់"):
                    break
                if inside and len(r) > 4:
                    total += _num(r[4])
            out["related_sales"] = total if inside else None
    wb.close()
    return out


def read_gl(path, year):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = _sheet(wb, "transaction")
    hi, c = _find_header(ws, {
        "code": ["account code"], "acct": ["account name"], "date": ["date"], "type": ["transaction type", "type"],
        "party": ["counterparty", lambda h: h == "name"], "split": ["split account"],
        "debit": ["debit"], "credit": ["credit"], "?bal": ["running balance"],
        "?num": [lambda h: h in ("num", "no", "no.", "ref", "ref.", "reference", "doc no", "doc no.", "document no", "document number",
                                 "voucher", "voucher no") or "invoice" in h or "វិក្កយបត្រ" in h or h.startswith(("ref ", "reference"))]})
    if not hi:
        raise ValueError(f"{Path(path).name}: cannot find Account Code / Date / Debit / Credit headings")
    rows = []
    for r in ws.iter_rows(min_row=hi + 1, values_only=True):
        d = _date(r[c["date"]])
        if d is None or d.year != year:
            continue
        rows.append({"code": str(r[c["code"]] or ""), "acct": str(r[c["acct"]] or ""), "date": d,
                     "type": str(r[c["type"]] or ""), "party": str(r[c["party"]] or "").strip(),
                     "split": str(r[c["split"]] or ""), "debit": _num(r[c["debit"]]), "credit": _num(r[c["credit"]]),
                     "bal": _num(r[c["?bal"]]) if "?bal" in c and r[c["?bal"]] is not None else None,
                     "num": str(r[c["?num"]] or "").strip() if "?num" in c else ""})
    wb.close()
    return rows


def is_sales_account(code, name):
    n = name.lower()
    return (code.startswith("4") and re.search(r"sale|revenue|service income|sewing", n)
            and not re.search(r"interest|other income|discount|return", n))


def classify_credit(desc, currency, hint=""):
    d, h = desc.upper(), hint.lower()
    if "payroll funding" in h:
        return "Payroll funding (not sales)"
    if "INTEREST" in d or "interest" in h:
        return "Interest"
    if "NATIONAL BANK OF CAMBODIA" in d or "BAKONG" in d:
        return "NBC / government fund transfer"
    if "IN HOUSE" in d or "IN-HOUSE" in d:
        if currency != "USD" or "EXCHANGE" in d or re.search(r"TO\s+KHR", d):
            return "Own-account exchange transfer"
        return "In-house transfer received"
    if "B/O" in d:
        return "Foreign remittance"
    if d.startswith("CHQ"):
        return "Cheque deposit"
    return "Other credit"


def payer_name(desc):
    m = re.search(r"B/O\s+(.+?)(?:,|$)", desc, re.I)
    return re.sub(r"\s+", " ", m[1]).strip().upper() if m else ""


def read_bank(folder, year):
    rows, files = [], []
    for p in _xlsx_in(folder):
        wb = load_workbook(p, read_only=True, data_only=True)
        ws = _sheet(wb, "transaction")
        hi, c = _find_header(ws, {
            "date": [lambda h: h.startswith("date") or h.startswith("posting date")],
            "in": ["money in", "credit", "bank receipts"], "desc": ["description", "bank notes"],
            "?cur": ["currency"], "?class": ["classification", "category"], "?acct": ["account"]})
        if not hi:
            raise ValueError(f"{p.name}: cannot find Date / Money In / Description headings on a Transactions sheet")
        heads = [str(x or "") for x in next(ws.iter_rows(min_row=hi, max_row=hi, values_only=True))]
        cur_default = "KHR" if "khr" in heads[c["in"]].lower() else "USD"
        bank = re.split(r"[_ ]", p.stem)[0]
        n = 0
        for r in ws.iter_rows(min_row=hi + 1, values_only=True):
            d = _date(r[c["date"]])
            if d is None or d.year != year:
                continue
            amt = _num(r[c["in"]])
            if amt <= 0:
                continue
            cur = str(r[c["?cur"]]).upper() if "?cur" in c and r[c["?cur"]] else cur_default
            desc = str(r[c["desc"]] or "")
            hint = str(r[c["?class"]] or "") if "?class" in c else ""
            acct = str(r[c["?acct"]] or "").strip() if "?acct" in c else ""
            rows.append({"bank": bank, "acct": acct, "date": d, "cur": cur, "amount": amt, "desc": desc,
                         "cls": classify_credit(desc, cur, hint), "payer": payer_name(desc)})
            n += 1
        files.append((p.name, bank, n))
        wb.close()
    return rows, files


# ------------------------------------------------------------------ analysis
def _norm_tokens(name):
    stop = {"CO", "LTD", "LIMITED", "INC", "THE", "AND", "COMPANY", "SA", "AB", "BV", "CORP", "PTE"}
    return {t for t in re.findall(r"[A-Z0-9]+", name.upper()) if t not in stop and len(t) > 2}


def analyse(listing, gl, bank, fx, tol, sr_monthly=None, annual_sales=None, pdf_total=None, sr_rows=None, toi=None):
    mon = lambda dct, m: dct.get(m, 0.0)
    list_m, gl_m = defaultdict(float), defaultdict(float)
    for r in listing:
        list_m[r["date"].month] += r["amount"]
    sales_accts = sorted({(r["code"], r["acct"]) for r in gl if is_sales_account(r["code"], r["acct"])})
    for r in gl:
        if is_sales_account(r["code"], r["acct"]):
            gl_m[r["date"].month] += r["credit"] - r["debit"]
    ar = [r for r in gl if re.search(r"receivable", r["acct"], re.I)]
    coll_bank, coll_cash, payers = defaultdict(float), defaultdict(float), defaultdict(float)
    for r in ar:
        if r["credit"] > 0:
            payers[r["party"] or "(blank)"] += r["credit"]
            (coll_cash if re.search(r"cash", r["split"], re.I) and not re.search(r"bank", r["split"], re.I)
             else coll_bank)[r["date"].month] += r["credit"]
    bank_m, bank_cls, bank_payers = defaultdict(float), defaultdict(float), defaultdict(float)
    cur_totals = defaultdict(float)
    for b in bank:
        khr = b["amount"] * (fx if b["cur"] == "USD" else 1)
        bank_cls[b["cls"]] += khr
        cur_totals[(b["bank"], b["cur"])] += b["amount"]
        if b["cls"] in RECEIPT_CLASSES:
            bank_m[b["date"].month] += khr
        if b["payer"]:
            bank_payers[b["payer"]] += khr
    out = dict(list_m=list_m, gl_m=gl_m, coll_bank=coll_bank, coll_cash=coll_cash, bank_m=bank_m,
               bank_cls=bank_cls, payers=payers, bank_payers=bank_payers, sales_accts=sales_accts,
               cur_totals=cur_totals, ar_open=0.0)
    out["sr_rows"] = sr_rows
    if sr_rows:
        sr_monthly = [sum(r["amount"] for r in sr_rows if r["date"].month == m) for m in range(1, 13)]
    if toi and toi.get("B0") is not None:
        annual_sales = toi["B0"]
    out["sr_monthly"], out["annual_sales"], out["toi"] = sr_monthly, annual_sales, toi or {}
    out["findings"] = findings(listing, gl, ar, out, fx, tol, sr_monthly, annual_sales, pdf_total)
    return out


def findings(listing, gl, ar, a, fx, tol, sr_monthly, annual_sales, pdf_total):
    F = []
    add = lambda sev, area, title, amount, evidence, action: F.append(
        dict(sev=sev, area=area, title=title, amount=amount, evidence=evidence, action=action))
    k = lambda v: f"{v:,.0f}"
    list_tot, gl_tot = sum(a["list_m"].values()), sum(a["gl_m"].values())
    # 1. sales listing vs GL
    bad = [m for m in range(1, 13) if abs(a["list_m"].get(m, 0) - a["gl_m"].get(m, 0)) > tol]
    if abs(list_tot - gl_tot) <= tol and not bad:
        add("OK", "Sales listing vs GL", "Sales listing agrees to GL sales in total and in every month", 0,
            f"Listing {k(list_tot)} = GL {k(gl_tot)} ({', '.join(c for c, _ in a['sales_accts'])}).",
            "None. Use this as the verified sales base for the tax comparisons.")
    else:
        add("High", "Sales listing vs GL", "Sales listing does not agree to GL sales", list_tot - gl_tot,
            f"Listing {k(list_tot)} vs GL {k(gl_tot)}; months with differences: {', '.join(MONTHS[m-1] for m in bad) or 'none (total only)'}.",
            "Obtain the missing/extra invoices and ask management to explain each month.")
    if pdf_total is not None:
        ok = abs(pdf_total - list_tot) <= tol
        add("OK" if ok else "High", "Sales listing PDF vs Excel",
            "Sales listing PDF total agrees to the Excel" if ok else "Sales listing PDF total differs from the Excel",
            pdf_total - list_tot, f"PDF total {k(pdf_total)} vs Excel {k(list_tot)}.",
            "None." if ok else "Re-convert the PDF; check dropped or edited rows.")
    # 2. SR and annual return
    if sr_monthly:
        diffs = [m for m in range(1, 13) if abs(sr_monthly[m - 1] - a["gl_m"].get(m, 0)) > tol]
        tot = sum(sr_monthly)
        sev = "OK" if abs(tot - gl_tot) <= tol and not diffs else "High"
        add(sev, "SR monthly returns vs GL",
            "SR monthly tax returns agree to GL sales" if sev == "OK" else "SR monthly tax returns differ from GL sales",
            tot - gl_tot, f"SR total {k(tot)} vs GL {k(gl_tot)}; months differing: {', '.join(MONTHS[m-1] for m in diffs) or 'none'}.",
            "None." if sev == "OK" else "Reconcile month by month: timing, exports, credit notes, unreported invoices.")
    else:
        add("Info", "SR monthly returns vs GL", "NOT TESTED: SR monthly totals not provided", None,
            "Enter the 12 SR monthly totals in the Inputs sheet (or inputs.json).", "Provide SR totals.")
    if annual_sales is not None:
        sev = "OK" if abs(annual_sales - gl_tot) <= tol else "High"
        add(sev, "Annual tax return vs GL",
            "Annual return sales agree to GL sales" if sev == "OK" else "Annual return sales differ from GL sales",
            annual_sales - gl_tot, f"Annual return {k(annual_sales)} vs GL {k(gl_tot)}."
            + (f" TOI products {k(a['toi'].get('B1', 0))} + services {k(a['toi'].get('B3', 0))}; file {a['toi'].get('file', '')}." if a.get("toi") else ""),
            "None." if sev == "OK" else "Obtain explanation; assess under/over-declared revenue.")
    else:
        add("Info", "Annual tax return vs GL", "NOT TESTED: annual return sales not provided", None,
            "Enter the annual return sales figure in the Inputs sheet.", "Provide TOI sales figure.")
    sr_rows = a.get("sr_rows")
    if sr_rows:
        lc = Counter((r["invoice"], round(r["amount"])) for r in listing)
        sc = Counter((r["invoice"], round(r["amount"])) for r in sr_rows)
        only_sr, only_list = sc - lc, lc - sc
        if only_sr or only_list:
            det = {(r["invoice"], round(r["amount"])): r for r in sr_rows}
            ex = [f"{i} {det[(i, v)]['date']} {det[(i, v)]['buyer'][:30]} {k(v)} (in SR, not in sales listing)" for (i, v) in only_sr]
            ex += [f"{i} {k(v)} (in sales listing, not in SR)" for (i, v) in only_list]
            diff = sum(v * n for (_, v), n in only_sr.items()) - sum(v * n for (_, v), n in only_list.items())
            add("High", "SR vs books at invoice level", "SR contains invoices that are not in the sales listing / GL", diff,
                "; ".join(ex) + ". If the invoice is a valid 2025 sale, revenue in the books and annual return is understated; "
                f"income tax at 20% would be about {k(max(diff, 0) * 0.20)} KHR.",
                "Obtain the invoice and customs declaration. Establish whether it is a real shipment (record it and amend the return) "
                "or a cancelled / duplicate declaration (correct the SR).")
        else:
            add("OK", "SR vs books at invoice level", "Every SR invoice agrees to the sales listing", 0, "Same invoice numbers and amounts.", "None.")
    toi = a.get("toi") or {}
    if sr_rows and toi.get("B0") is not None:
        sr_tot = sum(r["amount"] for r in sr_rows)
        if abs(sr_tot - toi["B0"]) > tol:
            add("High", "SR vs annual return (TOI)", "Sales declared in the 12 SR returns exceed sales in the annual return", sr_tot - toi["B0"],
                f"SR 12 months {k(sr_tot)} vs TOI {k(toi['B0'])}.", "See the invoice-level finding above; reconcile and amend whichever filing is wrong.")
    if toi.get("related_sales") is not None and toi["related_sales"] == 0:
        add("Medium", "Related-party disclosure", "TOI Step 14 declares no sales to related persons", 0,
            "All customer collections come from third-party payers rather than the invoiced buyers (see payer finding).",
            "Confirm whether the paying parties are related to the company or its shareholders; if so Step 14 is incomplete.")
    # 3. duplicate invoice numbers
    seen = defaultdict(list)
    for r in listing:
        if r["invoice"]:
            seen[r["invoice"]].append(r)
    dups = {i: v for i, v in seen.items() if len(v) > 1}
    if dups:
        amt = sum(r["amount"] for v in dups.values() for r in v)
        add("Medium", "Sales listing integrity", f"{len(dups)} invoice / customs numbers are used more than once", amt,
            "; ".join(f"{i}: " + " | ".join(f"{r['date']} {r['buyer'][:25]} {k(r['amount'])}" for r in v)
                      for i, v in list(dups.items())[:5]),
            "Check the invoices / customs declarations; the same number should not cover different buyers or dates.")
    # 4. payer vs buyer
    buyer_tokens = [_norm_tokens(r["buyer"]) for r in listing]
    all_buyer_tokens = set().union(*buyer_tokens) if buyer_tokens else set()
    pay_tot = sum(a["payers"].values())
    unmatched = {p: v for p, v in a["payers"].items() if not (_norm_tokens(p) & all_buyer_tokens)}
    um_tot = sum(unmatched.values())
    if pay_tot and um_tot / pay_tot > 0.5:
        top = sorted(unmatched.items(), key=lambda x: -x[1])[:3]
        add("High", "Payer vs buyer", "Customers pay through a party that is not the invoiced buyer", um_tot,
            f"{um_tot / pay_tot:.1%} of receivable collections come from " + ", ".join(f"{n} ({k(v)})" for n, v in top)
            + "; none of these names appear among the buyers on the sales listing.",
            "Establish who the payers are (related party / agent / group treasury), obtain agreements, and test transfer-pricing and revenue-completeness.")
    # 5. cash collections
    cash, bnk = sum(a["coll_cash"].values()), sum(a["coll_bank"].values())
    if cash > 0:
        share = cash / (cash + bnk)
        add("High" if share > 0.2 else "Medium", "Cash collections",
            f"{share:.0%} of customer collections are booked through cash on hand", cash,
            f"Receivable credited {k(cash)} against cash on hand vs {k(bnk)} against bank accounts. No bank statement supports the cash part. "
            "Largest cash months: " + ", ".join(f"{MONTHS[m-1]} {k(v)}" for m, v in sorted(a["coll_cash"].items(), key=lambda x: -x[1])[:3]) + ".",
            "Vouch cash receipts to receipt vouchers, bank deposit slips and withdrawals; look for round-tripping between bank and cash.")
    # 6. bank coverage / monthly agreement
    gaps = []
    for m in range(1, 13):
        g, b = a["coll_bank"].get(m, 0), a["bank_m"].get(m, 0)
        if g - b > max(tol, 0.01 * g):
            gaps.append((m, g, b))
    if gaps:
        tot = sum(g - b for _, g, b in gaps)
        add("High", "Bank vs GL collections", "GL bank collections are not supported by the bank statements provided", tot,
            "; ".join(f"{MONTHS[m-1]}: GL {k(g)} vs bank {k(b)}" for m, g, b in gaps),
            "Request the missing statements (the statement register shows which months are absent) and re-run.")
    else:
        add("OK", "Bank vs GL collections", "GL bank collections agree to bank receipts every month (within 1%)", 0,
            f"Compared at {fx:,.0f} KHR/USD.", "None.")
    inh = a["bank_cls"].get("In-house transfer received", 0)
    if inh:
        add("Medium", "Bank classification", "Bank credits described as in-house transfers are treated as customer receipts", inh,
            f"{k(inh)} KHR-equivalent received by in-house transfer from another account; the GL books bank receipts to receivables.",
            "Identify the source account holder. A transfer from an own, director or related account is not a sale.")
    nbc = a["bank_cls"].get("NBC / government fund transfer", 0)
    if nbc:
        vr = sum(r["credit"] for r in gl if re.search(r"vat refund", r["acct"], re.I))
        tied = abs(nbc - vr) <= tol
        add("Info", "Bank classification", "National Bank of Cambodia transfers are not sales" + (" (agree to GL VAT refund)" if tied else ""),
            nbc, f"NBC/Bakong credits {k(nbc)}; GL VAT refund credits {k(vr)}.",
            "Agree to the VAT refund claim and General Department of Taxation refund notice.")
    pf = a["bank_cls"].get("Payroll funding (not sales)", 0)
    if pf:
        add("Info", "Bank classification", "Large bank credits labelled payroll funding are excluded from sales", pf,
            f"{k(pf)} KHR-equivalent.", "Confirm the source (shareholder/related funding) and its accounting treatment.")
    cov = sum(a["bank_cls"].get(c, 0) for c in RECEIPT_CLASSES)
    if list_tot:
        add("Info", "Bank vs sales", "Bank receipts cover only part of annual sales", cov - list_tot,
            f"Receipt-type bank credits {k(cov)} = {cov / list_tot:.0%} of sales {k(list_tot)} (FX {fx:,.0f}). "
            "Bank cash-in is a reasonableness check, not a tie-out: statements may be incomplete and cash collections exist.",
            "Do not conclude under-declaration from this line alone; read it with the cash and coverage findings.")
    return sorted(F, key=lambda f: SEV_ORDER[f["sev"]])


# ------------------------------------------------------------------ workbook
HEAD = PatternFill("solid", fgColor="BDD7EE")
INPUT = PatternFill("solid", fgColor="FFFFCC")
SEVFILL = {"High": "F8CBAD", "Medium": "FFE699", "Low": "E2EFDA", "Info": "DDEBF7", "OK": "C6E0B4"}


def _hdr(ws, row, labels):
    for i, t in enumerate(labels, start=1):
        c = ws.cell(row, i, t)
        c.font, c.fill = Font(bold=True), HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")


def write_workbook(path, a, year, fx, tol, sr_monthly, annual_sales, listing, gl_sales_rows):
    wb = Workbook()
    f = wb.active
    f.title = "Findings"
    inp = wb.create_sheet("Inputs")
    inp["A1"], inp["A2"], inp["A3"] = "Tax year", "FX rate (KHR per USD)", "Tolerance (KHR)"
    inp["B1"], inp["B2"], inp["B3"] = year, fx, tol
    inp["A5"] = "Annual tax return: total sales (KHR)"
    inp["B5"] = annual_sales
    inp["A7"], inp["B7"] = "Month", "SR monthly total sales (KHR)"
    for c in ("B1", "B2", "B3", "B5"):
        inp[c].fill = INPUT
    for i in range(12):
        inp.cell(8 + i, 1, MONTHS[i])
        c = inp.cell(8 + i, 2, sr_monthly[i] if sr_monthly else None)
        c.fill, c.number_format = INPUT, "#,##0"
    inp["A21"] = "Yellow cells are inputs. FX and tolerance are assumptions used by the Python checks; SR and annual cells feed the live formulas."
    inp.column_dimensions["A"].width, inp.column_dimensions["B"].width = 42, 30
    # monthly comparison with live formulas
    m = wb.create_sheet("Monthly Comparison")
    _hdr(m, 1, ["Month", "Sales listing", "GL sales", "SR return (input)", "Listing - GL", "SR - GL",
                "Collected via bank (GL)", "Bank receipts (statements)", "Bank gap", "Collected via cash (GL)", "Flag"])
    for i in range(12):
        r, mm = 2 + i, i + 1
        m.cell(r, 1, MONTHS[i])
        m.cell(r, 2, round(a["list_m"].get(mm, 0)))
        m.cell(r, 3, round(a["gl_m"].get(mm, 0)))
        m.cell(r, 4, f'=IF(Inputs!B{8+i}="","",Inputs!B{8+i})')
        m.cell(r, 5, f"=B{r}-C{r}")
        m.cell(r, 6, f'=IF(D{r}="","",D{r}-C{r})')
        m.cell(r, 7, round(a["coll_bank"].get(mm, 0)))
        m.cell(r, 8, round(a["bank_m"].get(mm, 0)))
        m.cell(r, 9, f"=G{r}-H{r}")
        m.cell(r, 10, round(a["coll_cash"].get(mm, 0)))
        m.cell(r, 11, f'=IF(ABS(E{r})>Inputs!$B$3,"Listing/GL differ",IF(AND(F{r}<>"",ABS(N(F{r}))>Inputs!$B$3),"SR/GL differ",IF(I{r}>MAX(Inputs!$B$3,0.01*G{r}),"Bank gap","")))')
    t = 14
    m.cell(t, 1, "Total").font = Font(bold=True)
    for col in "BCDEFGHIJ":
        m[f"{col}{t}"] = f"=SUM({col}2:{col}13)"
        m[f"{col}{t}"].font = Font(bold=True)
    m.cell(16, 1, "Annual tax return")
    m["B16"] = '=IF(Inputs!B5="","",Inputs!B5)'
    m.cell(17, 1, "Annual return - GL")
    m["B17"] = '=IF(B16="","",B16-C14)'
    for row in m.iter_rows(min_row=2, max_row=17, min_col=2, max_col=10):
        for c in row:
            c.number_format = "#,##0;(#,##0);-"
    m.column_dimensions["A"].width = 20
    for col in "BCDEFGHIJ":
        m.column_dimensions[col].width = 20
    m.column_dimensions["K"].width = 22
    m.freeze_panes = "B2"
    # findings
    _hdr(f, 1, ["#", "Severity", "Area", "Finding", "Amount (KHR)", "Evidence", "Suggested audit procedure"])
    for i, x in enumerate(a["findings"], start=1):
        f.append([i, x["sev"], x["area"], x["title"], x["amount"], x["evidence"], x["action"]])
        f.cell(i + 1, 2).fill = PatternFill("solid", fgColor=SEVFILL[x["sev"]])
        f.cell(i + 1, 5).number_format = "#,##0;(#,##0);-"
        for c in range(1, 8):
            f.cell(i + 1, c).alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDEFG", (5, 10, 26, 48, 18, 80, 60)):
        f.column_dimensions[col].width = w
    f.freeze_panes = "A2"
    # bank classification
    b = wb.create_sheet("Bank Cash-In")
    _hdr(b, 1, ["Classification", "KHR equivalent", "Counted as customer receipt?"])
    for k_, v in sorted(a["bank_cls"].items(), key=lambda x: -x[1]):
        b.append([k_, round(v), "Yes" if k_ in RECEIPT_CLASSES else "No"])
    b.append([])
    _hdr(b, b.max_row + 1, ["Bank / currency", "Total credits (original currency)", ""])
    for (bk, cur), v in sorted(a["cur_totals"].items()):
        b.append([f"{bk} {cur}", round(v, 2)])
    b.append([])
    _hdr(b, b.max_row + 1, ["Remitter named on bank advice", "KHR equivalent", ""])
    for p, v in sorted(a["bank_payers"].items(), key=lambda x: -x[1]):
        b.append([p, round(v)])
    b.column_dimensions["A"].width, b.column_dimensions["B"].width = 50, 28
    for row in b.iter_rows(min_col=2, max_col=2):
        for c in row:
            c.number_format = "#,##0.##"
    # payers
    p = wb.create_sheet("GL Collections")
    _hdr(p, 1, ["Payer on receivable credits (GL)", "KHR"])
    for n, v in sorted(a["payers"].items(), key=lambda x: -x[1]):
        p.append([n, round(v)])
    p.column_dimensions["A"].width, p.column_dimensions["B"].width = 50, 22
    for row in p.iter_rows(min_col=2, max_col=2):
        for c in row:
            c.number_format = "#,##0"
    wb.active = 0
    wb.save(path)


# ------------------------------------------------------------------ driver
def run(company_folder, output, year=None, fx=4000.0, tol=1000.0, inputs=None):
    root = Path(company_folder)
    inputs = inputs or {}
    sales_dir, gl_dir, bank_dir = (_subfolder(root, {"SALES", "SALE", "REVENUE"}),
                                   _subfolder(root, {"GL", "GENERAL LEDGER", "LEDGER"}),
                                   _subfolder(root, {"BANK", "BANKS", "BANK STATEMENT", "BANK STATEMENTS"}))
    for name, d in (("Sales", sales_dir), ("GL", gl_dir), ("Bank", bank_dir)):
        if d is None:
            raise ValueError(f"{name} folder not found in {root}")
    xl = _xlsx_in(sales_dir)
    xl = [p for p in xl if ".pdf-extractions" not in p.parts]
    if not xl:
        raise ValueError("No sales listing .xlsx found; convert the PDF first")
    if not year:  # newest year that has the most sales rows
        counts = Counter(r["date"].year for r in read_sales_listing(xl[0], None))
        if not counts:
            raise ValueError("Sales listing has no dated rows")
        year = max(counts, key=lambda y: (counts[y], y))
    listing = read_sales_listing(xl[0], year)
    pdfs = sorted(sales_dir.glob("*.pdf"))
    pdf_total = read_pdf_total(pdfs[0]) if pdfs else None
    gl_files = _xlsx_in(gl_dir)
    if not gl_files:
        raise ValueError("No GL .xlsx found")
    gl = read_gl(gl_files[0], year)
    bank, _files = read_bank(bank_dir, year)
    sr_dir, toi_dir = _subfolder(root, {"SR"}), _subfolder(root, {"TOI"})
    sr_rows = read_sr(sr_dir, year) if sr_dir else None
    toi = read_toi(toi_dir) if toi_dir else None
    a = analyse(listing, gl, bank, fx, tol, inputs.get("sr_monthly"), inputs.get("annual_sales"), pdf_total, sr_rows, toi)
    a["year"] = year
    write_workbook(output, a, year, fx, tol, a["sr_monthly"], a["annual_sales"], listing, None)
    return a


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("company_folder", type=Path)
    ap.add_argument("output", type=Path)
    ap.add_argument("--year", type=int, default=None, help="default: year with most rows in the sales listing")
    ap.add_argument("--fx", type=float, default=4000.0, help="KHR per USD for bank credits (default 4000)")
    ap.add_argument("--tolerance", type=float, default=1000.0, help="KHR rounding tolerance")
    ap.add_argument("--inputs", type=Path, help='JSON: {"sr_monthly":[12 numbers],"annual_sales":N}')
    args = ap.parse_args()
    data = json.loads(args.inputs.read_text()) if args.inputs else {}
    res = run(args.company_folder, args.output, args.year, args.fx, args.tolerance, data)
    for x in res["findings"]:
        print(f"[{x['sev']}] {x['area']}: {x['title']}")
