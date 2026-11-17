"""Revenue and Expense analysis menus.

Revenue compares one year of sales per monthly tax return (SR), sales listing, general ledger,
trial balance and the annual return (TOI), month by month and invoice by invoice.
Expense compares purchases per monthly tax return (PR) with the GL (purchase invoices booked
to accounts payable, cost of sales 5xxxx, operating expenses 6xxxx), the TOI expense lines and
the trial balance, adds salary per TOS against the GL salary accounts, and lists PR suppliers.

build_revenue(folder) / build_expense(folder) -> audit_analysis.Analysis. Nothing here touches
the GUI. Missing sources are reported as "Not provided"; nothing is guessed. Amounts are taken
as recorded in each source (KHR in the tax returns).
"""
import re
from collections import Counter, defaultdict
from pathlib import Path

import audit_analysis as aa
import finance_view as fv
import sales_reconciliation as sr

TOL = 1000.0
SALARY_WORDS = re.compile(r"salar|wage|payroll|bonus|ប្រាក់បៀវត្ស|ប្រាក់ឈ្នួល", re.I)


def _try(an, label, fn):
    try:
        return fn()
    except Exception as exc:
        an.add_note(f"{label}: {str(exc).splitlines()[0] if str(exc) else type(exc).__name__}")
        return None


def _xlsx(folder):
    return [p for p in fv.table_files(folder) if p.suffix.lower() == ".xlsx"] if folder else []


def _years_in(folder):
    import consolidate as cons
    if folder is None:
        return []
    return sorted({int(y) for y, _p in cons.files_in(folder) if str(y).isdigit()})


def folders(root):
    root = Path(root)
    f = lambda names: aa.subfolder(root, names)
    return dict(sales=f({"SALES", "SALE", "REVENUE"}), gl=f({"GL", "GENERAL LEDGER", "LEDGER"}),
                sr=f({"SR"}), pr=f({"PR"}), tos=f({"TOS"}), toi=f({"TOI"}), tb=f({"TB", "TRIAL BALANCE"}))


def pick_year(an, d):
    """The tax year: the sales listing's year, else the newest SR / PR year, else the GL's main year."""
    listing = _xlsx(d["sales"])
    if listing:
        rows = _try(an, "Sales listing", lambda: sr.read_sales_listing(listing[0], None)) or []
        c = Counter(r["date"].year for r in rows)
        if c:
            return max(c, key=lambda y: (c[y], y))
    years = _years_in(d["sr"]) or _years_in(d["pr"])
    if years:
        return years[-1]
    gl = _xlsx(d["gl"])
    if gl:
        cols, rows = fv.read_rows(gl[0])
        t = fv.Table(cols, rows)
        c = Counter(m[:4] for m in t.months)
        if c:
            return int(max(c, key=lambda y: (c[y], y)))
    return None


def status(a, b, tol=TOL):
    if a is None or b is None:
        return "Not provided"
    return "Agrees" if abs(a - b) <= tol else "Differs"


def tag_for(st):
    return {"Agrees": "total", "Base": "total", "Differs": "error", "Not provided": "warning"}.get(st, "even")


def summary_view(an, rows, base_label):
    """rows: [(source, amount or None, base or None, note)]"""
    out, tags = [], []
    for src, amount, base, note in rows:
        if amount is None:
            st = "Not provided"
        elif base is None or src.startswith(base_label):
            st = "Base" if src.startswith(base_label) else "Info"
        else:
            st = status(amount, base)
        diff = None if amount is None or base is None or st == "Base" else amount - base
        out.append([src, amount, base if st not in ("Base", "Not provided") else None, diff, st, note])
        tags.append(tag_for(st))
    an.view("Summary", ["Source", "Amount", f"Compared with ({base_label})", "Difference", "Status", "Note"],
            out, tags, money=[1, 2, 3], total=False, desc=f"Year totals compared with the {base_label}.")


# ------------------------------------------------------------------ revenue
def build_revenue(company_folder, year=None, tol=TOL):
    an = aa.Analysis("REVENUE", "Revenue analysis")
    d = folders(company_folder)
    year = year or pick_year(an, d)
    if not year:
        an.add_note("Cannot tell the year: add SR monthly returns, a sales listing or a GL.")
        an.add_summary_only()
        return an
    an.period = str(year)
    listing_f, gl_f = _xlsx(d["sales"]), _xlsx(d["gl"])
    listing = _try(an, "Sales listing", lambda: sr.read_sales_listing(listing_f[0], year)) if listing_f else None
    gl = _try(an, "GL", lambda: sr.read_gl(gl_f[0], year)) if gl_f else None
    srr = _try(an, "SR", lambda: sr.read_sr(d["sr"], year)) if d["sr"] else None
    toi = _try(an, "TOI", lambda: sr.read_toi(d["toi"])) if d["toi"] else None
    import risk_view
    tb = _try(an, "TB", lambda: risk_view.read_tb(d["tb"])) if d["tb"] else None
    gl_sales_rows = [r for r in gl or [] if sr.is_sales_account(r["code"], r["acct"])]
    gl_total = sum(r["credit"] - r["debit"] for r in gl_sales_rows) if gl is not None else None
    sr_total = sum(r["amount"] for r in srr) if srr is not None else None
    li_total = sum(r["amount"] for r in listing) if listing is not None else None
    base_label, base = ("General ledger", gl_total) if gl_total is not None else ("SR monthly returns", sr_total)
    accts = ", ".join(sorted({f"{r['code']} {r['acct']}" for r in gl_sales_rows}))
    summary_view(an, [
        ("SR monthly returns", sr_total, base,
         f"{len(srr):,} invoices" if srr is not None else "No SR folder"),
        ("Sales listing", li_total, base, f"{len(listing):,} invoices" if listing is not None else "No sales listing Excel"),
        ("General ledger (sales accounts)", gl_total, base, accts or ("No GL Excel" if gl is None else "No sales account found")),
        ("Trial balance", tb["sales"] if tb else None, base, tb["file"] if tb else "No TB"),
        ("Annual return TOI (B0 operating revenue)", toi.get("B0") if toi else None, base,
         (f"Products B1 {toi.get('B1', 0):,.0f}; goods B2 {toi.get('B2', 0):,.0f}; services B3 {toi.get('B3', 0):,.0f}" if toi else "No TOI")),
    ], base_label)

    # monthly
    srm, lim, glm, cnt = defaultdict(float), defaultdict(float), defaultdict(float), defaultdict(int)
    for r in srr or []:
        srm[r["date"].month] += r["amount"]; cnt[r["date"].month] += 1
    for r in listing or []:
        lim[r["date"].month] += r["amount"]
    for r in gl_sales_rows:
        glm[r["date"].month] += r["credit"] - r["debit"]
    rows, tags = [], []
    for m in range(1, 13):
        a = srm[m] if srr is not None else None
        b = lim[m] if listing is not None else None
        c = glm[m] if gl is not None else None
        diffs = [x for x in ((a, c), (b, c), (a, b)) if x[0] is not None and x[1] is not None]
        st = "Not provided" if len(diffs) == 0 else ("Agrees" if all(abs(x - y) <= tol for x, y in diffs) else "Differs")
        rows.append([f"{year}-{m:02d}", cnt[m] if srr is not None else None, a, b, c,
                     None if a is None or c is None else a - c, None if b is None or c is None else b - c,
                     None if a is None or b is None else a - b, st])
        tags.append(tag_for(st))
    an.view("Monthly comparison", ["Month", "SR invoices", "SR (monthly tax)", "Sales listing", "GL sales", "SR − GL",
                                   "Listing − GL", "SR − Listing", "Status"], rows, tags, money=[2, 3, 4, 5, 6, 7], month_col=0,
            desc="Sales per month in each source. Differs = more than the tolerance.")

    # invoice match SR vs listing
    if srr is not None and listing is not None:
        key = lambda s: re.sub(r"[\s\-_/]", "", str(s or "")).upper()
        a = defaultdict(list); b = defaultdict(list)
        for r in srr:
            a[key(r["invoice"])].append(r)
        for r in listing:
            b[key(r["invoice"])].append(r)
        rows, tags = [], []
        for k in sorted(set(a) | set(b)):
            ra, rb = a.get(k, []), b.get(k, [])
            sa, sb = sum(x["amount"] for x in ra), sum(x["amount"] for x in rb)
            first = (ra or rb)[0]
            if ra and rb:
                res = "Agrees" if abs(sa - sb) <= tol else "Amount differs"
            else:
                res = "Only in SR" if ra else "Only in sales listing"
            if len(ra) > 1 or len(rb) > 1:
                res += " (invoice repeated)"
            rows.append([first["invoice"], f"{first['date'].year}-{first['date'].month:02d}", str(first["date"]),
                         first["buyer"], sa if ra else None, sb if rb else None, (sa if ra else 0) - (sb if rb else 0), res])
            tags.append("total" if res == "Agrees" else "error" if res.startswith(("Only", "Amount")) else "warning")
        order = {"error": 0, "warning": 1, "total": 2}
        pair = sorted(zip(rows, tags), key=lambda x: (order.get(x[1], 3), x[0][1], str(x[0][0])))
        an.view("Invoice match SR vs listing", ["Invoice", "Month", "Date", "Buyer", "Amount SR", "Amount listing", "Difference", "Result"],
                [p[0] for p in pair], [p[1] for p in pair], money=[4, 5, 6], month_col=1,
                desc="Each invoice in the SR returns matched to the sales listing by invoice number.")

    # by buyer
    buyers = defaultdict(lambda: [0.0, 0.0, 0, None])
    for r in srr or []:
        b = buyers[r["buyer"].upper() or "(blank)"]; b[0] += r["amount"]; b[2] += 1
    for r in listing or []:
        b = buyers[r["buyer"].upper() or "(blank)"]; b[1] += r["amount"]
    party = defaultdict(float)
    for r in gl or []:
        if re.search(r"receivable", r["acct"], re.I):
            party[r["party"].upper()] += r["debit"]
    rows, tags = [], []
    for name, (a, b, n, _) in sorted(buyers.items(), key=lambda x: -max(x[1][0], x[1][1])):
        g = party.get(name)
        rows.append([name, n if srr is not None else None, a if srr is not None else None, b if listing is not None else None,
                     None if srr is None or listing is None else a - b, g,
                     "Not in GL receivables" if gl is not None and g is None else ""])
        tags.append("warning" if rows[-1][-1] or (rows[-1][4] and abs(rows[-1][4]) > tol) else "even")
    an.view("By buyer", ["Buyer", "SR invoices", "SR amount", "Listing amount", "SR − Listing", "GL A/R invoiced", "Note"],
            rows, tags, money=[2, 3, 4, 5], desc="Sales per customer in each source.")
    if srr is not None:
        an.view("SR invoices", ["Date", "Month", "Invoice", "Buyer", "Amount", "File"],
                [[str(r["date"]), f"{r['date'].year}-{r['date'].month:02d}", r["invoice"], r["buyer"], r["amount"], r["file"]] for r in srr],
                None, money=[4], month_col=1, desc="Every invoice line of the SR monthly returns.")
    if gl is not None:
        an.view("GL sales lines", ["Date", "Month", "Account code", "Account", "Type", "Counterparty", "Debit", "Credit", "Net sales"],
                [[str(r["date"]), f"{r['date'].year}-{r['date'].month:02d}", r["code"], r["acct"], r["type"], r["party"],
                  r["debit"], r["credit"], r["credit"] - r["debit"]] for r in gl_sales_rows],
                None, money=[6, 7, 8], month_col=1, desc="GL lines on sales accounts (4xxxx sales / revenue).")
    return an


# ------------------------------------------------------------------ expense
def read_pr_rows(folder, year):
    """PR monthly purchase lines with supplier and amounts, labels taken from the return headings."""
    import consolidate as cons
    layout = cons.LAYOUTS["PR"]
    out = []
    for yr, path in cons.files_in(folder):
        if str(yr) != str(year):
            continue
        src = cons.read_source(path, layout)
        labels = aa.two_row_labels(src.headers, src.merges)
        low = [l.casefold() for l in labels]
        pick = lambda tests, default: next((i for i, l in enumerate(low[:layout.first_col - 1]) if any(t in l for t in tests)), default)
        supplier = pick(("ឈ្មោះ", "name", "supplier"), 5 if layout.first_col > 6 else None)
        tin = pick(("អត្តសញ្ញាណកម្ម", "tin", "vat no", "vattin"), None)
        vat = next((i for i in range(layout.first_col, layout.last_col) if i < len(low) and re.search(r"អាករ|vat", low[i])), None)
        for r in src.rows:
            dte = aa.to_date(r[1]) if len(r) > 1 else None
            if dte is None or dte.year != int(year):
                continue
            out.append({"date": dte, "invoice": str(r[2] or "").strip() if len(r) > 2 else "",
                        "supplier": str(r[supplier] or "").strip() if supplier is not None and supplier < len(r) else "",
                        "tin": str(r[tin] or "").strip() if tin is not None and tin < len(r) else "",
                        "amount": aa.num(r[layout.first_col - 1]), "vat": aa.num(r[vat]) if vat is not None and vat < len(r) else None,
                        "file": path.name})
    return out


def build_expense(company_folder, year=None, tol=TOL):
    an = aa.Analysis("EXPENSE", "Expense analysis")
    d = folders(company_folder)
    year = year or pick_year(an, d)
    if not year:
        an.add_note("Cannot tell the year: add PR monthly returns, a sales listing or a GL.")
        an.add_summary_only()
        return an
    an.period = str(year)
    gl_f = _xlsx(d["gl"])
    gl = _try(an, "GL", lambda: sr.read_gl(gl_f[0], year)) if gl_f else None
    pr = _try(an, "PR", lambda: read_pr_rows(d["pr"], year)) if d["pr"] else None
    toi = _try(an, "TOI", lambda: sr.read_toi(d["toi"])) if d["toi"] else None
    import risk_view
    tb = _try(an, "TB", lambda: risk_view.read_tb(d["tb"])) if d["tb"] else None
    tos = None
    if d["tos"]:
        import tos_analysis
        files = _try(an, "TOS", lambda: tos_analysis.read_tos(d["tos"])[0])
        if files is not None:
            tos = defaultdict(float)
            for y, m, _p, _l, _c, recs in files:
                if y == int(year):
                    tos[m] += sum(r["salary"] for r in recs)

    exp_rows = [r for r in gl or [] if r["code"][:1] in ("5", "6") and not re.search(r"income tax", r["acct"], re.I)]
    ap_rows = [r for r in gl or [] if re.search(r"payable", r["acct"], re.I) and not re.search(r"tax|nssf|salary", r["acct"], re.I)]
    sal_rows = [r for r in exp_rows if SALARY_WORDS.search(r["acct"])]
    cos = sum(r["debit"] - r["credit"] for r in exp_rows if r["code"].startswith("5")) if gl is not None else None
    opex = sum(r["debit"] - r["credit"] for r in exp_rows if r["code"].startswith("6")) if gl is not None else None
    gl_exp = None if gl is None else cos + opex
    ap = sum(r["credit"] for r in ap_rows) if gl is not None else None
    gl_sal = sum(r["debit"] - r["credit"] for r in sal_rows) if gl is not None else None
    pr_total = sum(r["amount"] for r in pr) if pr is not None else None
    toi_cos = sum(toi.get(c, 0) for c in ("B4", "B5", "B6")) if toi and any(c in toi for c in ("B4", "B5", "B6")) else None
    toi_opex = toi.get("B22") if toi else None
    tos_total = sum(tos.values()) if tos is not None else None

    rows, tags = [], []
    def line(src, amount, base, note, info=False):
        st = "Not provided" if amount is None else ("Info" if info or base is None else status(amount, base, tol))
        rows.append([src, amount, None if amount is None else base, None if amount is None or base is None else amount - base, st, note])
        tags.append(tag_for(st))
    line("PR purchases (monthly tax returns)", pr_total, ap, "Compared with GL purchase invoices booked (A/P credits). "
         "PR includes inventory and assets." if pr is not None else "No PR folder")
    line("GL expense total (5xxxx + 6xxxx)", gl_exp, None, "Income tax excluded" if gl is not None else "No GL Excel", info=True)
    line("TOI cost of sales (B4+B5+B6) vs GL 5xxxx", toi_cos, cos, "Annual return" if toi else "No TOI")
    line("TOI operating expenses (B22) vs GL 6xxxx", toi_opex, opex, "Annual return" if toi else "No TOI")
    line("Trial balance expense vs GL", tb["expense"] if tb else None, gl_exp, tb["file"] if tb else "No TB")
    line("Salary per TOS returns vs GL salary accounts", tos_total, gl_sal,
         ("GL accounts: " + ", ".join(sorted({f"{r['code']} {r['acct']}" for r in sal_rows}))) if sal_rows else
         ("No GL account named salary / wage / payroll" if gl is not None else "No GL Excel") if tos is not None else "No TOS folder")
    an.view("Summary", ["Comparison", "Amount", "Compared with", "Difference", "Status", "Note"], rows, tags, money=[1, 2, 3], total=False,
            desc="Expense totals per source. Differs = more than the tolerance.")

    # monthly
    prm, apm, exm, slm = defaultdict(float), defaultdict(float), defaultdict(float), defaultdict(float)
    for r in pr or []:
        prm[r["date"].month] += r["amount"]
    for r in ap_rows:
        apm[r["date"].month] += r["credit"]
    for r in exp_rows:
        exm[r["date"].month] += r["debit"] - r["credit"]
    for r in sal_rows:
        slm[r["date"].month] += r["debit"] - r["credit"]
    rows, tags = [], []
    for m in range(1, 13):
        a = prm[m] if pr is not None else None
        b = apm[m] if gl is not None else None
        t = tos.get(m, 0.0) if tos is not None else None
        s = slm[m] if gl is not None else None
        st = [x for x in ((a, b), (t, s)) if x[0] is not None and x[1] is not None]
        res = "Not provided" if not st else ("Agrees" if all(abs(x - y) <= tol for x, y in st) else "Differs")
        rows.append([f"{year}-{m:02d}", a, b, None if a is None or b is None else a - b, exm[m] if gl is not None else None,
                     t, s, None if t is None or s is None else t - s, res])
        tags.append(tag_for(res))
    an.view("Monthly comparison", ["Month", "PR purchases", "GL A/P credits", "PR − GL A/P", "GL expense 5/6xxxx",
                                   "TOS salary", "GL salary", "TOS − GL salary", "Status"], rows, tags,
            money=[1, 2, 3, 4, 5, 6, 7], month_col=0, desc="Purchases and salary per month in each source.")

    if gl is not None:
        acc = defaultdict(lambda: [0.0, 0.0])
        for r in exp_rows:
            acc[(r["code"], r["acct"])][0] += r["debit"]; acc[(r["code"], r["acct"])][1] += r["credit"]
        total = sum(dbt - crd for dbt, crd in acc.values()) or 1
        rows = [[c, n, "Cost of sales (TOI B4–B6)" if c.startswith("5") else "Operating expense (TOI B22)", dbt, crd, dbt - crd,
                 f"{(dbt - crd) / total:.1%}", "Salary" if SALARY_WORDS.search(n) else ""]
                for (c, n), (dbt, crd) in sorted(acc.items())]
        an.view("By GL account", ["Account code", "Account", "TOI line", "Debit", "Credit", "Net expense", "Share", "Note"],
                rows, None, money=[3, 4, 5], desc="Expense per GL account (5xxxx and 6xxxx; income tax excluded).")
    if pr is not None:
        sup = defaultdict(lambda: {"n": 0, "amt": 0.0, "vat": 0.0, "tin": "", "d": []})
        for r in pr:
            s = sup[r["supplier"].upper() or "(blank)"]
            s["n"] += 1; s["amt"] += r["amount"]; s["vat"] += r["vat"] or 0; s["tin"] = s["tin"] or r["tin"]; s["d"].append(r["date"])
        gl_party = {r["party"].upper() for r in gl or [] if r["party"]}
        rows, tags = [], []
        for name, s in sorted(sup.items(), key=lambda x: -x[1]["amt"]):
            found = name in gl_party or any(name[:12] and name[:12] in p for p in gl_party)
            note = "" if gl is None else ("" if found else "Supplier not found in GL counterparties")
            if not s["tin"]:
                note = (note + "; " if note else "") + "No TIN"
            rows.append([name, s["tin"], s["n"], s["amt"], s["vat"], str(min(s["d"])), str(max(s["d"])), note])
            tags.append("warning" if note else "even")
        an.view("PR by supplier", ["Supplier", "TIN", "Invoices", "Amount", "VAT", "First date", "Last date", "Note"],
                rows, tags, money=[3, 4], desc="Purchases per supplier in the PR returns, matched to GL counterparties by name.")
        an.view("PR purchases", ["Date", "Month", "Invoice", "Supplier", "TIN", "Amount", "VAT", "File"],
                [[str(r["date"]), f"{r['date'].year}-{r['date'].month:02d}", r["invoice"], r["supplier"], r["tin"], r["amount"], r["vat"], r["file"]]
                 for r in pr], None, money=[5, 6], month_col=1, desc="Every purchase line of the PR monthly returns.")
    if gl is not None:
        an.view("GL expense lines", ["Date", "Month", "Account code", "Account", "Type", "Counterparty", "Debit", "Credit", "Net"],
                [[str(r["date"]), f"{r['date'].year}-{r['date'].month:02d}", r["code"], r["acct"], r["type"], r["party"],
                  r["debit"], r["credit"], r["debit"] - r["credit"]] for r in exp_rows],
                None, money=[6, 7, 8], month_col=1, desc="GL lines on 5xxxx and 6xxxx accounts.")
    return an
