"""TOI (annual tax on income return) analysis: key risk checks, ratio analysis and a cash-flow
statement, all from the two periods printed in the TOI (current year N and prior year N-1).

The TOI is read by item code: A0..A52 balance sheet (step 4), B0..B48 income statement (step 5),
C1..C20 production cost (step 6), D1..D9 cost of goods (step 7). Step 10 (the fixed-asset /
depreciation table) gives the assets bought in the year when its columns can be recognised;
otherwise purchases are derived from the change in fixed assets plus depreciation.

build_ratios(folder) / build_cashflow(folder) -> audit_analysis.Analysis.  Nothing here touches the GUI.
"""
import re
from pathlib import Path

import audit_analysis as aa

CODE = re.compile(r"^[ABCD]\d{1,2}$")
TOL = 1000.0     # KHR rounding accepted in identity checks


# ------------------------------------------------------------------ reading the TOI
def toi_file(root):
    d = aa.subfolder(root, {"TOI"})
    if d is None:
        return None
    files = sorted(p for p in d.rglob("*.xlsx") if not any(part.startswith((".", "~$")) for part in p.relative_to(d).parts))
    return files[0] if files else None


def _sheet(wb, n):
    for ws in wb.worksheets:
        if ws.title.lower().replace(" ", "") == f"step{n}":
            return ws
    return None


def read_items(path):
    """{code: {"desc", "cur", "pri"}} from every step sheet (code in column B, N in C, N-1 in D)."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    items = {}
    try:
        for ws in wb.worksheets:
            if not ws.title.lower().replace(" ", "").startswith("step"):
                continue
            for r in ws.iter_rows(values_only=True):
                if len(r) < 3 or not isinstance(r[1], str):
                    continue
                code = r[1].strip().upper()
                if CODE.match(code) and code not in items:
                    items[code] = {"desc": re.sub(r"\s+", " ", str(r[0] or "")).strip(),
                                   "cur": aa.num(r[2]), "pri": aa.num(r[3]) if len(r) > 3 else 0.0}
    finally:
        wb.close()
    return items


ADD_WORDS = r"ទិញ|បន្ថែម|កើន|ទទួលបាន|addition|acquisi|purchase|increase|bought"
DISPOSE_WORDS = r"លក់|ថយ|ដកចេញ|disposal|sold|decrease|write"
TOTAL = re.compile(r"^\s*(grand\s+)?total\b|^\s*សរុប", re.I)


def read_step10(path):
    """Assets bought (and disposed of) in the year from step 10. Returns a dict with the column used,
    the rows read and the totals, or None when the sheet / columns are not recognised."""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = _sheet(wb, 10)
        if ws is None:
            return None
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    width = max((len(r) for r in rows), default=0)
    # heading text per column: everything written above the first row that holds numbers
    first_num = next((i for i, r in enumerate(rows) if sum(isinstance(v, (int, float)) and not isinstance(v, bool) and v != 0
                                                         for v in r[1:]) >= 1 and i > 0), None)
    if first_num is None:
        return None
    heads = [" ".join(re.sub(r"\s+", " ", str(rows[i][c])) for i in range(first_num)
                      if c < len(rows[i]) and isinstance(rows[i][c], str)) for c in range(width)]
    add = next((c for c in range(1, width) if re.search(ADD_WORDS, heads[c], re.I)), None)
    disp = next((c for c in range(1, width) if c != add and re.search(DISPOSE_WORDS, heads[c], re.I)), None)
    if add is None:
        return {"found": False, "heads": heads, "rows": []}
    detail, total_row = [], None
    for r in rows[first_num:]:
        label = " ".join(str(v) for v in r[:2] if isinstance(v, str)).strip()
        a = aa.num(r[add]) if add < len(r) else 0.0
        d = aa.num(r[disp]) if disp is not None and disp < len(r) else 0.0
        if not label and not a and not d:
            continue
        if TOTAL.search(label):
            total_row = (label, a, d)
            continue
        detail.append((label, a, d))
    adds = total_row[1] if total_row and total_row[1] else sum(x[1] for x in detail)
    disps = total_row[2] if total_row and total_row[2] else sum(x[2] for x in detail)
    return {"found": True, "add_col": aa.col_letter(add) + ": " + heads[add].strip(),
            "disp_col": (aa.col_letter(disp) + ": " + heads[disp].strip()) if disp is not None else "",
            "additions": adds, "disposals": disps, "rows": detail, "total_row": total_row}


class Toi:
    """Item values with helpers: t.c("A21") current, t.p("A21") prior, sums with t.c("A21", "A22")."""

    def __init__(self, items):
        self.items = items

    def c(self, *codes):
        return sum(self.items.get(k, {}).get("cur", 0.0) for k in codes)

    def p(self, *codes):
        return sum(self.items.get(k, {}).get("pri", 0.0) for k in codes)

    def d(self, *codes):
        return self.c(*codes) - self.p(*codes)

    def has(self, code):
        return code in self.items

    def desc(self, code):
        return self.items.get(code, {}).get("desc", "")


def load(root):
    path = toi_file(root)
    if path is None:
        return None, None, "No TOI folder / Excel file"
    items = read_items(path)
    if not items:
        return None, path, f"{path.name}: no item codes (A0, B0 ...) found in the step sheets"
    return Toi(items), path, ""


def year_of(path):
    m = re.search(r"20\d\d", path.stem) if path else None
    return int(m.group(0)) if m else None


# ------------------------------------------------------------------ measures used by several views
def liabilities(t, f):
    return f("A37") + f("A42")


def cost_of_sales(t, f):
    return f("B4", "B5", "B6")


def cash(t, f):
    return f("A21", "A22")


def div(a, b):
    return None if not b else a / b


# ------------------------------------------------------------------ key risk checks
def key_checks(t):
    """[(check, current, prior, flag, finding (EN), suggested procedure)] for the dashboard-style view."""
    out = []

    def add(name, cur, pri, hit, finding, action):
        out.append([name, cur, pri, "Check" if hit else "OK", finding if hit else "", action if hit else ""])

    profit_c, profit_p = t.c("B48"), t.p("B48")
    add("Profit made but cash decreased", t.d("A21", "A22"), profit_c,
        profit_c > 0 and t.d("A21", "A22") < 0,
        f"Profit after tax {profit_c:,.0f} but cash on hand + bank fell by {-t.d('A21', 'A22'):,.0f}.",
        "Trace where the profit went: A/R increase, stock, asset purchases, loans to related parties, dividends. "
        "Check bank statements and unrecorded withdrawals.")
    add("Profit made, A/R and cash both not increasing", t.d("A18"), t.d("A21", "A22"),
        profit_c > 0 and t.d("A18") <= 0 and t.d("A21", "A22") <= 0,
        "Profit is not visible in cash or receivables.", "Ask for the cash-flow / use of profit; check inventory and related-party balances.")
    debt_c, debt_p = liabilities(t, t.c), liabilities(t, t.p)
    add("Debt greater than assets", debt_c, t.c("A0"), debt_c > t.c("A0") > 0 or (debt_c > 0 and t.c("A0") == 0),
        f"Liabilities {debt_c:,.0f} exceed total assets {t.c('A0'):,.0f} (negative equity {t.c('A29'):,.0f}).",
        "Going-concern risk; check who funds the company (related-party loans A38 / A45), loan agreements and interest WHT.")
    add("Negative equity", t.c("A29"), t.p("A29"), t.c("A29") < 0, "Owner's equity is negative.",
        "Confirm accumulated losses are real (expenses not inflated); capital injection or debt conversion expected.")
    gp_c = t.c("B7")
    add("Gross profit negative", gp_c, t.p("B7"), gp_c < 0 and t.c("B0") > 0,
        f"Cost of sales {cost_of_sales(t, t.c):,.0f} is higher than revenue {t.c('B0'):,.0f}.",
        "Selling below cost: check pricing to related parties (transfer pricing), unrecorded sales, stock valuation.")
    cos_c = cost_of_sales(t, t.c)
    add("Operating expenses greater than cost of sales", t.c("B22"), cos_c, cos_c > 0 and t.c("B22") > cos_c,
        f"Operating expenses {t.c('B22'):,.0f} exceed cost of sales {cos_c:,.0f}.",
        "Review the large expense lines (B23–B41) for non-deductible, personal or undocumented items.")
    add("Operating expenses greater than revenue", t.c("B22"), t.c("B0"), t.c("B0") > 0 and t.c("B22") > t.c("B0"),
        "Operating expenses exceed total revenue.", "Check how the business is financed and whether revenue is fully declared.")
    add("No cost of sales but revenue declared", cos_c, t.c("B0"), t.c("B0") > 0 and cos_c == 0 and (t.c("A14", "A15", "A16") > 0),
        "Stock is held but no cost of sales is declared.", "Check stock movement (steps 6 and 7) and cost recognition.")
    add("Loss in both years", t.c("B46"), t.p("B46"), t.c("B46") < 0 and t.p("B46") < 0,
        "The company reports a loss in N and N-1.", "Continuous losses: check the funding source, related-party pricing and expense validity.")
    add("Loss but cash increased", t.d("A21", "A22"), t.c("B46"), t.c("B46") < 0 and t.d("A21", "A22") > 0,
        "Cash grew although the company made a loss.", "Find the source of cash: new loans, capital, unrecorded revenue.")
    rp = t.d("A38", "A45")
    add("Related-party funding increased", rp, t.p("A38", "A45"), rp > 0,
        f"Loans / payables to related parties rose by {rp:,.0f}.",
        "Check loan agreements, interest charged (WHT 14 % non-resident / 15 % resident) and interest limitation.")
    rev_g, stock_g = t.d("B0"), t.d("A14", "A15", "A16", "A17")
    add("Stock up while revenue down", stock_g, rev_g, stock_g > 0 and rev_g < 0,
        "Inventory increased while sales fell.", "Stock count, slow-moving stock, possible unrecorded sales.")
    ar_g = div(t.d("A18"), t.p("A18"))
    rv_g = div(t.d("B0"), t.p("B0"))
    add("A/R growing faster than revenue", t.d("A18"), t.d("B0"),
        ar_g is not None and rv_g is not None and ar_g > rv_g + 0.2 and t.d("A18") > 0,
        f"A/R +{ar_g:.0%} vs revenue {rv_g:+.0%}." if ar_g is not None and rv_g is not None else "",
        "Check collection after year end, credit notes and bad debts; revenue may be booked without cash.")
    capex = t.d("A1") + t.c("B36", "C9", "C12")
    add("Fixed assets bought in the year", capex, t.c("B36", "C9", "C12"), capex > TOL,
        f"Estimated purchases of long-term assets {capex:,.0f} (change in A1 + depreciation).",
        "Vouch purchases (step 10), VAT input claimed, WHT on construction / installation services.")
    add("Balance sheet does not balance", t.c("A0"), t.c("A28"), abs(t.c("A0") - t.c("A28")) > TOL,
        f"Assets {t.c('A0'):,.0f} vs equity + liabilities {t.c('A28'):,.0f}.", "Ask for a corrected TOI.")
    add("Profit in balance sheet differs from income statement", t.c("A36"), t.c("B48"),
        t.has("A36") and t.has("B48") and abs(t.c("A36") - t.c("B48")) > TOL,
        "A36 (profit of the year in equity) is not equal to B48.", "Reconcile; check prior-year adjustments booked to equity.")
    for code, close in (("C1", "A14"), ("D1", "A15"), ("C18", "A16"), ("C15", "A17")):
        if t.has(code) and (t.c(code) or t.p(close)):
            add(f"Opening stock {code} vs prior-year closing {close}", t.c(code), t.p(close), abs(t.c(code) - t.p(close)) > TOL,
                f"Opening stock {t.c(code):,.0f} differs from last year's closing {t.p(close):,.0f}.",
                "Stock should roll forward; ask for the explanation (write-off, reclassification, error).")
    for code, close in (("C5", "A14"), ("D8", "A15"), ("C19", "A16"), ("C16", "A17")):
        if t.has(code) and (t.c(code) or t.c(close)):
            add(f"Closing stock {code} vs balance sheet {close}", t.c(code), t.c(close), abs(t.c(code) - t.c(close)) > TOL,
                "Closing stock in the cost table differs from the balance sheet.", "Reconcile steps 6 / 7 with step 4.")
    min_tax = 0.01 * t.c("B0")
    add("Income tax below 1 % of turnover (minimum tax)", t.c("B47"), min_tax,
        t.c("B0") > 0 and t.c("B47") + TOL < min_tax,
        f"Income tax {t.c('B47'):,.0f} is below 1 % of revenue {min_tax:,.0f}.",
        "Check minimum tax / prepayment of profit tax, exemptions (QIP) and the tax computation.")
    for code, what in (("B29", "Entertainment"), ("B32", "Donations"), ("B37", "Provisions"), ("B40", "Unrealised FX loss")):
        add(f"{what} expense ({code})", t.c(code), t.p(code), t.c(code) > 0, f"{what} {t.c(code):,.0f} declared.",
            "Usually non-deductible or limited: confirm the add-back in the tax computation.")
    for code, what in (("B27", "Rental"), ("B33", "Management / technical services"), ("B34", "Royalty"), ("B43", "Interest (resident)"),
                       ("B44", "Interest (non-resident)")):
        add(f"{what} expense ({code}) – WHT", t.c(code), t.p(code), t.c(code) > 0, f"{what} {t.c(code):,.0f} declared.",
            "Compare with WHT returns (WHT tab): rental 10 %, services / interest 15 %, non-resident 14 %.")
    return out


# ------------------------------------------------------------------ ratios
def ratio_rows(t):
    """[(group, ratio, formula, N, N-1, change, flag, comment)]; ratios are fractions or days."""
    rows = []

    def r(group, name, formula, fn, kind="pct", warn=None, comment=""):
        cur, pri = fn(t.c), fn(t.p)
        chg = None if cur is None or pri is None else cur - pri
        flag = ""
        if warn and cur is not None:
            msg = warn(cur, pri)
            if msg:
                flag = msg
        rows.append([group, name, formula, cur, pri, chg, kind, flag, comment])

    rev = lambda f: f("B0")
    r("Profitability", "Gross profit margin", "B7 / B0", lambda f: div(f("B7"), rev(f)),
      warn=lambda c, p: "Negative gross margin" if c < 0 else ("Margin fell by more than 10 points" if p is not None and c < p - 0.10 else ""),
      comment="How much of each riel of sales is left after cost of sales.")
    r("Profitability", "Operating profit margin", "B42 / B0", lambda f: div(f("B42"), rev(f)),
      warn=lambda c, p: "Operating loss" if c < 0 else "")
    r("Profitability", "Net profit margin", "B48 / B0", lambda f: div(f("B48"), rev(f)),
      warn=lambda c, p: "Net loss" if c < 0 else "")
    r("Profitability", "Cost of sales to revenue", "(B4+B5+B6) / B0", lambda f: div(f("B4", "B5", "B6"), rev(f)),
      warn=lambda c, p: "Cost of sales above revenue" if c > 1 else "")
    r("Profitability", "Operating expenses to revenue", "B22 / B0", lambda f: div(f("B22"), rev(f)),
      warn=lambda c, p: "Expenses above revenue" if c > 1 else ("Up by more than 10 points" if p is not None and c > p + 0.10 else ""))
    r("Profitability", "Operating expenses to cost of sales", "B22 / (B4+B5+B6)", lambda f: div(f("B22"), f("B4", "B5", "B6")), kind="x",
      warn=lambda c, p: "Expenses bigger than cost of sales" if c > 1 else "")
    r("Profitability", "Return on assets (ROA)", "B48 / A0", lambda f: div(f("B48"), f("A0")))
    r("Profitability", "Return on equity (ROE)", "B48 / A29", lambda f: div(f("B48"), f("A29")) if f("A29") > 0 else None,
      comment="Not shown when equity is negative.")
    r("Profitability", "Effective tax rate", "B47 / B46", lambda f: div(f("B47"), f("B46")) if f("B46") > 0 else None,
      warn=lambda c, p: "Below the 20 % rate" if c < 0.19 else "")
    r("Profitability", "Income tax to revenue (minimum tax 1 %)", "B47 / B0", lambda f: div(f("B47"), rev(f)),
      warn=lambda c, p: "Below 1 % of turnover" if c < 0.0099 else "")
    r("Growth", "Revenue growth", "B0 N / B0 N-1 − 1", lambda f: None, kind="pct")
    rows[-1][3] = div(t.d("B0"), t.p("B0"))
    rows[-1][4] = None
    r("Growth", "Operating expense growth", "B22 N / B22 N-1 − 1", lambda f: None)
    rows[-1][3] = div(t.d("B22"), t.p("B22"))
    rows[-1][4] = None
    if rows[-1][3] is not None and rows[-2][3] is not None and rows[-1][3] > rows[-2][3] + 0.2:
        rows[-1][7] = "Expenses growing much faster than revenue"
    r("Liquidity", "Current ratio", "A13 / A42", lambda f: div(f("A13"), f("A42")), kind="x",
      warn=lambda c, p: "Below 1: short-term debts exceed current assets" if c < 1 else "")
    r("Liquidity", "Quick ratio", "(A13 − stock A14..A17) / A42", lambda f: div(f("A13") - f("A14", "A15", "A16", "A17"), f("A42")), kind="x",
      warn=lambda c, p: "Below 0.5" if c < 0.5 else "")
    r("Liquidity", "Cash ratio", "(A21+A22) / A42", lambda f: div(f("A21", "A22"), f("A42")), kind="x")
    r("Liquidity", "Cash on hand to total cash", "A21 / (A21+A22)", lambda f: div(f("A21"), f("A21", "A22")),
      warn=lambda c, p: "Most cash held as cash on hand: check cash count" if c > 0.5 else "")
    r("Solvency", "Debt to assets", "(A37+A42) / A0", lambda f: div(f("A37", "A42"), f("A0")), kind="pct",
      warn=lambda c, p: "Debt greater than assets" if c > 1 else ("High leverage (over 80 %)" if c > 0.8 else ""))
    r("Solvency", "Debt to equity", "(A37+A42) / A29", lambda f: div(f("A37", "A42"), f("A29")) if f("A29") > 0 else None, kind="x",
      warn=lambda c, p: "Over 3 times equity" if c > 3 else "", comment="Not shown when equity is negative.")
    r("Solvency", "Equity ratio", "A29 / A0", lambda f: div(f("A29"), f("A0")),
      warn=lambda c, p: "Negative equity" if c < 0 else "")
    r("Solvency", "Related-party debt share", "(A38+A45) / (A37+A42)", lambda f: div(f("A38", "A45"), f("A37", "A42")),
      warn=lambda c, p: "Mostly funded by related parties" if c > 0.5 else "")
    r("Solvency", "Interest cover", "B42 / (B43+B44)", lambda f: div(f("B42"), f("B43", "B44")), kind="x",
      warn=lambda c, p: "Interest not covered by operating profit" if c < 1 else "")
    r("Efficiency", "A/R days", "A18 / B0 × 365", lambda f: (div(f("A18"), rev(f)) or 0) * 365 if rev(f) else None, kind="days",
      warn=lambda c, p: "Over 90 days" if c > 90 else "")
    r("Efficiency", "Inventory days", "Stock A14..A17 / cost of sales × 365",
      lambda f: div(f("A14", "A15", "A16", "A17"), f("B4", "B5", "B6")) * 365 if f("B4", "B5", "B6") else None, kind="days",
      warn=lambda c, p: "Over 180 days" if c > 180 else "")
    r("Efficiency", "Payables days", "A46 / (cost of sales + B22) × 365",
      lambda f: div(f("A46"), f("B4", "B5", "B6", "B22")) * 365 if f("B4", "B5", "B6", "B22") else None, kind="days",
      warn=lambda c, p: "Over 180 days: unpaid or fictitious payables?" if c > 180 else "")
    r("Efficiency", "Asset turnover", "B0 / A0", lambda f: div(rev(f), f("A0")), kind="x")
    return rows


# ------------------------------------------------------------------ cash flow
def cashflow_lines(t, step10=None):
    """Indirect-method cash flow for year N. Returns [(section, code(s), label, amount, note, level)]."""
    L = []

    def line(section, codes, label, amount, note="", level=1):
        L.append([section, codes, label, amount, note, level])

    dep = t.c("B36", "C9", "C12")
    line("A. Operating activities", "B48", "Profit / (loss) after tax", t.c("B48"))
    line("A. Operating activities", "B36+C9+C12", "Add: depreciation and amortisation", dep, "Non-cash expense")
    gain = t.c("B16") - t.c("B38")
    if gain:
        line("A. Operating activities", "B16−B38", "Less: gain / add: loss on sale of long-term assets", -gain, "Moved to investing")
    wc = [("A14+A15+A16+A17", ("A14", "A15", "A16", "A17"), "(Increase) / decrease in inventories", -1),
          ("A18", ("A18",), "(Increase) / decrease in trade receivables", -1),
          ("A19", ("A19",), "(Increase) / decrease in other receivables", -1),
          ("A20", ("A20",), "(Increase) / decrease in prepaid expenses", -1),
          ("A23+A24+A25", ("A23", "A24", "A25"), "(Increase) / decrease in tax credits (prepayment, VAT)", -1),
          ("A26+A27", ("A26", "A27"), "(Increase) / decrease in other current assets", -1),
          ("A45", ("A45",), "Increase / (decrease) in payables to related parties", 1),
          ("A46", ("A46",), "Increase / (decrease) in other payables", 1),
          ("A47", ("A47",), "Increase / (decrease) in unearned revenue", 1),
          ("A48", ("A48",), "Increase / (decrease) in accruals and other current liabilities", 1),
          ("A40+A49", ("A40", "A49"), "Increase / (decrease) in provisions", 1),
          ("A50+A51", ("A50", "A51"), "Increase / (decrease) in taxes payable", 1),
          ("A52", ("A52",), "Increase / (decrease) in exchange differences (liabilities)", 1)]
    for label_codes, codes, label, sign in wc:
        v = sign * t.d(*codes)
        if v:
            line("A. Operating activities", label_codes, label, v)
    op = sum(x[3] for x in L)
    line("A. Operating activities", "", "Net cash from operating activities", op, level=0)

    # investing: fixed assets A1 (A2..A10, A12) and investments A11
    fa_change = t.d("A1") - t.d("A11")
    reval = t.d("A33")
    derived = fa_change + dep - reval
    inv = []
    if step10 and step10.get("found") and step10.get("additions"):
        capex = step10["additions"]
        inv.append(("Step 10", "Purchase of long-term assets (TOI step 10)", -capex, step10.get("add_col", "")))
        other = capex - derived
        if abs(other) > TOL:
            inv.append(("A1", "Disposals / other movements in long-term assets (balancing)", other,
                        f"Derived purchases {derived:,.0f} vs step 10 {capex:,.0f}; includes disposals at book value"))
    else:
        inv.append(("A1+dep", "Purchase of long-term assets (derived: change in A1 + depreciation)", -derived,
                    "Step 10 not recognised: net of disposals"))
    if gain:
        inv.append(("B16−B38", "Gain / (loss) on sale of long-term assets", gain, "Part of the sale proceeds"))
    if t.d("A11"):
        inv.append(("A11", "(Investment) / withdrawal in other companies", -t.d("A11"), ""))
    for codes, label, v, note in inv:
        line("B. Investing activities", codes, label, v, note)
    inv_total = sum(x[2] for x in inv)
    line("B. Investing activities", "", "Net cash from investing activities", inv_total, level=0)

    fin = [("A30+A31", "Capital contributed / (withdrawn)", t.d("A30", "A31")),
           ("A38", "Loans from related parties (long-term)", t.d("A38")),
           ("A39", "Bank and other loans (long-term)", t.d("A39")),
           ("A41", "Other long-term liabilities", t.d("A41")),
           ("A43+A44", "Overdraft and current portion of loans", t.d("A43", "A44"))]
    other_eq = t.d("A29") - t.d("A30", "A31") - reval - t.c("B48")
    if abs(other_eq) > 0.5:
        fin.append(("A29", "Dividends / other equity movements (derived)", other_eq,))
    for row in fin:
        if row[2]:
            line("C. Financing activities", row[0], row[1], row[2],
                 "Negative = paid out (dividends, prior-year adjustments)" if row[0] == "A29" else "")
    fin_total = sum(r[2] for r in fin)
    line("C. Financing activities", "", "Net cash from financing activities", fin_total, level=0)

    net = op + inv_total + fin_total
    line("D. Cash", "", "Net increase / (decrease) in cash", net, level=0)
    line("D. Cash", "A21+A22 N-1", "Cash at beginning of year (cash on hand + bank)", t.p("A21", "A22"))
    line("D. Cash", "A21+A22 N", "Cash at end of year per balance sheet", t.c("A21", "A22"))
    diff = t.p("A21", "A22") + net - t.c("A21", "A22")
    line("D. Cash", "", "Difference (should be 0)", diff,
         "OK" if abs(diff) <= TOL else "Balance sheet lines do not reconcile: check A0 = A28 in both years", level=0)
    return L, dict(op=op, inv=inv_total, fin=fin_total, net=net, diff=diff, dep=dep, derived_capex=derived)


# ------------------------------------------------------------------ analyses
def _not_ready(an, msg):
    an.add_note(msg)
    an.add_summary_only()
    return an


def build_ratios(company_folder):
    an = aa.Analysis("RATIO", "TOI ratio analysis")
    t, path, err = load(Path(company_folder))
    if t is None:
        return _not_ready(an, err)
    y = year_of(path)
    an.period = f"{y} vs {y - 1}" if y else "N vs N-1"
    rows = ratio_rows(t)

    def show(v, kind):
        if v is None:
            return None
        return round(v * 100, 1) if kind == "pct" else round(v, 1) if kind == "days" else round(v, 2)
    unit = {"pct": "%", "x": "times", "days": "days"}
    out, tags = [], []
    for g, name, formula, cur, pri, chg, kind, flag, comment in rows:
        out.append([g, name, show(cur, kind), show(pri, kind),
                    None if chg is None else show(chg, kind), unit[kind], flag or ("OK" if cur is not None else "n/a"), formula, comment])
        tags.append("error" if flag else "even" if cur is None else "total")
    checks = key_checks(t)
    flagged = [c for c in checks if c[3] == "Check"]
    S = [["TOI file", path.name, None, None, "", "section"],
         ["Ratios flagged", f"{sum(1 for r in rows if r[7]):,} of {len(rows)} ratios outside the normal range", sum(1 for r in rows if r[7]), None,
          "See 'Ratios'.", "error" if any(r[7] for r in rows) else "total"],
         ["Key risk checks", f"{len(flagged):,} of {len(checks)} checks need attention", len(flagged), None, "See 'Key risk checks'.",
          "error" if flagged else "total"]]
    for c in flagged[:12]:
        S.append(["   " + c[0], c[4], None, c[1], c[5], "warning"])
    S.append(["Revenue (B0)", "Current year", None, t.c("B0"), f"Prior year {t.p('B0'):,.0f}", "even"])
    S.append(["Profit after tax (B48)", "Current year", None, t.c("B48"), f"Prior year {t.p('B48'):,.0f}", "even"])
    S.append(["Total assets (A0) / liabilities", "Current year", None, t.c("A0"), f"Liabilities {liabilities(t, t.c):,.0f}", "even"])
    an.view("Summary", ["Area", "Result", "Count", "Amount (KHR)", "Note / action"], [r[:5] for r in S], [r[5] for r in S],
            money=[3], total=False, desc="Main ratios and risk checks from the TOI (current year N vs prior year N-1).")
    an.view("Ratios", ["Group", "Ratio", "N", "N-1", "Change", "Unit", "Flag", "Formula (TOI codes)", "Meaning"], out, tags,
            money=[], total=False, desc="Percent ratios are shown as %, others as times or days.")
    an.view("Key risk checks", ["Check", "Current / value", "Compared with", "Flag", "Finding", "Suggested procedure"], checks,
            ["error" if c[3] == "Check" else "total" for c in checks], money=[1, 2], total=False,
            desc="Cross-checks between TOI lines (same checks as the Audit Dashboard key-risk section).")
    an.view("TOI items", ["Code", "Description", "N", "N-1", "Change", "Change %"],
            [[k, v["desc"], v["cur"], v["pri"], v["cur"] - v["pri"], round((v["cur"] - v["pri"]) / v["pri"] * 100, 1) if v["pri"] else None]
             for k, v in sorted(t.items.items(), key=lambda kv: (kv[0][0], int(kv[0][1:])))],
            None, money=[2, 3, 4], total=False, desc="Every coded TOI line read from the file.")
    return an


def build_cashflow(company_folder):
    an = aa.Analysis("CASHFLOW", "TOI cash flow statement")
    root = Path(company_folder)
    t, path, err = load(root)
    if t is None:
        return _not_ready(an, err)
    y = year_of(path)
    an.period = str(y) if y else "Year N"
    try:
        step10 = read_step10(path)
    except Exception as exc:
        step10 = None
        an.add_note(f"Step 10: {str(exc).splitlines()[0]}")
    lines, tot = cashflow_lines(t, step10)
    rows, tags = [], []
    for section, codes, label, amount, note, level in lines:
        rows.append([section, codes, ("    " if level else "") + label, amount, note])
        tags.append("section" if level == 0 and not label.startswith("Difference") else
                    ("error" if label.startswith("Difference") and abs(amount) > TOL else "total" if label.startswith("Difference") else "even"))
    an.view("Cash flow", ["Section", "TOI codes", "Line", "Amount (KHR)", "Note"], rows, tags, money=[3], total=False,
            desc="Indirect method from the two balance sheets in the TOI and the income statement of year N.")
    ins = []
    def add(text, hit, action):
        if hit:
            ins.append([text, action])
    add(f"Profit {t.c('B48'):,.0f} but cash decreased by {-tot['net']:,.0f}", t.c("B48") > 0 and tot["net"] < -TOL,
        "See where the cash went in the lines above (stock, A/R, assets, repayments, dividends) and vouch the large ones.")
    add(f"Profit {t.c('B48'):,.0f} but operating cash flow {tot['op']:,.0f}", t.c("B48") > 0 and tot["op"] < 0,
        "Profit tied up in receivables / stock, or not real: check A/R, stock and revenue recognition.")
    add(f"Operating cash flow negative ({tot['op']:,.0f}) financed by {tot['fin']:,.0f} of financing", tot["op"] < 0 and tot["fin"] > 0,
        "Check the loans / capital received (agreements, related parties, interest WHT).")
    add(f"Large asset purchases {-tot['inv']:,.0f}", tot["inv"] < -TOL, "Vouch step 10 additions, VAT input and WHT on construction / installation.")
    add("Dividends / other equity outflow derived", any(l[1] == "A29" and l[3] < -TOL for l in lines),
        "Check dividend declarations and WHT on dividends (14 %), or prior-year adjustments.")
    add("Cash flow does not reconcile", abs(tot["diff"]) > TOL, "Balance sheet totals are inconsistent: check A0 = A28 in N and N-1.")
    if not ins:
        ins.append(["No unusual cash-flow pattern found", ""])
    an.view("Findings", ["Finding", "Suggested procedure"], ins, ["warning" if a else "total" for _, a in ins], total=False,
            desc="What the cash flow shows.")
    if step10 and step10.get("found"):
        an.view("Step 10 assets", ["Asset / line", "Additions (bought)", "Disposals"],
                [list(r) for r in step10["rows"]] + ([["TOTAL row: " + step10["total_row"][0], step10["total_row"][1], step10["total_row"][2]]]
                                                     if step10.get("total_row") else []),
                None, money=[1, 2], total=False,
                desc=f"Read from step 10. Additions column {step10['add_col']}; disposals column {step10['disp_col'] or 'not found'}.")
    else:
        an.add_note("Step 10: purchase column not recognised (looked for ទិញ / បន្ថែម / កើន / addition / acquisition); "
                    "purchases were derived from the change in long-term assets plus depreciation.")
    return an
