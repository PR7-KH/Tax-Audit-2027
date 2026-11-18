"""Audit Dashboard: one risk summary that collects every flagged item from every test.

Sources (whatever exists for the company):
  * ADT rules on each TOI line (rows whose Flag is "Check"), key risk checks and ratio flags;
  * cash-flow findings; audit report vs TOI; revenue, expense, TOS, WHT, tax payments, A/R & A/P;
  * the sales tie-out (tax risk finding).
Each finding: area, item, TOI code / source, finding, amount, severity, suggested action, where to look.

build(folder, analyses, risk_report=None, workbook=None, calc=None) -> audit_analysis.Analysis with
two views: "Tax Risk Overview" (flagged items only) and "ADT (all TOI items)" (every line of the template).
Nothing here touches the GUI.
"""
import re

import audit_analysis as aa

AREAS = [  # (analysis key, area title, menu label used in "Where to look")
    ("KEY", "1. TOI key risk checks", "Ratio analysis (TOI) → Key risk checks"),
    ("ADT", "2. TOI line checks (ADT rules)", "Audit Dashboard → ADT (all TOI items)"),
    ("an_RATIO", "3. TOI ratios", "Ratio analysis (TOI) → Ratios"),
    ("an_CASHFLOW", "4. Cash flow", "Cash flow (TOI) → Findings"),
    ("an_AUDITREP", "5. Audit report vs TOI", "Audit report vs TOI"),
    ("tax_risk", "6. Sales tie-out (tax risk finding)", "Tax risk finding"),
    ("an_REVENUE", "7. Revenue", "Revenue"),
    ("an_EXPENSE", "8. Expense", "Expense"),
    ("an_TOS", "9. Salary tax (TOS)", "TOS (salary tax)"),
    ("an_WHT", "10. Withholding tax (WHT)", "WHT (withholding tax)"),
    ("an_TAXPAID", "11. Tax payments", "Tax Payment History"),
    ("an_ARAP", "12. A/R and A/P", "A/R & A/P"),
]
SEV = {"error": "High", "warning": "Medium"}
COLUMNS = ["Area", "Item", "TOI code / source", "Finding", "Amount (KHR)", "Severity", "Suggested action", "Where to look"]


def _first_number(row, money):
    for i in money:
        if i < len(row) and isinstance(row[i], (int, float)) and not isinstance(row[i], bool) and row[i]:
            return row[i]
    return None


def _clean(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def from_summary(an, where):
    """Generic: the Summary rows an analysis marked red (High) or amber (Medium)."""
    t = an.views.get("Summary")
    if t is None or not t.row_tags:
        return []
    out = []
    for row, tag in zip(t.rows, t.row_tags):
        if tag not in SEV:
            continue
        texts = [_clean(v) for v in row if isinstance(v, str) and _clean(v)]
        item = texts[0] if texts else ""
        finding = texts[1] if len(texts) > 1 else ""
        action = texts[-1] if len(texts) > 2 else ""
        if item.lower().startswith("not ready"):
            continue
        out.append([item, "", finding, _first_number(row, t.money), SEV[tag], action, where + " → Summary"])
    return out


def adt_rows(workbook, calc):
    """[(code, description, N, N-1, change, change %, flag, issue KH, issue EN)] from the ADT sheet."""
    if workbook is None or calc is None or "ADT" not in workbook.sheetnames:
        return []
    ws = workbook["ADT"]
    out = []
    for r in range(1, ws.max_row + 1):
        code = ws.cell(r, 4).value
        if not (isinstance(code, str) and re.fullmatch(r"([ABCD]\d{1,2}|KEY\d+)", code.strip())):
            continue
        v = lambda c: calc.value("ADT", r, c)
        def num(x):
            return x if isinstance(x, (int, float)) and not isinstance(x, bool) else None
        out.append([code.strip(), _clean(ws.cell(r, 3).value), num(v(5)), num(v(6)), num(v(7)), num(v(8)),
                    _clean(v(10)), _clean(v(11)), _clean(v(12))])
    return out


def build(company_folder, analyses, risk_report=None, workbook=None, calc=None):
    an = aa.Analysis("DASHBOARD", "Audit Dashboard")
    rows, tags = [], []
    counts = {}
    adt = adt_rows(workbook, calc)
    for key, area, where in AREAS:
        items = []
        if key == "KEY":
            ratio = analyses.get("an_RATIO")
            t = ratio.views.get("Key risk checks") if ratio else None
            for r in (t.rows if t else []):
                if r[4] == "Check":
                    items.append([r[0], r[1], r[5], r[2], "High", r[6], where])
        elif key == "ADT":
            for code, desc, cur, pri, chg, pct, flag, kh, en in adt:
                if code.startswith("KEY") or flag != "Check":
                    continue
                items.append([desc or code, code, en or kh, cur, "Medium", kh if en else "", where])
        elif key == "an_RATIO":
            ratio = analyses.get("an_RATIO")
            t = ratio.views.get("Ratios") if ratio else None
            for r in (t.rows if t else []):
                if r[6] not in ("OK", "n/a", "", None):
                    items.append([r[1], r[7], f"{r[6]}: N {r[2]} {r[5]} (N-1 {r[3] if r[3] is not None else '–'})", None, "Medium", r[8], where])
        elif key == "an_CASHFLOW":
            cf = analyses.get("an_CASHFLOW")
            t = cf.views.get("Findings") if cf else None
            for r in (t.rows if t else []):
                if r[1]:
                    items.append([r[0], "TOI", r[0], None, "Medium", r[1], where])
        elif key == "tax_risk":
            for f in (risk_report or {}).get("findings", []):
                if f["sev"] in ("High", "Medium"):
                    items.append([f"{f['area']}: {f['title']}", "Sales / GL / bank", _clean(f.get("evidence")), f.get("amount"), f["sev"],
                                  _clean(f.get("action")), where])
        elif key == "an_TAXPAID":
            a = analyses.get(key)
            for label, status, diff in getattr(a, "findings", []) if a else []:
                items.append([label, "Payment History vs returns", status, diff, "High" if status == "Not paid" else "Medium",
                              "Compare the payment receipts with the return; late or missing payment means penalties.",
                              where + " → Paid vs declared"])
        else:
            a = analyses.get(key)
            if a is not None:
                items = from_summary(a, where)
        counts[area] = (sum(1 for i in items if i[4] == "High"), sum(1 for i in items if i[4] == "Medium"), key in analyses or
                        (key in ("KEY", "an_RATIO") and "an_RATIO" in analyses) or (key == "ADT" and bool(adt)) or
                        (key == "tax_risk" and risk_report is not None))
        if not items:
            continue
        rows.append([area, "", "", f"{counts[area][0]} High, {counts[area][1]} Medium", None, "", "", where.split(" → ")[0]])
        tags.append("headline")
        items.sort(key=lambda i: (i[4] != "High",))
        for i in items:
            rows.append([area] + i)
            tags.append("error" if i[4] == "High" else "warning")
    high = sum(c[0] for c in counts.values())
    med = sum(c[1] for c in counts.values())
    overview, otags = [], []
    overview.append(["Overview", "All areas", "", f"{high} High, {med} Medium items to check", None, "",
                     "Start with the High items. Double-click a row to open its menu.", ""])
    otags.append("headline")
    for key, area, where in AREAS:
        h, m, ran = counts.get(area, (0, 0, False))
        overview.append(["Overview", area, "", ("Not run: source files missing" if not ran else
                                                "No findings" if not (h or m) else f"{h} High, {m} Medium"), None,
                         "", "", where])
        otags.append("even" if not ran else "total" if not (h or m) else "error" if h else "warning")
    an.view("Tax Risk Overview", COLUMNS, overview + rows, otags + tags, money=[4], total=False, wide=True,
            desc="Every flagged item from every test, in one place. Red = High, amber = Medium. Click a heading to filter.")
    an.view("ADT (all TOI items)", ["Code", "Description", "N", "N-1", "Change", "Change %", "Flag", "Tax/Audit issue (KH)", "Tax/Audit issue (EN)"],
            [r[:5] + [round(r[5] * 100, 1) if isinstance(r[5], (int, float)) else None] + r[6:] for r in adt],
            ["warning" if r[6] == "Check" else "even" for r in adt], money=[2, 3, 4], total=False, wide=True,
            desc="Every line of the ADT template (rules sheet) with its result. Edit the rules in rules/ADT_rules_template.xlsx.")
    if not adt:
        an.add_note("ADT rules not evaluated: the TOI and Info (company information) files are needed to build the ADT sheet.")
    return an
