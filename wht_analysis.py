"""WHT (withholding tax) risk: purchases in the monthly PR that should carry WHT, compared with
what was declared in the monthly WHT returns.

WHT concept used here (Cambodia): when the company pays for a SERVICE, RENTAL, FEE, interest or
royalty to a supplier that is a physical person / not VAT registered / gives no tax invoice, it must
withhold tax (rental 10 %, services and other fees 15 %; 14 % for non-residents). Purchases of
GOODS, raw materials, imports and fixed assets are not subject to WHT, and a VAT-registered supplier
with a tax invoice (VAT shown on the PR line) is not either.

For each PR month the purchases that look subject to WHT are grouped by supplier and matched with
the WHT return of the same month (or the next month, when the payment was later) by TIN, else by
name. Declared items are kept in their own view as "OK"; only what is not (or only partly) declared,
or cannot be classified, is raised as a risk.

build(company_folder) -> audit_analysis.Analysis. Nothing here touches the GUI.
"""
import re
from collections import defaultdict
from pathlib import Path

import audit_analysis as aa

TOL = 0.02   # 2 % of the PR amount is accepted as rounding / gross-up difference

RATES = {"Rental": 0.10, "Service / fee": 0.15, "Interest": 0.15, "Royalty": 0.15, "Unclear": 0.15}
NATURE = [  # (nature, regex on description + supplier name); first match wins
    ("Excluded: fixed asset", r"\basset|equipment|machine|vehicle|\bcar\b|computer|laptop|furniture|ទ្រព្យ|ម៉ាស៊ីន|ឧបករណ៍|រថយន្ត"),
    ("Excluded: goods / material", r"goods|material|raw\b|fabric|stock|inventory|merchandise|spare part|stationery|fuel|"
                                   r"petrol|diesel|ទំនិញ|វត្ថុធាតុ|សម្ភារ|ប្រេង"),
    ("Excluded: utility", r"\bedc\b|electric|electricity|water supply|ppwsa|អគ្គិសនី|ទឹកស្អាត"),
    ("Rental", r"rent|rental|lease|leasing|ជួល"),
    ("Interest", r"interest|ការប្រាក់"),
    ("Royalty", r"royalt|license fee|franchise|សួយសារ"),
    ("Service / fee", r"service|fee|consult|repair|maintenance|transport|delivery|clean|security|guard|legal|"
                      r"audit|accounting|advertis|marketing|commission|training|design|install|labou?r|freight|"
                      r"logistic|courier|agent|broker|management|សេវា|ថ្លៃ|ជួសជុល|ដឹកជញ្ជូន|ប្រឹក្សា"),
]
NON_TAXABLE = re.compile(r"បុគ្គល\s*មិន\s*ជាប់\s*អាករ|non[- ]?taxable\s+person|non[- ]?vat\s+person", re.I)
TIN_FORMAT = re.compile(r"[A-Z]\d{3}\d{6,}")      # L001-100163327, K001-..., after removing - and spaces
INDIVIDUAL = re.compile(r"individual|physical|person|natural|non[- ]?vat|non[- ]?registered|unregistered|រូបវន្ត|មិនមែន|មិនបាន", re.I)


def _norm_tin(v):
    t = re.sub(r"[^0-9A-Z]", "", str(v or "").upper())
    return "" if t in ("", "0", "NA", "NONE") or set(t) == {"0"} else t


def _tokens(name):
    stop = {"CO", "LTD", "LIMITED", "INC", "THE", "AND", "COMPANY", "MR", "MRS", "MS", "MISS", "LOK", "LOKSREY", "PLC", "ENTERPRISE"}
    return {t for t in re.findall(r"[A-Z0-9ក-៿]+", str(name or "").upper()) if t not in stop and len(t) > 1}


def norm_inv(v):
    return re.sub(r"[^0-9A-Z]", "", str(v or "").upper())


def same_invoice(a, b):
    """Invoice numbers match when equal, or when one ends with the other (INV-001 vs 001), at least 3 characters."""
    a, b = norm_inv(a), norm_inv(b)
    if not a or not b:
        return False
    if a == b:
        return True
    short, long_ = sorted((a, b), key=len)
    return len(short) >= 3 and long_.endswith(short) and not long_[-len(short) - 1:-len(short)].isdigit()


def same_party(tin_a, name_a, tin_b, name_b):
    if tin_a and tin_b:
        return tin_a == tin_b
    a, b = _tokens(name_a), _tokens(name_b)
    if not a or not b:
        return False
    return len(a & b) / min(len(a), len(b)) >= 0.6


def _pick(low, tests, avoid=(), limit=None, after=0):
    for i, l in enumerate(low[:limit] if limit else low):
        if i < after or not l:
            continue
        if any(re.search(t, l) for t in tests) and not any(re.search(a, l) for a in avoid):
            return i
    return None


def map_pr(labels, first_col, last_col):
    low = [l.casefold() for l in labels]
    text_end = first_col - 1
    cols = {
        "date": _pick(low, (r"កាលបរិច្ឆេទ", r"\bdate\b"), limit=text_end),
        "invoice": _pick(low, (r"វិក្កយបត្រ", r"invoice"), limit=text_end),
        "type": _pick(low, (r"ប្រភេទ", r"\btype\b", r"status"), limit=text_end),
        "tin": _pick(low, (r"អត្តសញ្ញាណកម្ម", r"\btin\b", r"vat ?no", r"vattin"), limit=text_end),
        "name": _pick(low, (r"ឈ្មោះ", r"name", r"supplier"), limit=text_end),
        "desc": _pick(low, (r"បរិយាយ", r"ពិពណ៌នា", r"description", r"particular", r"detail", r"memo", r"item"), avoid=(r"ឈ្មោះ", r"name")),
        "total": first_col - 1,
    }
    if cols["date"] is None:
        cols["date"] = 1
    if cols["invoice"] is None and text_end > 2:
        cols["invoice"] = 2
    if cols["name"] is None and text_end > 5:
        cols["name"] = 5
    amount = range(first_col - 1, min(last_col, len(low)))
    cols["import"] = [i for i in amount if re.search(r"នាំចូល|import", low[i])]
    cols["vat"] = [i for i in amount if re.search(r"អាករ|\bvat\b", low[i]) and i not in cols["import"]
                   and not re.search(r"មិន|ជាប់|non|exempt|taxable|លើកលែង|តម្លៃ|value|សរុប|total", low[i])]
    cols["nonvat"] = [i for i in amount if re.search(r"មិនជាប់|non[- ]?vat|not subject|exempt|លើកលែង", low[i])]
    return cols


def map_wht(labels, first_col):
    low = [l.casefold() for l in labels]
    end = first_col - 1
    cols = {
        "date": _pick(low, (r"កាលបរិច្ឆេទ", r"\bdate\b"), limit=end),
        "tin": _pick(low, (r"អត្តសញ្ញាណកម្ម", r"\btin\b", r"vat ?no"), limit=end),
        "name": _pick(low, (r"ឈ្មោះ", r"name", r"payee", r"recipient"), limit=end),
        "nature": _pick(low, (r"ប្រភេទ", r"បរិយាយ", r"nature", r"type", r"description", r"income"), avoid=(r"អត្រា", r"rate"), limit=end),
        "rate": _pick(low, (r"អត្រា", r"rate", r"%"), limit=end),
        "invoice": _pick(low, (r"វិក្កយបត្រ", r"invoice", r"\binv\b", r"លេខយោង", r"reference", r"\bref\b"),
                         avoid=(r"កាលបរិច្ឆេទ", r"date"), limit=end),
        "resident": _pick(low, (r"និវាសនជន", r"resident"), limit=end),
        "base": first_col - 1,
        "tax": first_col,
    }
    return cols


def has_tin(tin):
    return bool(TIN_FORMAT.fullmatch(_norm_tin(tin)))


def classify(desc, name, supplier_type, vat, imported, tin="", type_column=True):
    """(nature, reason) for one PR line. Natures starting with 'Excluded' are not WHT.
    A WHT candidate must meet BOTH criteria: supplier type បុគ្គលមិនជាប់អាករ (non-taxable person) and no
    tax registration number (TIN such as L001-XXXXXXXXX / K00X-XXXXXXXXX). Goods, assets and imports are not WHT."""
    if imported:
        return "Excluded: import", "Import columns used"
    if has_tin(tin):
        return "Excluded: registered taxpayer (has TIN)", f"TIN {tin}"
    if type_column and not NON_TAXABLE.search(supplier_type or ""):
        return "Excluded: supplier type", f"Supplier type '{supplier_type or 'blank'}' is not បុគ្គលមិនជាប់អាករ"
    if not type_column and vat > 0:
        return "Excluded: VAT invoice", "VAT shown on the PR line (no supplier-type column in this PR file)"
    text = f"{desc} {name}".casefold()
    nature = next((n for n, rx in NATURE if re.search(rx, text, re.I)), None)
    if nature and nature.startswith("Excluded"):
        return nature, "Keyword in description / supplier"
    why = ("បុគ្គលមិនជាប់អាករ, no TIN" if type_column else "No TIN, no VAT (no supplier-type column: check the type)")
    if nature:
        return nature, why
    return "Unclear", why + "; nature not shown: check the invoice (goods are not WHT)"


def read_files(folder, kind):
    import consolidate as cons
    layout = cons.LAYOUTS[kind]
    out, notes = [], []
    if folder is None:
        return out, notes
    for year, path in cons.files_in(Path(folder)):
        if not str(year).isdigit():
            continue
        month = aa.file_month(path.name, year)
        if month is None:
            notes.append(f"{kind} {path.name}: month not recognised in the file name (use e.g. {kind}_03_{year}.xlsx)")
            continue
        try:
            src = cons.read_source(path, layout)
        except Exception as exc:
            notes.append(f"{kind} {path.name}: {str(exc).splitlines()[0]}")
            continue
        labels = aa.two_row_labels(src.headers, src.merges)
        out.append(((int(year), month), path, labels, [list(r) for r in src.rows]))
    return out, notes


def _cell(r, i):
    return r[i] if i is not None and i < len(r) else None


def _text(v):
    return re.sub(r"\s+", " ", str(v if v is not None else "")).strip()


TOTAL_ROW = re.compile(r"^\s*(grand\s+)?(sub\s*)?total\b|^\s*សរុប", re.I)


def build(company_folder, tol=TOL):
    import consolidate as cons
    root = Path(company_folder)
    an = aa.Analysis("WHT", "WHT withholding tax risk")
    pr_dir, wht_dir = aa.subfolder(root, {"PR"}), aa.subfolder(root, {"WHT"})
    if pr_dir is None:
        an.add_note("No PR folder: WHT is tested from the monthly purchase records (PR/<year>/*.xlsx).")
        an.add_summary_only()
        return an
    pr_files, n1 = read_files(pr_dir, "PR")
    wht_files, n2 = read_files(wht_dir, "WHT")
    for n in n1 + n2:
        an.add_note(n)
    if wht_dir is None:
        an.add_note("No WHT folder: every WHT-type purchase is shown as not declared.")
    if not pr_files:
        an.add_note("No readable monthly PR file.")
        an.add_summary_only()
        return an
    prl, whl = cons.LAYOUTS["PR"], cons.LAYOUTS["WHT"]
    M = aa.month_label

    # ---- PR lines
    lines, mapping, mtags = [], [], []
    for ym, path, labels, rows in pr_files:
        c = map_pr(labels, prl.first_col, prl.last_col)
        lab = lambda i: f"{aa.col_letter(i)}: {labels[i]}" if i is not None and i < len(labels) else "not found"
        mapping.append(["PR", M(ym), path.name, lab(c["tin"]), lab(c["name"]), lab(c["type"]), lab(c["desc"]),
                        lab(c["total"]), ", ".join(lab(i) for i in c["vat"]) or "not found",
                        ", ".join(lab(i) for i in c["import"]) or "—"])
        mtags.append("warning" if c["desc"] is None or not c["vat"] else "even")
        for r in rows:
            amount = aa.num(_cell(r, c["total"]))
            name = _text(_cell(r, c["name"]))
            first = " ".join(_text(x) for x in r[:prl.first_col - 1] if isinstance(x, str))
            if not amount or TOTAL_ROW.search(first):
                continue
            vat = sum(aa.num(_cell(r, i)) for i in c["vat"])
            imported = any(aa.num(_cell(r, i)) for i in c["import"])
            stype, desc = _text(_cell(r, c["type"])), _text(_cell(r, c["desc"]))
            nature, why = classify(desc, name, stype, vat, imported, _text(_cell(r, c["tin"])), c["type"] is not None)
            d = aa.to_date(_cell(r, c["date"]))
            lines.append(dict(ym=ym, date=str(d) if d else _text(_cell(r, c["date"])), invoice=_text(_cell(r, c["invoice"])),
                              tin=_norm_tin(_cell(r, c["tin"])), tin_text=_text(_cell(r, c["tin"])), name=name, type=stype,
                              desc=desc, amount=amount, vat=vat, nature=nature, why=why, file=path.name))

    # ---- WHT lines
    wht = []
    for ym, path, labels, rows in wht_files:
        c = map_wht(labels, whl.first_col)
        lab = lambda i: f"{aa.col_letter(i)}: {labels[i]}" if i is not None and i < len(labels) else "not found"
        mapping.append(["WHT", M(ym), path.name, lab(c["tin"]), lab(c["name"]), lab(c["nature"]),
                        lab(c["rate"]) + " · invoice " + lab(c["invoice"]),
                        lab(c["base"]), lab(c["tax"]), lab(c["resident"])])
        mtags.append("warning" if c["name"] is None else "even")
        for r in rows:
            base, tax = aa.num(_cell(r, c["base"])), aa.num(_cell(r, c["tax"]))
            first = " ".join(_text(x) for x in r[:whl.first_col - 1] if isinstance(x, str))
            if not (base or tax) or TOTAL_ROW.search(first):
                continue
            wht.append(dict(ym=ym, tin=_norm_tin(_cell(r, c["tin"])), tin_text=_text(_cell(r, c["tin"])),
                            name=_text(_cell(r, c["name"])), nature=_text(_cell(r, c["nature"])),
                            rate=_text(_cell(r, c["rate"])), invoice=_text(_cell(r, c["invoice"])),
                            base=base, tax=tax, file=path.name, used=0.0, left=base))

    # ---- group WHT-type purchases by supplier and month, then match
    groups = defaultdict(list)
    for ln in lines:
        if not ln["nature"].startswith("Excluded"):
            groups[(ln["ym"], ln["tin"] or "NAME:" + ln["name"].upper())].append(ln)
    # Each WHT line's base is used once: first by purchases of the same month, then (if the payment
    # was later) by the previous month's purchases that are still not covered.
    ordered = sorted(groups.items(), key=lambda x: (x[0][0], x[1][0]["name"].upper()))
    alloc = {k: {"base": 0.0, "tax": 0.0, "when": set(), "by": set()} for k, _ in ordered}

    def take(key, items, targets, label, by):
        first, need = items[0], sum(i["amount"] for i in items) - alloc[key]["base"]
        invoices = [i["invoice"] for i in items if i["invoice"]]
        for w in wht:
            if need <= 0:
                break
            if w["ym"] not in targets or w["left"] <= 0:
                continue
            if by == "invoice":
                if not (w["invoice"] and any(same_invoice(inv, w["invoice"]) for inv in invoices)):
                    continue
            elif not same_party(first["tin"], first["name"], w["tin"], w["name"]):
                continue
            part = min(w["left"], need) if (label or by == "invoice") else w["left"]   # same month by name: whole line
            share = part / w["base"] if w["base"] else 1.0
            w["left"] -= part
            w["used"] += 1
            need -= part
            alloc[key]["base"] += part
            alloc[key]["tax"] += w["tax"] * share
            alloc[key]["when"].add(label)
            alloc[key]["by"].add(by)
    later = lambda ym, n: [ym] if n == 0 else later(aa.next_month(ym), n - 1) + [ym]
    for key, items in ordered:          # 1. invoice number: one month before to three months after
        ym = key[0]
        take(key, items, {aa.prev_month(ym), *later(ym, 3)}, "", "invoice")
    for key, items in ordered:          # 2. TIN / name, same month
        take(key, items, {key[0]}, "", "name / TIN")
    for key, items in ordered:          # 3. TIN / name, paid later (next one or two months)
        n1 = aa.next_month(key[0])
        take(key, items, {n1, aa.next_month(n1)}, "later month", "name / TIN")

    results = []
    for key, items in ordered:
        ym, first = key[0], items[0]
        amount = sum(i["amount"] for i in items)
        natures = sorted({i["nature"] for i in items})
        nature = "Rental" if "Rental" in natures else natures[0] if len(natures) == 1 else " / ".join(natures)
        rate = max(RATES.get(n, 0.15) for n in natures)
        base, tax = alloc[key]["base"], alloc[key]["tax"]
        found = base > 0
        when = ", ".join(sorted(alloc[key]["by"])) + (" · later month" if alloc[key]["when"] == {"later month"} else "")
        gap = max(0.0, amount - base)
        if not found:
            result = "Not declared"
        elif gap <= tol * amount:
            result = "Declared" + (f" ({when})" if when else "")
        else:
            result = "Partly declared" + (f" ({when})" if when else "")
        unclear = all(n == "Unclear" for n in natures)
        risk = 0.0 if result.startswith("Declared") else gap * rate
        tag = ("total" if result.startswith("Declared") else "warning" if unclear or result.startswith("Partly") else "error")
        action = ("OK: declared in the WHT return" if result.startswith("Declared") else
                  "Check the invoice: if it is a service / rental from a non-VAT supplier, WHT is due" if unclear else
                  f"Ask why WHT {rate:.0%} was not declared; request the contract / invoice and payment proof")
        results.append(dict(ym=ym, name=first["name"], tin=first["tin_text"], type=first["type"], nature=nature,
                            n=len(items), invoices=", ".join(i["invoice"] for i in items if i["invoice"])[:120],
                            desc="; ".join(sorted({i["desc"] for i in items if i["desc"]}))[:160],
                            amount=amount, base=base, tax=tax, gap=gap, rate=rate, risk=risk,
                            result=result, tag=tag, action=action, unclear=unclear))

    def res_row(x):
        return [aa.month_key(x["ym"]), x["name"], x["tin"], x["type"], x["nature"], x["desc"], x["n"], x["invoices"],
                x["amount"], x["base"], x["tax"], x["gap"], f"{x['rate']:.0%}", x["risk"], x["result"], x["action"]]
    RES_COLS = ["Month", "Supplier", "TIN", "Supplier type", "Nature", "Description", "PR lines", "Invoices",
                "PR amount", "WHT base declared", "WHT declared", "Not declared", "Rate", "Estimated WHT at risk", "Result", "Suggested action"]
    risky = sorted([x for x in results if not x["result"].startswith("Declared")], key=lambda x: (x["unclear"], -x["risk"]))
    ok = [x for x in results if x["result"].startswith("Declared")]

    # recurring suppliers not declared
    rec = defaultdict(list)
    for x in risky:
        rec[(_norm_tin(x["tin"]) or x["name"].upper())].append(x)
    rec_rows = []
    for _k, xs in rec.items():
        rec_rows.append([xs[0]["name"], xs[0]["tin"], xs[0]["nature"], len(xs), aa.month_list([x["ym"] for x in xs]),
                         sum(x["amount"] for x in xs), sum(x["gap"] for x in xs), sum(x["risk"] for x in xs),
                         "Recurring: likely a contract (rent / service); request it" if len(xs) >= 3 else ""])
    rec_rows.sort(key=lambda r: -r[7])

    # WHT declared with no PR purchase
    orphan = [[aa.month_key(w["ym"]), w["name"], w["tin_text"], w["nature"], w["rate"], w["base"], w["tax"], w["file"],
               "No matching WHT-type purchase in PR this or last month: check the expense is recorded and the PR is complete"]
              for w in wht if not w["used"]]

    # monthly summary
    months = sorted({ln["ym"] for ln in lines} | {w["ym"] for w in wht})
    mon_rows, mon_tags = [], []
    for ym in months:
        L = [l for l in lines if l["ym"] == ym]
        cand = [l for l in L if not l["nature"].startswith("Excluded")]
        R = [x for x in results if x["ym"] == ym]
        W = [w for w in wht if w["ym"] == ym]
        nd = sum(x["gap"] for x in R if not x["result"].startswith("Declared"))
        risk = sum(x["risk"] for x in R)
        st = "No PR file" if not L else "OK" if not nd else "Risk"
        mon_rows.append([aa.month_key(ym), len(L), sum(l["amount"] for l in L), len(cand), sum(l["amount"] for l in cand),
                         len(W), sum(w["base"] for w in W), sum(w["tax"] for w in W), nd, risk, st])
        mon_tags.append("error" if st == "Risk" and risk else "warning" if st != "OK" else "total")

    # ---- summary
    period = f"{M(months[0])} to {M(months[-1])}" if months else ""
    an.period = period
    total_risk = sum(x["risk"] for x in risky)
    S = []
    def s(area, result, count, amount, action, tag):
        S.append(([area, result, count, amount, action], tag))
    s("Period", f"{period}: {len(pr_files)} PR file(s), {len(wht_files)} WHT file(s), {len(lines):,} purchase lines", len(lines), None, "", "section")
    cand = [l for l in lines if not l["nature"].startswith("Excluded")]
    s("WHT-type purchases in PR", "Supplier type បុគ្គលមិនជាប់អាករ AND no TIN: services, rental, management fees, interest, royalty, unclear",
      len(cand), sum(l["amount"] for l in cand), "See 'PR lines classified' (filter the Nature column).", "even")
    s("Declared in WHT (OK)", f"{len(ok):,} supplier-month(s) found in the WHT return", len(ok), sum(x["amount"] for x in ok),
      "See 'Declared OK'. No action.", "total")
    nd = [x for x in risky if not x["unclear"]]
    s("RISK: service / rental not declared", f"{len(nd):,} supplier-month(s)", len(nd), sum(x["risk"] for x in nd),
      "Amount = estimated WHT at risk. See 'WHT risk'.", "error" if nd else "total")
    un = [x for x in risky if x["unclear"]]
    s("CHECK: nature unclear, not declared", f"{len(un):,} supplier-month(s) from non-VAT suppliers", len(un), sum(x["risk"] for x in un),
      "Look at the invoices: goods are not WHT; services and rental are.", "warning" if un else "total")
    s("Recurring suppliers not declared", f"{sum(1 for r in rec_rows if r[3] >= 3):,} supplier(s) in 3 or more months", len(rec_rows),
      sum(r[7] for r in rec_rows), "See 'Recurring not declared'. Request the rent / service contract.", "warning" if rec_rows else "total")
    s("WHT declared but not in PR", f"{len(orphan):,} WHT line(s)", len(orphan), sum(o[6] for o in orphan),
      "See 'WHT not in PR'. Check the purchase is recorded.", "warning" if orphan else "total")
    s("Total estimated WHT at risk", "Rental 10 %, services / fees / unclear 15 % on the amount not declared", None, total_risk,
      "Estimate only: confirm resident status (non-resident 14 %) and whether the PR amount is gross or net.", "error" if total_risk else "total")

    an.view("Summary", ["Area", "Result", "Count", "Amount (KHR)", "Suggested audit action"], [r for r, _ in S], [t for _, t in S],
            money=[3], total=False, desc=f"WHT risk from PR vs WHT returns, {period}. Rental 10 %, services / fees 15 %.")
    an.view("Monthly summary", ["Month", "PR lines", "PR amount", "WHT-type lines", "WHT-type amount", "WHT lines", "WHT base declared",
                                "WHT declared", "Not declared", "Estimated WHT at risk", "Status"],
            mon_rows, mon_tags, money=[2, 4, 6, 7, 8, 9], month_col=0, desc="Each month: WHT-type purchases in PR vs WHT declared.")
    an.view("WHT risk", RES_COLS, [res_row(x) for x in risky], [x["tag"] for x in risky], money=[8, 9, 10, 11, 13], month_col=0,
            wide=True, desc="WHT-type purchases not (or only partly) declared, by supplier and month. Red = service / rental; amber = check the invoice.")
    an.view("Recurring not declared", ["Supplier", "TIN", "Nature", "Months", "Months (list)", "PR amount", "Not declared",
                                       "Estimated WHT at risk", "Note"], rec_rows,
            ["error" if r[3] >= 3 else "warning" for r in rec_rows], money=[5, 6, 7], desc="Suppliers with undeclared WHT, all months together.")
    an.view("Declared OK", RES_COLS, [res_row(x) for x in ok], ["total"] * len(ok), money=[8, 9, 10, 11, 13], month_col=0,
            wide=True, desc="WHT-type purchases found in the WHT return of the same or next month.")
    an.view("WHT not in PR", ["Month", "Payee", "TIN", "Nature", "Rate", "WHT base", "WHT", "File", "Note"], orphan,
            ["warning"] * len(orphan), money=[5, 6], month_col=0, desc="WHT return lines with no WHT-type purchase in PR.")
    an.view("PR lines classified", ["Month", "Date", "Invoice", "Supplier", "TIN", "Supplier type", "Description", "Amount", "VAT",
                                    "Nature", "Reason", "File"],
            [[aa.month_key(l["ym"]), l["date"], l["invoice"], l["name"], l["tin_text"], l["type"], l["desc"], l["amount"], l["vat"],
              l["nature"], l["why"], l["file"]] for l in lines],
            ["even" if l["nature"].startswith("Excluded") else "warning" for l in lines], money=[7, 8], month_col=0,
            wide=True, desc="Every PR line and why it is (or is not) WHT-type. Filter Nature to review.")
    an.view("WHT declarations", ["Month", "Payee", "TIN", "Nature", "Rate", "WHT base", "WHT", "Matched to PR", "File"],
            [[aa.month_key(w["ym"]), w["name"], w["tin_text"], w["nature"], w["rate"], w["base"], w["tax"],
              "Yes" if w["used"] else "No", w["file"]] for w in wht], None, money=[5, 6], month_col=0,
            desc="Every line of the WHT returns.")
    an.view("Column mapping", ["Return", "Month", "File", "TIN", "Name", "Type / nature", "Description / rate", "Amount", "VAT / WHT",
                               "Import / resident"], mapping, mtags,
            desc="Which column of each PR / WHT file was used. Check this first if a result looks wrong.")
    return an
