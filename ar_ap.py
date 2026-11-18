"""A/R and A/P listings compared with the TOI, the GL and the audited financial statements, with
aging and per-customer / per-supplier detail.

AR folder: customer listing or aging (Excel / CSV).  AP folder: supplier listing or aging.
Columns are found by heading text: name (customer / supplier / vendor / party / name), balance
(balance / outstanding / amount due / amount / total), date (invoice date / date), and aging buckets
(headings with "days", "current", "over", "not due" or "0-30" ...). When there are no bucket columns
but a date column, the age is worked out to the year end (31 December of the TOI year).

build(company_folder) -> audit_analysis.Analysis.  Nothing here touches the GUI.
"""
import re
from collections import defaultdict
from datetime import date
from pathlib import Path

import audit_analysis as aa
import finance_view as fv

TOTAL = re.compile(r"^\s*(grand\s+)?(sub\s*)?total\b|^\s*សរុប", re.I)
BUCKET = re.compile(r"\bdays?\b|^current$|not due|over|older|above|\d+\s*[-–]\s*\d+|>\s*\d+|\+$|ថ្ងៃ", re.I)
AGE_BUCKETS = ((90, "0–90 days"), (180, "91–180 days"), (365, "181–365 days"), (None, "Over 365 days"))


def _pick(cols, words, avoid=()):
    low = [c.lower() for c in cols]
    for w in words:
        for i, c in enumerate(low):
            if w in c and not any(a in c for a in avoid):
                return i
    return None


def read_listing(folder, kind):
    """dict(file, rows=[{name, balance, date, buckets{}}], bucket_cols, columns used) or None."""
    if folder is None:
        return None
    files = [p for p in fv.table_files(folder)]
    if not files:
        return None
    cols, rows = fv.read_rows(files[0])
    party_words = ("customer", "client", "buyer", "debtor") if kind == "AR" else ("supplier", "vendor", "creditor", "payee")
    name = _pick(cols, party_words + ("party", "name", "counterparty", "ឈ្មោះ"), avoid=("account name", "file"))
    buckets = [i for i, c in enumerate(cols) if BUCKET.search(c.strip()) and i != name]
    bal = _pick(cols, ("balance", "outstanding", "amount due", "closing", "total", "amount", "សមតុល្យ"), avoid=("credit limit", "days"))
    if bal in buckets:
        buckets.remove(bal)
    dte = _pick(cols, ("invoice date", "due date", "date", "កាលបរិច្ឆេទ"))
    if name is None or (bal is None and not buckets):
        raise ValueError(f"{files[0].name}: need a {'customer' if kind == 'AR' else 'supplier'} / name column and a balance or aging columns")
    out = []
    for r in rows:
        nm = str(r[name] or "").strip()
        if not nm or TOTAL.search(nm):
            continue
        b = {cols[i]: aa.num(r[i]) for i in buckets}
        balance = aa.num(r[bal]) if bal is not None else sum(b.values())
        if not balance and not any(b.values()):
            continue
        out.append(dict(name=nm, balance=balance, date=aa.to_date(r[dte]) if dte is not None else None, buckets=b))
    return dict(file=files[0].name, rows=out, bucket_cols=[cols[i] for i in buckets],
                used=f"name: {cols[name]}; balance: {cols[bal] if bal is not None else 'sum of aging columns'}"
                     f"{'; date: ' + cols[dte] if dte is not None else ''}{'; aging: ' + ', '.join(cols[i] for i in buckets) if buckets else ''}")


def age_bucket(days):
    for limit, label in AGE_BUCKETS:
        if limit is None or days <= limit:
            return label
    return AGE_BUCKETS[-1][1]


def is_old(label):
    t = label.lower()
    return bool(re.search(r"over\s*365|>\s*365|365\s*\+|over\s*1\s*year|\b(1|one)\s*year|older|over\s*360|>\s*360", t))


def gl_balance(root, year, pattern):
    """Closing GL balance of the accounts whose name matches pattern (last running balance, else debit − credit)."""
    import sales_reconciliation as sr
    gl_dir = aa.subfolder(root, {"GL", "GENERAL LEDGER", "LEDGER"})
    files = [p for p in fv.table_files(gl_dir) if p.suffix.lower() == ".xlsx"] if gl_dir else []
    if not files or not year:
        return None
    gl = sr.read_gl(files[0], year)
    by_acct = defaultdict(list)
    for r in gl:
        if re.search(pattern, r["acct"], re.I) and not re.search(r"tax|nssf|salary|wage", r["acct"], re.I):
            by_acct[(r["code"], r["acct"])].append(r)
    if not by_acct:
        return None
    total = 0.0
    for lines in by_acct.values():
        last = lines[-1]
        total += last["bal"] if last["bal"] is not None else sum(l["debit"] - l["credit"] for l in lines)
    return abs(total)


def audit_figure(root, item):
    """Current-year KHR figure of one compared item from the Audit Report, or None."""
    try:
        import audit_report as ar
        _d, files = ar.report_files(root)
        if not files:
            return None
        rep = ar.parse_report(ar.read_pages(files[0]))
        hit = ar.match_items(rep["lines"]).get(item)
        return sum(h["cur_k"] or 0 for h in hit[1]) if hit else None
    except Exception:
        return None


def build(company_folder, tol=1000.0):
    root = Path(company_folder)
    an = aa.Analysis("ARAP", "A/R and A/P analysis")
    import toi_analysis as ta
    t, toi_path, err = ta.load(root)
    year = ta.year_of(toi_path) if toi_path else None
    year_end = date(year, 12, 31) if year else None
    if t is None:
        an.add_note("TOI: " + err)
    an.period = str(year or "")
    S, S_tags = [], []
    for kind, folder_names, toi_codes, gl_rx, audit_items, label in (
            ("AR", {"AR", "A-R", "A_R", "A/R", "RECEIVABLE", "RECEIVABLES"}, ["A18"], r"receivable", ("Trade receivables", "Trade and other receivables"), "A/R"),
            ("AP", {"AP", "A-P", "A_P", "A/P", "PAYABLE", "PAYABLES"}, ["A46"], r"payable", ("Trade and other payables",), "A/P")):
        folder = aa.subfolder(root, folder_names)
        try:
            lst = read_listing(folder, kind)
        except Exception as exc:
            an.add_note(f"{label}: {str(exc).splitlines()[0]}")
            lst = None
        if lst is None:
            S.append([label, "Not provided" if folder is None else "No readable listing (Excel / CSV)", None, None, None, "", ""])
            S_tags.append("warning")
            continue
        total = sum(r["balance"] for r in lst["rows"])
        toi_v = t.c(*toi_codes) if t else None
        gl_v = None
        try:
            gl_v = gl_balance(root, year, gl_rx)
        except Exception as exc:
            an.add_note(f"GL ({label}): {str(exc).splitlines()[0]}")
        afs = next((v for v in (audit_figure(root, i) for i in audit_items) if v is not None), None)
        for src, v in ((f"{label} listing ({lst['file']})", total), (f"TOI {'+'.join(toi_codes)}", toi_v),
                       (f"GL {label} accounts (closing)", gl_v), ("Audit report", afs)):
            diff = None if v is None or src.startswith(label + " listing") else total - v
            st = "Base" if src.startswith(label + " listing") else ("Not provided" if v is None else "Agrees" if abs(diff) <= tol else "Differs")
            S.append([label, src, v, diff, None, st, lst["used"] if st == "Base" else ""])
            S_tags.append({"Base": "section", "Agrees": "total", "Differs": "error"}.get(st, "warning"))

        # aging
        aging = defaultdict(float)
        per = defaultdict(lambda: {"bal": 0.0, "n": 0, "old": 0.0, "buckets": defaultdict(float), "oldest": None})
        for r in lst["rows"]:
            p = per[r["name"].upper()]
            p["bal"] += r["balance"]
            p["n"] += 1
            if lst["bucket_cols"]:
                for col, v in r["buckets"].items():
                    aging[col] += v
                    p["buckets"][col] += v
                    if is_old(col):
                        p["old"] += v
            elif r["date"] and year_end:
                b = age_bucket((year_end - r["date"]).days)
                aging[b] += r["balance"]
                p["buckets"][b] += r["balance"]
                if b == AGE_BUCKETS[-1][1]:
                    p["old"] += r["balance"]
                p["oldest"] = min(p["oldest"] or r["date"], r["date"])
        cols_b = lst["bucket_cols"] or ([b for _, b in AGE_BUCKETS] if aging else [])
        old_total = sum(p["old"] for p in per.values())
        neg = [n for n, p in per.items() if p["bal"] < 0]
        S.append([label, "Aging", None, None, old_total, "Info" if not old_total else "Check",
                  " · ".join(f"{b}: {aging[b]:,.0f}" for b in cols_b) if cols_b else "No aging columns or dates in the listing"])
        S_tags.append("error" if old_total else "even")
        if neg:
            S.append([label, f"{len(neg)} {'customer' if kind == 'AR' else 'supplier'}(s) with a {'credit' if kind == 'AR' else 'debit'} balance",
                      None, None, sum(per[n]["bal"] for n in neg), "Check", ", ".join(neg[:5])])
            S_tags.append("warning")
        rows, tags = [], []
        grand = total or 1
        for n, p in sorted(per.items(), key=lambda x: -abs(x[1]["bal"])):
            note = []
            if p["old"]:
                note.append("Over 1 year: " + ("collectable? bad debt / unrecorded cash receipt" if kind == "AR"
                                                else "still owed? possible fictitious payable / unclaimed expense; WHT on payment"))
            if p["bal"] < 0:
                note.append("Credit balance (advance / overpayment)" if kind == "AR" else "Debit balance (advance to supplier)")
            if p["bal"] / grand > 0.3:
                note.append("Over 30 % of the total: confirm the balance")
            rows.append([n, p["n"], p["bal"], f"{p['bal'] / grand:.1%}"] + [p["buckets"].get(b, 0.0) for b in cols_b] +
                        [p["old"], str(p["oldest"] or ""), "; ".join(note)])
            tags.append("error" if p["old"] else "warning" if note else "even")
        who = "Customer" if kind == "AR" else "Supplier"
        an.view(f"{label} by {who.lower()}", [who, "Lines", "Balance", "Share"] + cols_b + ["Over 1 year", "Oldest date", "Note"],
                rows, tags, money=[2] + list(range(4, 4 + len(cols_b))) + [4 + len(cols_b)],
                desc=f"{label} listing {lst['file']} grouped by {who.lower()}, with aging.")
        an.view(f"{label} lines", [who, "Balance", "Date"] + lst["bucket_cols"],
                [[r["name"], r["balance"], str(r["date"] or "")] + [r["buckets"].get(c, 0.0) for c in lst["bucket_cols"]] for r in lst["rows"]],
                None, money=[1] + list(range(3, 3 + len(lst["bucket_cols"]))), desc=f"Every line of {lst['file']}.")
    an.view("Summary", ["Area", "Source", "Amount (KHR)", "Listing − source", "Over 1 year", "Status", "Detail"], S, S_tags,
                      money=[2, 3, 4], total=False, desc="Listing totals compared with the TOI, GL and audited accounts, and aging.")
    an.views.move_to_end("Summary", last=False)
    return an
