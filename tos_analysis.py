"""TOS (tax on salary) risk analysis from the monthly TOS employee lists.

Reads every monthly TOS workbook (TOS/<year>/*.xlsx, the same files the consolidation uses),
finds the employee columns by their heading text (Khmer or English) and tests four things:

  1. Foreign staff (not Khmer) declared with a base salary of 4,000,000 KHR (≈ USD 1,000) or less,
     month by month and per person.
  2. Foreign staff missing from a month between their first and last declaration
     (declared Jan, Feb, not Mar, again in Apr).
  3. Foreign staff who stop being declared before the last month on file: assumed to have left
     after their last month, so the resignation letter can be requested. Joiners are listed too.
  4. Khmer staff missing from a month between their first and last declaration (no leaver list).

Staff are matched month to month by ID (the VLOOKUP key); when a row has no ID, its name is used.
Months with no TOS file at all are reported once and are not counted as an employee's missing month.

build(company_folder) -> audit_analysis.Analysis (views shown in the Control Center and saved to Excel).
Nothing here touches the GUI.
"""
import re
from collections import Counter, defaultdict, OrderedDict
from pathlib import Path

import audit_analysis as aa

TAXABLE_BASE = 1_500_000   # KHR (≈ USD 375): Khmer staff from this base salary usually pay tax on salary
THRESHOLD = 4_000_000        # KHR, base salary at or below this is flagged for foreign staff
FX = 4000.0                  # KHR per USD, for the USD column only
KHMER_WORDS = ("khmer", "ខ្មែរ", "cambod", "កម្ពុជា")
TOTAL_ROW = re.compile(r"^\s*(grand\s+)?(sub\s*)?total\b|^\s*សរុប", re.I)

# heading tokens, most specific first. Each entry: (tokens that must all appear, tokens that must not)
HEADINGS = {
    "id": [(("លិខិតឆ្លងដែន",), ()), (("passport",), ()), (("អត្តសញ្ញាណ",), ()),
           ((r"\bid\b",), ("រូបវន្ត",)), (("identity",), ()), (("អត្តលេខ",), ()), (("staff no",), ()), (("employee no",), ())],
    "name": [(("ឈ្មោះ", "ឡាតាំង"), ()), (("name", "latin"), ()), (("name", "english"), ()), (("ឈ្មោះ",), ()), ((r"\bname\b",), ())],
    "nationality": [(("សញ្ជាតិ",), ()), (("nationality",), ()), (("citizenship",), ())],
    "gender": [(("ភេទ",), ()), (("gender",), ()), ((r"\bsex\b",), ())],
    "position": [(("មុខតំណែង",), ()), (("មុខងារ",), ()), (("position",), ())],
    "fb_tax": [(("ពន្ធ", "អត្ថប្រយោជន៍"), ()), (("tax", "fringe"), ()), ((r"\bfbt\b",), ())],
    "fb": [(("អត្ថប្រយោជន៍បន្ថែម",), ("ពន្ធ",)), (("fringe",), ("tax",)), ((r"\bfb\b",), ("tax",))],
    "tax": [(("ពន្ធ", "ប្រាក់បៀវត្ស"), ("អត្ថប្រយោជន៍",)), (("tax on salary",), ()), (("salary tax",), ()),
            (("ពន្ធ",), ("អត្ថប្រយោជន៍", "មូលដ្ឋាន")), ((r"\btos\b",), ())],
}
FIELDS = ("id", "name", "nationality", "gender", "position", "salary", "tax", "fb", "fb_tax")
FIELD_LABEL = {"id": "ID", "name": "Name", "nationality": "Nationality", "gender": "Gender", "position": "Position",
               "salary": "Salary base", "tax": "Tax on salary", "fb": "Fringe benefit", "fb_tax": "Tax on FB"}


def _norm(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def _find(labels, key, used):
    for need, avoid in HEADINGS[key]:
        for i, lab in enumerate(labels):
            low = lab.casefold()
            if i in used or not low:
                continue
            if all(re.search(t, low) if t.startswith(r"\b") else t in low for t in need) and not any(a in low for a in avoid):
                return i
    return None


def map_columns(labels, rows, salary_col):
    """{field: column index or None} from the heading labels, with value-based fallbacks."""
    cols, used = {}, {salary_col}
    cols["salary"] = salary_col
    for key in ("fb_tax", "fb", "tax", "nationality", "id", "name", "gender", "position"):
        i = _find(labels, key, used)
        # salary-type numbers live in the amount block (J..N); text fields before it
        if i is not None and key in ("tax", "fb", "fb_tax") and i <= salary_col:
            i = None
        cols[key] = i
        if i is not None:
            used.add(i)
    if cols["nationality"] is None:  # a column whose values say Khmer / ខ្មែរ
        best = None
        for i in range(min(salary_col, max((len(r) for r in rows), default=0))):
            if i in used:
                continue
            hits = sum(1 for r in rows[:300] if i < len(r) and any(w in str(r[i] or "").casefold() for w in KHMER_WORDS))
            if hits and (best is None or hits > best[1]):
                best = (i, hits)
        cols["nationality"] = best[0] if best else None
    return cols


def is_khmer(nationality):
    t = _norm(nationality).casefold()
    return bool(t) and (any(w in t for w in KHMER_WORDS) or t in ("kh", "khm", "cam"))


def _id_key(v):
    t = re.sub(r"\s+", "", str(v or "")).upper()
    t = re.sub(r"\.0$", "", t)
    return t if t not in ("", "-", "N/A", "NA", "NONE", "0") else ""


def read_tos(folder):
    """[(year, month, path, labels, cols, records)], notes. records: dicts with the mapped fields."""
    import consolidate as cons
    layout = cons.LAYOUTS["TOS"]
    out, notes = [], []
    for year, path in cons.files_in(Path(folder)):
        month = aa.file_month(path.name, year)
        if month is None:
            notes.append(f"{path.name}: month not recognised in the file name; rename it like TOS_03_{year}.xlsx")
            continue
        try:
            src = cons.read_source(path, layout)
        except Exception as exc:
            notes.append(f"{path.name}: {str(exc).splitlines()[0]}")
            continue
        labels = aa.two_row_labels(src.headers, src.merges)
        rows = [list(r) for r in src.rows]
        cols = map_columns(labels, rows, layout.first_col - 1)
        recs = []
        for r in rows:
            get = lambda f: (r[cols[f]] if cols.get(f) is not None and cols[f] < len(r) else None)
            name, ident = _norm(get("name")), _id_key(get("id"))
            if not name and not ident:
                continue
            first_text = " ".join(_norm(x) for x in r[:layout.first_col - 1] if isinstance(x, str))
            if TOTAL_ROW.search(first_text) or TOTAL_ROW.search(name):
                continue
            recs.append({"id": ident, "name": name, "nationality": _norm(get("nationality")),
                         "gender": _norm(get("gender")), "position": _norm(get("position")),
                         "salary": aa.num(get("salary")), "tax": aa.num(get("tax")),
                         "fb": aa.num(get("fb")), "fb_tax": aa.num(get("fb_tax"))})
        out.append((int(year) if str(year).isdigit() else year, month, path, labels, cols, recs))
    return out, notes


def analyse(files, threshold=THRESHOLD, fx=FX):
    """The four tests on what read_tos returned. Returns a dict of plain lists (no GUI)."""
    files = [f for f in files if isinstance(f[0], int)]
    months_on_file = sorted({(y, m) for y, m, *_ in files})
    if not months_on_file:
        return None
    first, last = months_on_file[0], months_on_file[-1]
    timeline = aa.month_range(first, last)
    no_file = [ym for ym in timeline if ym not in set(months_on_file)]

    people = OrderedDict()          # key -> dict(id, names Counter, nats Counter, months {ym: rec})
    duplicates, decl = [], []
    for year, month, path, _labels, _cols, recs in sorted(files, key=lambda f: (f[0], f[1], f[2].name)):
        ym = (year, month)
        for rec in recs:
            key = rec["id"] or "NAME:" + rec["name"].casefold()
            p = people.setdefault(key, {"id": rec["id"], "names": Counter(), "nats": Counter(), "months": {}, "files": {}})
            p["names"][rec["name"]] += 1
            if rec["nationality"]:
                p["nats"][rec["nationality"]] += 1
            if ym in p["months"]:
                duplicates.append((ym, key, rec["name"], path.name, p["files"][ym]))
                continue
            p["months"][ym] = rec
            p["files"][ym] = path.name
            decl.append((ym, key, rec, path.name))
    for p in people.values():
        p["name"] = p["names"].most_common(1)[0][0] if p["names"] else ""
        p["nationality"] = p["nats"].most_common(1)[0][0] if p["nats"] else ""
        p["group"] = "Khmer" if is_khmer(p["nationality"]) else ("Foreigner" if p["nationality"] else "Unknown")
        ms = sorted(p["months"])
        p["first"], p["last"] = ms[0], ms[-1]
        span = aa.month_range(p["first"], p["last"])
        p["missing"] = [ym for ym in span if ym not in p["months"] and ym not in no_file]
        p["no_file"] = [ym for ym in span if ym in no_file]
    foreign = lambda p: p["group"] != "Khmer"   # Unknown (blank nationality) is checked with foreigners
    return dict(people=people, decl=decl, duplicates=duplicates, timeline=timeline, months_on_file=months_on_file,
                no_file=no_file, first=first, last=last, threshold=threshold, fx=fx, foreign=foreign)


PAYROLL_WORDS = re.compile(r"salar|payroll|wage|bonus|ប្រាក់ខែ|ប្រាក់បៀវត្ស|ប្រាក់ឈ្នួល", re.I)


def salary_compare(root, decl, fx=FX, tol=0.05):
    """Salary per TOS (12 months) vs GL salary accounts, TOI (B23 + C8) and bank payroll payments.
    Returns (year, monthly rows, tags, summary lines, notes)."""
    import finance_view as fv
    notes = []
    years = defaultdict(set)
    for ym, *_ in decl:
        years[ym[0]].add(ym[1])
    if not years:
        return None
    year = max(years, key=lambda y: (len(years[y]), y))
    tos_m, tax_m = defaultdict(float), defaultdict(float)
    for ym, _k, rec, _f in decl:
        if ym[0] == year:
            tos_m[ym[1]] += rec["salary"] + rec["fb"]
            tax_m[ym[1]] += rec["tax"] + rec["fb_tax"]
    gl_m = None
    gl_dir = aa.subfolder(root, {"GL", "GENERAL LEDGER", "LEDGER"})
    gl_files = [p for p in fv.table_files(gl_dir) if p.suffix.lower() == ".xlsx"] if gl_dir else []
    if gl_files:
        try:
            import sales_reconciliation as sr
            gl_m = defaultdict(float)
            accts = set()
            for r in sr.read_gl(gl_files[0], year):
                if r["code"][:1] in ("5", "6") and PAYROLL_WORDS.search(r["acct"]):
                    gl_m[r["date"].month] += r["debit"] - r["credit"]
                    accts.add(f"{r['code']} {r['acct']}")
            notes.append("GL salary accounts: " + (", ".join(sorted(accts)) or "none found (account names with salary / wage / payroll)"))
        except Exception as exc:
            notes.append(f"GL: {str(exc).splitlines()[0]}")
            gl_m = None
    bank_m = None
    bank_dir = aa.subfolder(root, {"BANK", "BANKS", "BANK STATEMENT", "BANK STATEMENTS"})
    if bank_dir and fv.table_files(bank_dir):
        try:
            t = fv.load_table(fv.table_files(bank_dir), kind="BANK")
            col = {c: i for i, c in enumerate(t.columns)}
            if "Money Out" in col and "Description" in col and t.date_col is not None:
                bank_m = defaultdict(float)
                for r in t.rows:
                    mk = fv.month_key(r[t.date_col])
                    out = r[col["Money Out"]]
                    if not mk or int(mk[:4]) != year or not isinstance(out, (int, float)) or not out:
                        continue
                    if PAYROLL_WORDS.search(str(r[col["Description"]] or "")):
                        cur = str(r[col["Currency"]] or "") if "Currency" in col else ""
                        bank_m[int(mk[5:])] += out * (fx if cur.upper() == "USD" else 1)
            else:
                notes.append("Bank: no Date / Description / Money Out columns")
        except Exception as exc:
            notes.append(f"Bank: {str(exc).splitlines()[0]}")
    toi_salary = None
    try:
        import toi_analysis as ta
        t, path, _err = ta.load(root)
        if t is not None and (ta.year_of(path) in (None, year)):
            toi_salary = t.c("B23", "C8")
    except Exception:
        pass
    rows, tags = [], []
    for m in range(1, 13):
        a = tos_m.get(m, 0.0)
        g = gl_m.get(m, 0.0) if gl_m is not None else None
        b = bank_m.get(m, 0.0) if bank_m is not None else None
        diffs = [x for x in (g, b) if x is not None]
        st = "Not provided" if not diffs else ("Agrees" if all(abs(a - x) <= max(1000, tol * max(a, x)) for x in diffs) else "Differs")
        rows.append([f"{year}-{m:02d}", a, tax_m.get(m, 0.0), g, b, None if g is None else a - g, None if b is None else a - b, st])
        tags.append({"Agrees": "total", "Differs": "error"}.get(st, "warning"))
    tos_total = sum(tos_m.values())
    lines = [("TOS returns (12 months, salary + fringe benefits)", tos_total)]
    lines.append(("General ledger salary / wage accounts", sum(gl_m.values()) if gl_m is not None else None))
    lines.append(("TOI salary (B23 + production salary C8)", toi_salary))
    lines.append(("Bank payroll payments (description: salary / payroll / wage)", sum(bank_m.values()) if bank_m is not None else None))
    return year, rows, tags, lines, notes


def build(company_folder, threshold=THRESHOLD, fx=FX):
    root = Path(company_folder)
    an = aa.Analysis("TOS", "TOS salary tax risk")
    tos_dir = aa.subfolder(root, {"TOS"})
    if tos_dir is None:
        an.add_note("No TOS folder: create TOS/<year> and put the monthly TOS Excel files in it.")
        an.add_summary_only()
        return an
    files, notes = read_tos(tos_dir)
    for n in notes:
        an.add_note(n)
    res = analyse(files, threshold, fx)
    if res is None:
        an.add_note("No readable monthly TOS file with a month in its name.")
        an.add_summary_only()
        return an
    people, timeline = res["people"], res["timeline"]
    foreign = res["foreign"]
    M = aa.month_label
    period = f"{M(res['first'])} to {M(res['last'])}"
    an.period = period
    thr = f"{threshold:,.0f} KHR (≈ USD {threshold / fx:,.0f})"

    # ---- mapping check
    map_rows, map_tags = [], []
    missing_fields = Counter()
    for year, month, path, labels, cols, recs in files:
        if not isinstance(year, int):
            continue
        cells = []
        for f in FIELDS:
            i = cols.get(f)
            cells.append(f"{aa.col_letter(i)}: {labels[i]}" if i is not None and i < len(labels) else "not found")
            if i is None and f in ("id", "name", "nationality", "salary", "tax"):
                missing_fields[f] += 1
        bad = cols.get("nationality") is None or (cols.get("id") is None and cols.get("name") is None)
        map_rows.append([M((year, month)), path.name, len(recs)] + cells)
        map_tags.append("error" if bad else "warning" if cols.get("id") is None or cols.get("tax") is None else "even")

    # ---- 1. foreigners at or below threshold
    low_rows, low_tags, low_people = [], [], defaultdict(list)
    for ym, key, rec, fname in res["decl"]:
        p = people[key]
        if foreign(p) and rec["salary"] <= threshold:
            low_people[key].append((ym, rec["salary"]))
            low_rows.append([aa.month_key(ym), p["id"], p["name"], p["nationality"] or "(blank)", rec["position"],
                             rec["salary"], rec["salary"] / fx, rec["tax"], rec["fb"], fname])
            low_tags.append("error" if rec["salary"] < threshold else "warning")
    low_rows.sort(key=lambda r: (r[0], str(r[2]).casefold()))
    low_tags = ["error" if r[5] < threshold else "warning" for r in low_rows]
    lowp_rows = []
    for key, items in low_people.items():
        p = people[key]
        sal = [s for _, s in items]
        lowp_rows.append([p["id"], p["name"], p["nationality"] or "(blank)", len(p["months"]), len(items),
                          f"{len(items) / len(p['months']):.0%}", min(sal), max(sal),
                          aa.month_list([ym for ym, _ in items]),
                          "Request employment contract, work permit and payroll; check salary is not under-declared"])
    lowp_rows.sort(key=lambda r: (-r[4], str(r[1]).casefold()))

    # ---- 2. foreigners missing months; 4. Khmer missing months
    def gap_rows(group_test):
        rows = []
        for p in people.values():
            if group_test(p) and p["missing"]:
                rows.append([p["id"], p["name"], p["nationality"] or "(blank)", M(p["first"]), M(p["last"]),
                             len(p["months"]), len(p["missing"]), aa.month_list(p["missing"]),
                             "Ask why not declared in " + aa.month_list(p["missing"]) + " (unpaid leave, transfer, missed declaration?)"])
        rows.sort(key=lambda r: (-r[6], str(r[1]).casefold()))
        return rows
    gapf = gap_rows(foreign)
    gapk = gap_rows(lambda p: p["group"] == "Khmer")
    # Khmer gaps with each month's salary and tax: staff with taxable salary first, then the others
    span_cols = [aa.month_key(ym) for ym in timeline]
    gapk_rows, gapk_tags = [], []
    for p in people.values():
        if p["group"] != "Khmer" or not p["missing"]:
            continue
        sal = [r["salary"] for r in p["months"].values()]
        tax = [r["tax"] for r in p["months"].values()]
        taxable = max(sal or [0]) >= TAXABLE_BASE or sum(tax) > 0
        cells = []
        for ym in timeline:
            if ym in p["months"]:
                cells.append(f"{p['months'][ym]['salary']:,.0f} / {p['months'][ym]['tax']:,.0f}")
            elif ym in p["missing"]:
                cells.append("MISSING")
            else:
                cells.append("")
        gapk_rows.append([p["id"], p["name"], "Taxable" if taxable else "Below tax", len(p["missing"]), aa.month_list(p["missing"]),
                          sum(sal) / len(sal) if sal else 0, sum(tax), M(p["first"]), M(p["last"])] + cells +
                         ["Ask why not declared in " + aa.month_list(p["missing"]) + (": salary tax may be missing" if taxable else "")])
        gapk_tags.append("error" if taxable else "even")
    order = sorted(range(len(gapk_rows)), key=lambda i: (gapk_rows[i][2] != "Taxable", -gapk_rows[i][5], str(gapk_rows[i][1]).casefold()))
    gapk_rows, gapk_tags = [gapk_rows[i] for i in order], [gapk_tags[i] for i in order]

    # ---- 3. foreigners who stop / join
    left_rows, join_rows = [], []
    for p in people.values():
        if not foreign(p):
            continue
        if p["last"] < res["last"]:
            after = aa.next_month(p["last"])
            left_rows.append([p["id"], p["name"], p["nationality"] or "(blank)", M(p["first"]), M(p["last"]), len(p["months"]),
                              M(after), f"Request resignation / termination letter dated about {M(p['last'])} – {M(after)}"])
        if p["first"] > res["first"]:
            join_rows.append([p["id"], p["name"], p["nationality"] or "(blank)", M(p["first"]), M(p["last"]), len(p["months"]),
                              "Request employment contract and work permit from " + M(p["first"])])
    left_rows.sort(key=lambda r: (aa.parse_label(r[4]), str(r[1]).casefold()))
    join_rows.sort(key=lambda r: (aa.parse_label(r[3]), str(r[1]).casefold()))

    # ---- monthly headcount
    by_month = defaultdict(list)
    for ym, key, rec, _f in res["decl"]:
        by_month[ym].append((people[key], rec))
    head_rows, head_tags, prev_f = [], [], None
    for ym in timeline:
        items = by_month.get(ym, [])
        fset = {id(p) for p, _ in items if foreign(p)}
        joined = sum(1 for p, _ in items if foreign(p) and p["first"] == ym and ym != res["first"])
        left = sum(1 for p in people.values() if foreign(p) and p["last"] == aa.prev_month(ym) and ym <= res["last"])
        low = sum(1 for p, r in items if foreign(p) and r["salary"] <= threshold)
        head_rows.append([aa.month_key(ym), "No TOS file" if ym in res["no_file"] else "",
                          len(items), sum(1 for p, _ in items if p["group"] == "Khmer"), len(fset),
                          (len(fset) - prev_f) if prev_f is not None and ym not in res["no_file"] else None,
                          joined, left, low,
                          sum(r["salary"] for _, r in items), sum(r["tax"] for _, r in items),
                          sum(r["fb"] for _, r in items), sum(r["fb_tax"] for _, r in items)])
        head_tags.append("error" if ym in res["no_file"] else "warning" if low or left else "even")
        if ym not in res["no_file"]:
            prev_f = len(fset)

    # ---- matrix and data checks
    mcols = [aa.month_key(ym) for ym in timeline]
    matrix, mtags = [], []
    for p in sorted(people.values(), key=lambda p: (p["group"] == "Khmer", str(p["name"]).casefold())):
        cells = []
        for ym in timeline:
            if ym in p["months"]:
                cells.append(p["months"][ym]["salary"])
            elif p["first"] <= ym <= p["last"]:
                cells.append("NO FILE" if ym in res["no_file"] else "MISSING")
            else:
                cells.append("")
        status = "Missing month" if p["missing"] else ("Left after " + M(p["last"]) if p["last"] < res["last"] and foreign(p) else "")
        matrix.append([p["id"], p["name"], p["nationality"] or "(blank)", p["group"], len(p["months"]), status] + cells)
        mtags.append("error" if p["missing"] and foreign(p) else "warning" if status else "even")

    checks = []
    for ym, key, name, f1, f0 in res["duplicates"]:
        checks.append(["Declared twice in one month", people[key]["id"], name, M(ym), f"{f0} and {f1}",
                       "Only the first row is used. Check for an amended return or a duplicate line."])
    for p in people.values():
        if len(p["names"]) > 1 and p["id"]:
            checks.append(["Same ID, different names", p["id"], " / ".join(p["names"]), "", "",
                           "Confirm this is one person (spelling) and not an ID reused for another employee."])
        if len(p["nats"]) > 1:
            checks.append(["Nationality changes", p["id"], p["name"], "", " / ".join(p["nats"]), "Confirm the nationality used for the tax rate."])
        if not p["id"]:
            checks.append(["No ID declared", "", p["name"], aa.month_list(sorted(p["months"])), "", "Matched by name only; ask for the ID / passport number."])
        if p["group"] == "Unknown":
            checks.append(["Nationality blank", p["id"], p["name"], aa.month_list(sorted(p["months"])), "", "Checked together with foreign staff; confirm the nationality."])
    by_name = defaultdict(set)
    for p in people.values():
        if p["id"]:
            by_name[p["name"].casefold()].add(p["id"])
    for name, ids in by_name.items():
        if len(ids) > 1 and name:
            checks.append(["Same name, different IDs", ", ".join(sorted(ids)), name.title(), "", "",
                           "One person declared under two IDs could hide a missing month or a leaver."])

    all_rows = []
    for ym, key, rec, fname in sorted(res["decl"], key=lambda d: (d[0], str(people[d[1]]["name"]).casefold())):
        p = people[key]
        all_rows.append([aa.month_key(ym), p["id"], p["name"], p["nationality"] or "(blank)", p["group"], rec["gender"], rec["position"],
                         rec["salary"], rec["tax"], rec["fb"], rec["fb_tax"], fname])

    # ---- summary
    nf = [p for p in people.values() if foreign(p)]
    nk = [p for p in people.values() if p["group"] == "Khmer"]
    S = []
    def s(area, result, count, action, sev):
        S.append(([area, result, count, action], sev))
    s("Period on file", f"{period}: {len(res['months_on_file'])} month(s) with a TOS file, "
      f"{len(people):,} different staff ({len(nk):,} Khmer, {len(nf):,} foreign or blank nationality)", len(people), "", "section")
    if res["no_file"]:
        s("Months with no TOS file", aa.month_list(res["no_file"]), len(res["no_file"]),
          "Add the missing TOS returns; staff are not marked missing for these months", "error")
    s("1. Foreign staff at or below " + thr,
      f"{len(low_people):,} person(s), {len(low_rows):,} person-month(s)" if low_rows else "None found",
      len(low_people), "See '1 Foreign ≤ 4M monthly' and '1 by person'. Compare with work permit, contract and bank payroll; "
      "salary below the market rate for foreign staff is a salary-tax risk.", "error" if low_rows else "total")
    s("2. Foreign staff with a missing month",
      f"{len(gapf):,} person(s), {sum(r[6] for r in gapf):,} missing month(s)" if gapf else "None found",
      len(gapf), "See '2 Foreign missing months'. Ask for the reason and the payroll of those months.", "error" if gapf else "total")
    s("3. Foreign staff who stopped being declared",
      f"{len(left_rows):,} person(s) last declared before {M(res['last'])}" if left_rows else "None found",
      len(left_rows), "See '3 Foreign left'. Request the resignation / termination letter for the month after the last declaration.",
      "warning" if left_rows else "total")
    s("    Foreign staff who joined after " + M(res["first"]), f"{len(join_rows):,} person(s)" if join_rows else "None",
      len(join_rows), "See '3 Foreign joined'. Request contract and work permit.", "even")
    s("4. Khmer staff with a missing month",
      f"{len(gapk):,} person(s), {sum(r[6] for r in gapk):,} missing month(s)" if gapk else "None found",
      len(gapk), "See '4 Khmer missing months'. Ask why they were not declared in those months.", "warning" if gapk else "total")
    if checks:
        s("Data checks", f"{len(checks):,} item(s): duplicates, ID / name differences, blank ID or nationality", len(checks),
          "See 'Data checks'.", "warning")
    for f, n in missing_fields.items():
        s("Column not found", f"{FIELD_LABEL[f]} column not recognised in {n} file(s)", n,
          "See 'Column mapping'. Results that need this column are incomplete.", "error" if f in ("nationality", "salary") else "warning")
    tot_sal = sum(r["salary"] for _, _, r, _ in res["decl"])
    tot_tax = sum(r["tax"] for _, _, r, _ in res["decl"])
    s("Total declared", f"Salary {tot_sal:,.0f} KHR; tax on salary {tot_tax:,.0f} KHR", None, "", "even")
    comp = salary_compare(root, res["decl"], fx)
    if comp:
        cyear, crow, ctag, clines, cnotes = comp
        for n in cnotes:
            an.add_note(n)
        base = clines[0][1]
        for label, v in clines:
            if label.startswith("TOS"):
                s(f"Salary {cyear}: {label}", f"{base:,.0f} KHR", None, "Base for the comparison below", "section")
                continue
            if v is None:
                s(f"Salary {cyear}: {label}", "Not provided", None, "", "even")
                continue
            d = base - v
            ok = abs(d) <= max(1000, 0.05 * max(abs(base), abs(v)))
            s(f"Salary {cyear}: {label}", f"{v:,.0f} KHR; TOS − this = {d:+,.0f}", None,
              "" if ok else "Differs by more than 5 %: unreported staff / salary, or salary booked elsewhere. See 'Salary vs GL, TOI, bank'.",
              "total" if ok else "error")

    an.view("Summary", ["Area", "Result", "Count", "Suggested audit action"], [r for r, _ in S], [t for _, t in S],
            money=[], desc=f"TOS salary tax risk, {period}. Threshold {thr}. Staff are matched by ID (name when no ID).")
    if comp:
        an.view("Salary vs GL, TOI, bank", ["Month", "TOS salary + FB", "TOS tax + FB tax", "GL salary", "Bank payroll paid",
                                            "TOS − GL", "TOS − bank", "Status"], crow, ctag, money=[1, 2, 3, 4, 5, 6], month_col=0,
                desc=f"{cyear}: salary declared in TOS each month vs GL salary accounts and bank payroll payments (5 % tolerance).")
    an.view("Monthly headcount", ["Month", "Note", "Total staff", "Khmer", "Foreign", "Foreign change", "Foreign joined",
                                  "Foreign left", f"Foreign ≤ {threshold / 1e6:g}M", "Salary base", "Tax on salary", "Fringe benefit", "Tax on FB"],
            head_rows, head_tags, money=[9, 10, 11, 12], month_col=0,
            desc="Number of staff each month; Foreign left = last declared in the previous month.")
    an.view("1 Foreign ≤ 4M monthly", ["Month", "ID", "Name", "Nationality", "Position", "Salary base (KHR)", "Salary (USD)",
                                        "Tax on salary", "Fringe benefit", "File"], low_rows, low_tags, money=[5, 6, 7, 8], month_col=0,
            desc=f"Foreign staff declared with salary ≤ {thr}, each month.")
    an.view("1 by person", ["ID", "Name", "Nationality", "Months declared", "Months ≤ 4M", "Share", "Lowest salary", "Highest salary",
                            "Months ≤ 4M (list)", "Suggested action"], lowp_rows, ["error"] * len(lowp_rows), money=[6, 7],
            desc="One line per foreign person with at least one month at or below the threshold.")
    an.view("2 Foreign missing months", ["ID", "Name", "Nationality", "First declared", "Last declared", "Months declared",
                                          "Missing", "Missing months", "Suggested action"], gapf, ["error"] * len(gapf),
            desc="Foreign staff absent in a month between their first and last declaration.")
    an.view("3 Foreign left", ["ID", "Name", "Nationality", "First declared", "Last declared", "Months declared",
                               "Assumed left from", "Request"], left_rows, ["warning"] * len(left_rows),
            desc=f"Foreign staff whose last declaration is before {M(res['last'])}: request the resignation letter.")
    an.view("3 Foreign joined", ["ID", "Name", "Nationality", "First declared", "Last declared", "Months declared", "Request"],
            join_rows, ["even"] * len(join_rows), desc=f"Foreign staff first declared after {M(res['first'])}.")
    an.view("4 Khmer missing months", ["ID", "Name", "Salary level", "Missing", "Missing months", "Average salary", "Tax on salary (total)",
                                        "First declared", "Last declared"] + span_cols + ["Suggested action"], gapk_rows, gapk_tags,
            money=[5, 6], wide=True,
            desc=f"Khmer staff absent in a month between first and last declaration. Each month shows salary / tax. "
                 f"Taxable (base ≥ {TAXABLE_BASE:,.0f} KHR or tax paid) first, in red; leavers are not listed.")
    an.view("Staff × month", ["ID", "Name", "Nationality", "Group", "Months", "Status"] + mcols, matrix, mtags,
            money=list(range(6, 6 + len(mcols))), wide=True, desc="Salary declared per person and month; MISSING = gap between first and last month.")
    an.view("Data checks", ["Check", "ID", "Name", "Month(s)", "Detail", "Suggested action"], checks, ["warning"] * len(checks),
            desc="Duplicates and ID / name / nationality differences.")
    an.view("All declarations", ["Month", "ID", "Name", "Nationality", "Group", "Gender", "Position", "Salary base", "Tax on salary",
                                 "Fringe benefit", "Tax on FB", "File"], all_rows, None, money=[7, 8, 9, 10], month_col=0,
            desc="Every employee line read from the TOS files.")
    an.view("Column mapping", ["Month", "File", "Staff rows"] + [FIELD_LABEL[f] for f in FIELDS], map_rows, map_tags,
            desc="Which column of each TOS file was used for each field. Check this first if a result looks wrong.")
    return an
