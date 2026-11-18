"""Audited financial statements (Audit Report folder, PDF or Word .docx) compared with the TOI.

Reads the text of the report (a PDF with a text layer, or a Word file: no conversion needed; a scanned
PDF has no text and must be OCR'd or exported again as text PDF / Word). Then:
  * finds the statement of financial position, profit or loss and cash flows, and reads every line
    with its current / prior-year amounts (US$ and / or KHR columns, KHR'000 handled);
  * compares the main lines with the TOI (balance sheet A-codes, income statement B-codes) and with the
    TOI cash flow (operating / investing / financing, asset purchases);
  * detects the audit opinion (unmodified, qualified, adverse, disclaimer, going concern, emphasis of
    matter) and indexes the notes and key words (related parties, contingencies, subsequent events, tax).

build(company_folder) -> audit_analysis.Analysis.  Nothing here touches the GUI.
"""
import re
from pathlib import Path

import audit_analysis as aa

FX_DEFAULT = 4000.0
NUM = re.compile(r"\(\s*[\d,]+(?:\.\d+)?\s*\)|-?\d[\d,]*(?:\.\d+)?|(?<!\w)[-–—](?!\w)")
STATEMENTS = [
    ("BS", re.compile(r"statement of financial position|balance sheet", re.I)),
    ("IS", re.compile(r"statement of (profit or loss|comprehensive income|income)|income statement|profit and loss statement", re.I)),
    ("CF", re.compile(r"statement of cash ?flows?|cash ?flow statement", re.I)),
    ("EQ", re.compile(r"statement of changes in (equity|owner)", re.I)),
    ("NOTES", re.compile(r"^\s*notes to (the )?financial statements", re.I)),
]
STATEMENT_NAME = {"BS": "Balance sheet", "IS": "Income statement", "CF": "Cash flow", "EQ": "Changes in equity", "NOTES": "Notes"}

# (statement, label pattern, TOI codes or cash-flow key, item name, compare absolute values, add up all matching lines)
MAP = [
    ("BS", r"^total assets", ["A0"], "Total assets", False, False),
    ("BS", r"total non.?current assets", ["A1"], "Non-current (long-term) assets", False, False),
    ("BS", r"total current assets", ["A13"], "Current assets", False, False),
    ("BS", r"total (equity and liabilities|liabilities and (shareholders.? |owners.? )?equity)", ["A28"], "Total equity and liabilities", False, False),
    ("BS", r"total (shareholders.? |owners.? )?equity|^total capital and reserves|net (assets|liabilities)$", ["A29"], "Equity", False, False),
    ("BS", r"total non.?current liabilities", ["A37"], "Non-current liabilities", False, False),
    ("BS", r"total current liabilities", ["A42"], "Current liabilities", False, False),
    ("BS", r"^total liabilities$", ["A37", "A42"], "Total liabilities", False, False),
    ("BS", r"property,? plant|fixed assets|plant and equipment", ["A2", "A3", "A4", "A5", "A6", "A7"], "Property, plant and equipment", False, True),
    ("BS", r"intangible", ["A8"], "Intangible assets", False, True),
    ("BS", r"right.of.use", ["A10"], "Right-of-use assets", False, True),
    ("BS", r"inventor", ["A14", "A15", "A16", "A17"], "Inventories", False, True),
    ("BS", r"trade and other receivables", ["A18", "A19"], "Trade and other receivables", False, False),
    ("BS", r"^trade receivables|accounts receivable", ["A18"], "Trade receivables", False, False),
    ("BS", r"other receivables|amounts? due from", ["A19"], "Other receivables / due from", False, True),
    ("BS", r"prepayment|deposits", ["A20"], "Prepayments and deposits", False, True),
    ("BS", r"^cash|cash and (cash equivalents|bank)|cash on hand|cash at bank", ["A21", "A22"], "Cash and bank", False, False),
    ("BS", r"share capital|paid.up capital|registered capital|^capital$", ["A30"], "Share capital", False, False),
    ("BS", r"retained earnings|accumulated (losses|deficit)|retained (losses|profits)", ["A35", "A36"], "Retained earnings / losses", False, False),
    ("BS", r"amounts? due to|due to (a )?(shareholder|director|related)|loans? from (shareholder|director|related)", ["A38", "A45"], "Due to related parties", False, True),
    ("BS", r"borrowing|bank loan|loans? payable|term loan", ["A39", "A44"], "Borrowings", False, True),
    ("BS", r"trade and other payables|accounts payable|trade payables|other payables", ["A46"], "Trade and other payables", False, True),
    ("BS", r"accru", ["A48"], "Accruals", False, True),
    ("BS", r"income tax payable|current (income )?tax liabilit|tax payable", ["A50"], "Income tax payable", False, True),
    ("IS", r"^(revenue|sales|turnover)", ["B0"], "Revenue", False, False),
    ("IS", r"cost of (sales|goods|revenue|services)", ["B4", "B5", "B6"], "Cost of sales", True, False),
    ("IS", r"^gross (profit|loss)", ["B7"], "Gross profit", False, False),
    ("IS", r"^other (operating )?income", ["B8", "B12"], "Other income", False, True),
    ("IS", r"administrative|general and admin|operating expenses|selling|distribution|other (operating )?expenses", ["B22"],
     "Operating expenses", True, True),
    ("IS", r"finance (costs?|expenses?)|interest expense", ["B43", "B44", "B45"], "Finance costs", True, True),
    ("IS", r"(profit|loss|\(loss\)).{0,25}before (income )?tax", ["B46"], "Profit / (loss) before tax", False, False),
    ("IS", r"^income tax|tax expense", ["B47"], "Income tax expense", True, False),
    ("IS", r"(profit|loss|\(loss\)).{0,15}for the (year|period)|net (profit|loss)|(profit|loss) after tax", ["B48"], "Profit / (loss) for the year", False, False),
    ("CF", r"net cash.*operating", "op", "Net cash from operating activities", False, False),
    ("CF", r"net cash.*investing", "inv", "Net cash from investing activities", False, False),
    ("CF", r"net cash.*financing", "fin", "Net cash from financing activities", False, False),
    ("CF", r"(purchases?|acquisitions?|additions?) (of|to) (property|plant|fixed|equipment)", "capex", "Purchase of property, plant and equipment", True, True),
    ("CF", r"^depreciation", "dep", "Depreciation", True, True),
    ("CF", r"cash and cash equivalents at (the )?end|cash at (the )?end", "cash_end", "Cash at end of year", False, False),
]
OPINION = [
    ("Disclaimer of opinion", r"disclaimer of opinion|do not express an opinion", "error"),
    ("Adverse opinion", r"adverse opinion", "error"),
    ("Qualified opinion", r"(?<!un)qualified opinion|basis for qualified|except for the (effects?|possible effects?)", "error"),
    ("Material uncertainty – going concern", r"material uncertainty.{0,40}going concern", "error"),
    ("Emphasis of matter", r"emphasis of (a )?matter", "warning"),
    ("Other matter", r"other matter", "warning"),
]
KEYWORDS = [("Related parties", r"related part"), ("Going concern", r"going concern"), ("Contingent liabilities", r"contingen"),
            ("Commitments", r"commitment"), ("Subsequent events", r"subsequent event|events after the reporting"),
            ("Tax (uncertain / assessment / audit)", r"tax (audit|assessment|reassessment)|uncertain tax|tax contingenc|tax penalt"),
            ("Borrowings / guarantees", r"guarantee|collateral|pledge"), ("Prior-year restatement", r"restate|prior.year adjustment|prior period error"),
            ("Dividends", r"dividend")]


def report_files(root):
    import engine
    root = Path(root)
    d = next((p for p in sorted(root.iterdir()) if p.is_dir() and p.name.upper() in engine.AUDIT_REPORT_NAMES), None) if root.is_dir() else None
    if d is None:
        return None, []
    files = sorted(p for p in d.rglob("*") if p.is_file() and p.suffix.lower() in (".pdf", ".docx")
                   and not any(part.startswith((".", "~$")) for part in p.relative_to(d).parts))
    return d, files


def read_pages(path):
    """[(page number, text)]. Word files: tables are written one row per line, cells tab-separated."""
    path = Path(path)
    if path.suffix.lower() == ".docx":
        import docx
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        doc = docx.Document(str(path))
        lines = []
        for child in doc.element.body.iterchildren():      # paragraphs and tables in document order
            tag = child.tag.rsplit("}", 1)[-1]
            if tag == "p":
                lines.append(Paragraph(child, doc).text)
            elif tag == "tbl":
                for row in Table(child, doc).rows:
                    cells, last = [], None
                    for c in row.cells:
                        if c._tc is not last:
                            cells.append(c.text.strip())
                        last = c._tc
                    lines.append("\t".join(cells))
        return [(1, "\n".join(lines))]
    import pdfplumber
    out = []
    with pdfplumber.open(str(path)) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            out.append((i, page.extract_text() or ""))
    return out


def _value(tok):
    t = tok.strip()
    if t in ("-", "–", "—"):
        return 0.0
    neg = t.startswith("(") or t.startswith("-")
    v = float(t.strip("()- ").replace(",", "") or 0)
    return -v if neg else v


def parse_line(line):
    """(label, note, [numbers]) or None. Note references are the small first number before the amounts."""
    text = line.replace("\t", "  ").strip()
    m = NUM.search(text)
    if not m:
        return None
    label = text[:m.start()].strip(" .:")
    if len(re.sub(r"[^A-Za-zក-៿]", "", label)) < 3:
        return None
    toks = NUM.findall(text[m.start():])
    if re.search(r"[A-Za-z]{3,}", text[m.start():]):   # words after the numbers: not a table line
        return None
    nums = [_value(t) for t in toks]
    note = None
    if len(nums) in (3, 5) and re.fullmatch(r"\d{1,2}(\.\d{1,2})?", toks[0].strip()) and float(toks[0]) < 60:
        note, nums = toks[0].strip(), nums[1:]
    if len(nums) < 1 or all(abs(n) < 100 for n in nums) and len(nums) <= 2 and all(float(n).is_integer() for n in nums) \
            and re.fullmatch(r"(19|20)\d\d", toks[-1].strip()):
        return None
    return label, note, nums


def parse_report(pages, fx=None):
    """Statement lines, year, currencies and full text from the report pages."""
    full = "\n".join(t for _, t in pages)
    low = full.lower()
    ym = re.search(r"(?:year|period) ended\s+(\d{1,2}\s+\w+\s+(20\d\d)|\w+\s+\d{1,2},?\s+(20\d\d))", full, re.I)
    year = int(ym.group(2) or ym.group(3)) if ym else None
    rate = None
    m = re.search(r"(?:US\$|USD)\s*1\s*=\s*(?:KHR|riel|៛)\s*([\d,]{4,6})", full, re.I) or \
        re.search(r"(?:KHR|riel)\s*([\d,]{4,6})\s*(?:per|to|=|:)\s*(?:US\$|USD)\s*1", full, re.I)
    if m:
        rate = float(m.group(1).replace(",", ""))
    fx = fx or rate or FX_DEFAULT
    lines, state = [], None
    for page, text in pages:
        page_low = text.lower()
        usd = bool(re.search(r"us\$|usd", page_low))
        khr = bool(re.search(r"khr|riel|៛", page_low))
        khr_thousand = bool(re.search(r"khr\s*['’`]?\s*000|riel\s*['’`]?\s*000|'000\s*khr", page_low))
        usd_thousand = bool(re.search(r"(us\$|usd)\s*['’`]?\s*000", page_low))
        for raw in text.splitlines():
            stripped = raw.strip()
            heading = next((k for k, rx in STATEMENTS if rx.search(stripped)), None)
            if heading and len(stripped) < 90 and not re.search(r"\d{1,3}\s*$", stripped):
                state = heading
                continue
            if state not in ("BS", "IS", "CF"):
                continue
            p = parse_line(raw)
            if p is None:
                # a wrapped label: words only, continuing the previous line ("Property," / "plant and" / "equipment")
                prev = lines[-1] if lines and lines[-1]["page"] == page and lines[-1].get("open") else None
                if prev and stripped and not NUM.search(stripped) and len(stripped) <= 45 and \
                        (stripped[:1].islower() or prev["label"].endswith(("-", ",", "and", "of", "to", "from", "in"))):
                    joiner = "" if prev["label"].endswith("-") else " "
                    prev["label"] = (prev["label"] + joiner + stripped).strip()
                elif prev:
                    prev["open"] = False
                continue
            label, note, nums = p
            if label.strip().lower() in ("note", "notes", "description", "particulars"):
                continue
            if lines and lines[-1]["page"] == page:
                lines[-1]["open"] = False
            cur_k = pri_k = None
            if usd and khr and len(nums) >= 4:
                cur_u, pri_u, cur_k, pri_k = nums[-4:]
                mult = 1000 if khr_thousand else 1
                cur_k, pri_k = cur_k * mult, pri_k * mult
                cur, pri, ccy = cur_u, pri_u, "US$ + KHR"
            else:
                cur = nums[0]
                pri = nums[1] if len(nums) > 1 else None
                if usd and not khr:
                    ccy, mult = "US$", (1000 if usd_thousand else 1)
                    cur_k = cur * mult * fx
                    pri_k = pri * mult * fx if pri is not None else None
                else:
                    ccy, mult = "KHR", (1000 if khr_thousand else 1)
                    cur_k, pri_k = cur * mult, (pri * mult if pri is not None else None)
            lines.append(dict(stmt=state, label=label, note=note, cur=cur, pri=pri, ccy=ccy, cur_k=cur_k, pri_k=pri_k, page=page, open=True))
    return dict(lines=lines, year=year, fx=fx, fx_found=rate is not None, text=full, low=low, pages=pages)


def match_items(lines):
    """{item name: (definition, matched lines)} using MAP; first matching line wins unless 'add up'."""
    out = {}
    used = set()
    for stmt, pattern, target, name, absolute, add_all in MAP:
        rx = re.compile(pattern, re.I)
        hits = [i for i, l in enumerate(lines) if l["stmt"] == stmt and i not in used and rx.search(l["label"].strip())]
        if not hits:
            continue
        hits = hits if add_all else hits[:1]
        used.update(hits)
        out[name] = ((stmt, target, absolute), [lines[i] for i in hits])
    return out


def build(company_folder, fx=None, tol=0.01):
    root = Path(company_folder)
    an = aa.Analysis("AUDITREPORT", "Audit report vs TOI")
    folder, files = report_files(root)
    if folder is None:
        an.add_note("No Audit Report folder: create 'Audit Report' and put the audited financial statements (PDF or Word) in it.")
        an.add_summary_only()
        return an
    if not files:
        an.add_note(f"{folder.name} has no PDF or Word (.docx) file.")
        an.add_summary_only()
        return an
    path = files[0]
    if len(files) > 1:
        an.add_note(f"{len(files)} files in {folder.name}; {path.name} is used (keep only the final signed report there).")
    try:
        pages = read_pages(path)
    except Exception as exc:
        an.add_note(f"{path.name}: cannot be read ({str(exc).splitlines()[0]})")
        an.add_summary_only()
        return an
    chars = sum(len(t.strip()) for _, t in pages)
    if chars < 300:
        an.add_note(f"{path.name}: no readable text: it is a scanned PDF. Ask for the text PDF or Word version, "
                    "or OCR it (Mac: open in Preview / Adobe and export as text PDF), then click Refresh.")
        an.add_summary_only()
        return an
    rep = parse_report(pages, fx)
    an.period = str(rep["year"] or "")
    lines = rep["lines"]

    # ---- TOI side
    import toi_analysis as ta
    t, toi_path, err = ta.load(root)
    cf = None
    if t is not None:
        try:
            step10 = ta.read_step10(toi_path)
        except Exception:
            step10 = None
        _l, cf = ta.cashflow_lines(t, step10)
        cf["capex"] = step10["additions"] if step10 and step10.get("found") and step10.get("additions") else cf["derived_capex"]
        cf["cash_end"] = t.c("A21", "A22")
        if rep["year"] and ta.year_of(toi_path) and ta.year_of(toi_path) != rep["year"]:
            an.add_note(f"Audit report year {rep['year']} but TOI file {toi_path.name}: check they are the same year.")
    else:
        an.add_note("TOI: " + err)

    # ---- comparison
    matched = match_items(lines)
    cmp_rows, cmp_tags, big = [], [], []
    for stmt, pattern, target, name, absolute, add_all in MAP:
        if name not in matched:
            continue
        (_s, _t, _a), hits = matched[name]
        ar_c = sum(h["cur_k"] or 0 for h in hits)
        ar_p = sum(h["pri_k"] or 0 for h in hits) if all(h["pri_k"] is not None for h in hits) else None
        if isinstance(target, list):
            toi_c = t.c(*target) if t else None
            toi_p = t.p(*target) if t else None
            codes = "+".join(target)
        else:
            toi_c = cf.get(target) if cf else None
            toi_p = None
            codes = {"op": "TOI cash flow A", "inv": "TOI cash flow B", "fin": "TOI cash flow C", "capex": "Step 10 / derived",
                     "dep": "B36+C9+C12", "cash_end": "A21+A22"}[target]
            if target == "dep" and t:
                toi_c = t.c("B36", "C9", "C12")
        a, b = (abs(ar_c), abs(toi_c)) if absolute and toi_c is not None else (ar_c, toi_c)
        if toi_c is None:
            st, diff = "No TOI", None
        else:
            diff = a - b
            limit = max(1000.0, tol * max(abs(a), abs(b)))
            st = "Agrees" if abs(diff) <= limit else ("Differs (over 5 %)" if abs(diff) > 0.05 * max(abs(a), abs(b), 1) else "Differs")
        pri_diff = None if ar_p is None or toi_p is None else ((abs(ar_p) - abs(toi_p)) if absolute else ar_p - toi_p)
        label = "; ".join(f"{h['label']} (p.{h['page']}{', note ' + h['note'] if h['note'] else ''})" for h in hits)[:200]
        if absolute:
            ar_p, toi_p = (abs(ar_p) if ar_p is not None else None), (abs(toi_p) if toi_p is not None else None)
        cmp_rows.append([STATEMENT_NAME[stmt], name, codes, a, b, diff, st, ar_p, toi_p, pri_diff, label])
        cmp_tags.append("total" if st == "Agrees" else "error" if st.startswith("Differs (over") else "warning")
        if st.startswith("Differs"):
            big.append((name, diff))

    # ---- opinion and keywords
    low = rep["low"]
    opinion = [(n, tag) for n, rx, tag in OPINION if re.search(rx, low)]
    if not opinion:
        opinion = [("Unmodified (clean) opinion", "total")] if re.search(r"present fairly|true and fair", low) else \
                  [("Opinion paragraph not found", "warning")]
    hits = []
    for name, rx in KEYWORDS:
        for page, text in rep["pages"]:
            for m in re.finditer(rx, text, re.I):
                s = max(0, m.start() - 160)
                hits.append([name, page, re.sub(r"\s+", " ", text[s:m.end() + 200]).strip()])
                break
    notes = []
    for page, text in rep["pages"]:
        for line in text.splitlines():
            m = re.match(r"^\s*(\d{1,2}(?:\.\d{1,2})?)\.?\s+([A-Z][A-Za-z ,&'()/\-]{3,70})\s*$", line)
            if m and not re.search(r"\d{3}", line) and m.group(2).strip().lower() not in ("page",):
                notes.append([m.group(1), m.group(2).strip(), page])
    seen, notes_u = set(), []
    for n in notes:
        k = (n[0], n[1].lower())
        if k not in seen:
            seen.add(k)
            notes_u.append(n)

    # ---- summary
    counts = {k: sum(1 for l in lines if l["stmt"] == k) for k in ("BS", "IS", "CF")}
    S = [["Report", f"{path.name} · {len(pages)} page(s) · year {rep['year'] or 'not found'}", None, None, "", "section"],
         ["Currency", f"{'Rate US$1 = KHR ' + format(rep['fx'], ',.0f') + (' (from the report)' if rep['fx_found'] else ' (default: not found in the report)')}",
          None, None, "US$ figures are converted to KHR with this rate when the report has no KHR column.", "even"]]
    for n, tag in opinion:
        S.append(["Audit opinion", n, None, None, "" if tag == "total" else "Read the basis paragraph; consider the impact on the TOI and tax risk.", tag])
    S.append(["Statements read", f"Balance sheet {counts['BS']} lines, income statement {counts['IS']}, cash flow {counts['CF']}",
              sum(counts.values()), None, "See 'Statement lines'." if sum(counts.values()) else
              "No statement table recognised: check the PDF has text tables, or use the Word version.", "even" if sum(counts.values()) else "error"])
    agree = sum(1 for r in cmp_rows if r[6] == "Agrees")
    S.append(["Compared with TOI", f"{agree} of {len(cmp_rows)} items agree", len(cmp_rows), None, "See 'Compare with TOI'.",
              "total" if cmp_rows and agree == len(cmp_rows) else "warning"])
    for name, diff in big[:10]:
        S.append(["   Difference: " + name, "Audited FS vs TOI", None, diff,
                  "The TOI should follow the audited accounts: ask for the reconciliation / tax adjustments.", "error"])
    for name in sorted({h[0] for h in hits}):
        S.append(["Note to read", name, None, None, "See 'Keyword hits'.", "warning" if name in ("Going concern", "Tax (uncertain / assessment / audit)",
                                                                                                "Contingent liabilities", "Prior-year restatement") else "even"])
    an.view("Summary", ["Area", "Result", "Count", "Amount (KHR)", "Note / action"], [r[:5] for r in S], [r[5] for r in S],
            money=[3], total=False, desc="Audited financial statements read from the Audit Report folder and compared with the TOI.")
    an.view("Compare with TOI", ["Statement", "Item", "TOI codes", "Audit report N (KHR)", "TOI N", "Difference N", "Status",
                                 "Audit report N-1 (KHR)", "TOI N-1", "Difference N-1", "Audit report line (page, note)"],
            cmp_rows, cmp_tags, money=[3, 4, 5, 7, 8, 9], total=False, wide=True,
            desc="Expenses and asset purchases are compared as positive amounts. Agrees = within 1 % or 1,000 KHR.")
    an.view("Statement lines", ["Statement", "Line", "Note", "N", "N-1", "Currency", "N (KHR)", "N-1 (KHR)", "Page"],
            [[STATEMENT_NAME[l["stmt"]], l["label"], l["note"], l["cur"], l["pri"], l["ccy"], l["cur_k"], l["pri_k"], l["page"]] for l in lines],
            None, money=[3, 4, 6, 7], total=False, desc="Every line read from the three statements.")
    an.view("Notes index", ["Note", "Title", "Page"], notes_u, None, total=False, desc="Note headings found in the report.")
    an.view("Keyword hits", ["Topic", "Page", "Text around the first mention"], hits,
            ["warning"] * len(hits), total=False, desc="Where the report mentions topics that matter for tax audit.")
    an.view("Full text", ["Page", "Text"], [[p, line] for p, text in rep["pages"] for line in text.splitlines() if line.strip()],
            None, total=False, desc="The report text, line by line: use Find or the heading filter to search it.")
    return an
