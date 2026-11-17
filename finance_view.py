"""Load GL / Sales / Bank (and other finance) workbooks into a table that can be
filtered, searched and sorted. No GUI code here, so it is fully testable.

    table = load_table(table_files(folder), kind="BANK")
    idx = table.view(find="globex", column="Description", month="2025-04", sort=(col, True))
    rows = [table.display(i) for i in idx[:200]]
"""
import csv
import re
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook

PREFERRED_SHEETS = ("transaction", "sale", "ledger", "detail", "data")
AMOUNT_WORDS = ("debit", "credit", "amount", "money in", "money out", "receipt", "payment", "sales (")
MONEY_FORMAT_WORDS = AMOUNT_WORDS + ("balance",)
NOT_AMOUNT = ("running balance", "no", "month", "page", "row", "code", "year", "count", "txn")
ALIASES = {
    "posting date": "Date", "date": "Date",
    "money in": "Money In", "money in / credit": "Money In", "money in (usd)": "Money In", "money in (khr)": "Money In",
    "bank receipts": "Money In", "credit": "Money In",
    "money out": "Money Out", "money out / debit": "Money Out", "money out (usd)": "Money Out",
    "money out (khr)": "Money Out", "bank payments": "Money Out", "debit": "Money Out",
    "bank description": "Description", "description / bank notes": "Description", "description": "Description",
    "running balance": "Running Balance", "running balance (usd)": "Running Balance",
    "currency": "Currency", "account": "Account",
    "category": "Bank category", "classification": "Bank category",
    "reference (ocr)": "Reference", "transaction id": "Reference", "txn code": "Reference", "cheque #": "Cheque #",
    "source statement pdf": "Source PDF", "source pdf": "Source PDF",
}
BANK_ORDER = ("Bank", "Date", "Account", "Currency", "Description", "Money In", "Money Out", "Running Balance", "Type",
              "Bank category", "Reference", "Cheque #", "Source PDF")
OPS = re.compile(r"^\s*(>=|<=|>|<|=)?\s*(-?[\d,]*\.?\d+)\s*(?:\.\.\s*(-?[\d,]*\.?\d+))?\s*$")


def table_files(folder):
    """Readable data files below a folder, ignoring hidden folders such as .pdf-extractions."""
    if folder is None:
        return []
    folder = Path(folder)
    return sorted(p for p in folder.rglob("*")
                  if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".csv")
                  and not any(part.startswith((".", "~$")) for part in p.relative_to(folder).parts))


def _clean(v):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    if isinstance(v, float) and v == int(v) and abs(v) < 1e15:
        return int(v)
    if isinstance(v, str):
        return v.strip()
    return v


def _is_blank(row):
    return all(v in (None, "") for v in row)


def _header_index(rows):
    """First row (within 25) with 3+ cells that are all text: that is the heading row."""
    for i, row in enumerate(rows[:25]):
        cells = [c for c in row if c not in (None, "")]
        if len(cells) >= 3 and all(isinstance(c, str) for c in cells):
            return i
    return None


def read_rows(path):
    """(columns, rows) of the main data sheet of one file."""
    path = Path(path)
    if path.suffix.lower() == ".csv":
        with open(path, encoding="utf-8-sig", errors="replace", newline="") as f:
            raw = [list(r) for r in csv.reader(f)]
    else:
        wb = load_workbook(path, read_only=True, data_only=True)
        ws = None
        for token in PREFERRED_SHEETS:
            ws = next((s for s in wb.worksheets if token in s.title.lower()), None)
            if ws:
                break
        ws = ws or wb.worksheets[0]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    h = _header_index(raw)
    if h is None:
        raise ValueError(f"{path.name}: no heading row found")
    width = max((i + 1 for i, c in enumerate(raw[h]) if c not in (None, "")), default=0)
    cols = [str(c).strip() if c not in (None, "") else f"Column {i + 1}" for i, c in enumerate(raw[h][:width])]
    rows = []
    for r in raw[h + 1:]:
        r = (list(r) + [None] * width)[:width]
        if not _is_blank(r):
            rows.append([_clean(v) for v in r])
    return cols, rows


def month_key(v):
    """'2025-10' for 2025-10-21, '21-10-2025', datetime; None otherwise."""
    if isinstance(v, (datetime, date)):
        return f"{v.year:04d}-{v.month:02d}"
    t = str(v or "").strip()
    m = re.match(r"(\d{4})-(\d{2})", t)
    if m:
        return f"{m[1]}-{m[2]}"
    m = re.match(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", t)
    if m:
        return f"{m[3]}-{int(m[2]):02d}"
    return None


class Table:
    def __init__(self, columns, rows, notes=None):
        self.columns, self.rows, self.notes = columns, rows, notes or []
        self.row_tags = None       # optional colour tag per row (analysis views)
        low = [c.lower() for c in columns]
        self.date_col = next((i for i, c in enumerate(columns) if c == "Date" or low[i] == "date"), None)
        if self.date_col is None:
            self.date_col = next((i for i, c in enumerate(low) if "date" in c), None)
        self.money = [i for i, c in enumerate(low) if any(w in c for w in MONEY_FORMAT_WORDS)]
        self.amount = [i for i in self.money if not any(re.search(rf"\b{w}\b", low[i]) for w in NOT_AMOUNT)]  # A/R "Balance" counts; "Running balance" does not
        self.currency_col = next((i for i, c in enumerate(columns) if c == "Currency"), None)
        self._text = None
        self.refresh_months()

    def refresh_months(self):
        self.months = sorted({m for r in self.rows for m in [month_key(r[self.date_col]) if self.date_col is not None else None] if m})

    # ---- display
    def display(self, i):
        out = []
        for c, v in enumerate(self.rows[i]):
            if v is None:
                out.append("")
            elif isinstance(v, (int, float)) and not isinstance(v, bool) and c in self.money:
                out.append(f"{v:,.2f}" if isinstance(v, float) and v != int(v) else f"{int(v):,}")
            else:
                out.append(str(v))
        return out

    # ---- filtering
    def _texts(self):
        if self._text is None:
            self._text = [[("" if v is None else str(v)).casefold() for v in r] for r in self.rows]
        return self._text

    def _numeric_test(self, query):
        m = OPS.match(query)
        if not m:
            return None
        op, a, b = m.group(1), float(m.group(2).replace(",", "")), m.group(3)
        if b is not None:
            lo, hi = sorted((a, float(b.replace(",", ""))))
            return lambda v: lo <= v <= hi
        return {"": lambda v: v == a, "=": lambda v: v == a, ">": lambda v: v > a, "<": lambda v: v < a,
                ">=": lambda v: v >= a, "<=": lambda v: v <= a}[op or ""]

    # ---- Excel-style heading filters: {column index: {"values": set of value keys or None, "text": str}}
    BLANK = "(Blanks)"

    def value_key(self, i, c):
        """Text a heading filter matches on: the value as shown in the table, (Blanks) for empty cells."""
        v = self.rows[i][c]
        if v in (None, ""):
            return self.BLANK
        if isinstance(v, (int, float)) and not isinstance(v, bool) and c in self.money:
            return f"{v:,.2f}" if isinstance(v, float) and v != int(v) else f"{int(v):,}"
        return str(v)

    def _filter_test(self, c, spec):
        values, text = spec.get("values"), (spec.get("text") or "").strip()
        tests = []
        if values is not None:
            values = set(values)
            tests.append(lambda i: self.value_key(i, c) in values)
        if text:
            num = self._numeric_test(text)
            op_given = bool(re.match(r"^\s*(>=|<=|>|<|=)", text) or ".." in text)
            if num and (op_given or c in self.money):
                tests.append(lambda i: isinstance(self.rows[i][c], (int, float)) and not isinstance(self.rows[i][c], bool)
                             and num(self.rows[i][c]))
            else:
                words = text.casefold().split()
                texts = self._texts()
                tests.append(lambda i: all(w in texts[i][c] for w in words))
        return lambda i: all(t(i) for t in tests)

    def apply_filters(self, idx, filters, skip=None):
        for c, spec in (filters or {}).items():
            if c == skip or c >= len(self.columns) or not spec:
                continue
            test = self._filter_test(c, spec)
            idx = [i for i in idx if test(i)]
        return list(idx)

    MAX_FILTER_VALUES = 5000

    def distinct(self, c, idx):
        """[(value key, row count)] of one column over rows idx, sorted like Excel (numbers, text, blanks).
        Returns (values, complete); complete is False when the column has too many different values."""
        counts = defaultdict(int)
        raw = {}
        for i in idx:
            k = self.value_key(i, c)
            counts[k] += 1
            raw.setdefault(k, self.rows[i][c])
        def order(k):
            v = raw[k]
            if k == self.BLANK:
                return (3, 0, "")
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                return (0, v, "")
            return (1, 0, str(v).casefold())
        keys = sorted(counts, key=order)
        complete = len(keys) <= self.MAX_FILTER_VALUES
        return [(k, counts[k]) for k in keys[:self.MAX_FILTER_VALUES]], complete

    def filter_text(self, filters):
        """Short description of the active heading filters, for the totals bar."""
        parts = []
        for c, spec in sorted((filters or {}).items()):
            if not spec or c >= len(self.columns):
                continue
            bits = []
            if spec.get("values") is not None:
                vals = sorted(spec["values"])
                bits.append(" / ".join(vals[:3]) + (f" +{len(vals) - 3} more" if len(vals) > 3 else ""))
            if spec.get("text"):
                bits.append(f"contains '{spec['text']}'" if not self._numeric_test(spec["text"]) else spec["text"])
            parts.append(f"{self.columns[c]}: " + ", ".join(bits))
        return "; ".join(parts)

    def view(self, find="", column=None, month=None, sort=None, field=None, value=None, filters=None):
        """Row indexes after filtering and sorting. column: heading text or None for all columns.
        find: words (all must appear) or, for number columns, >1000000  <=500  100..200  1,234.
        filters: Excel-style heading filters, see apply_filters."""
        idx = range(len(self.rows))
        if filters:
            idx = self.apply_filters(idx, filters)
        col = self.columns.index(column) if column in self.columns else None
        if field in self.columns and value not in (None, ""):
            fc = self.columns.index(field)
            idx = [i for i in idx if ("" if self.rows[i][fc] is None else str(self.rows[i][fc])) == str(value)]
        if month and self.date_col is not None:
            idx = [i for i in idx if month_key(self.rows[i][self.date_col]) == month]
        find = (find or "").strip()
        if find:
            test = self._numeric_test(find)
            op_given = bool(re.match(r"^\s*(>=|<=|>|<|=)", find) or ".." in find)
            numeric_cols = [col] if col is not None else self.amount
            if test and (op_given or (col is not None and col in self.money)):
                def hit(i):
                    for c in numeric_cols:
                        v = self.rows[i][c]
                        if isinstance(v, (int, float)) and not isinstance(v, bool) and test(v):
                            return True
                    return False
                idx = [i for i in idx if hit(i)]
            else:
                words = find.casefold().split()
                texts = self._texts()
                if col is None:
                    idx = [i for i in idx if all(any(w in t for t in texts[i]) for w in words)]
                else:
                    idx = [i for i in idx if all(w in texts[i][col] for w in words)]
        idx = list(idx)
        if sort is not None and sort[0] is not None:
            c, desc = sort
            filled = [i for i in idx if self.rows[i][c] not in (None, "")]
            blank = [i for i in idx if self.rows[i][c] in (None, "")]
            nums = sorted([i for i in filled if not isinstance(self.rows[i][c], str)], key=lambda i: self.rows[i][c], reverse=desc)
            strs = sorted([i for i in filled if isinstance(self.rows[i][c], str)], key=lambda i: self.rows[i][c].casefold(), reverse=desc)
            idx = (strs + nums if desc else nums + strs) + blank
        return idx

    MAX_CHOICES = 1500

    def choices(self, column):
        """[(label shown, value to match)] for the value list of one column; [] if there are too many."""
        if column not in self.columns:
            return []
        c = self.columns.index(column)
        values = {str(r[c]) for r in self.rows if r[c] not in (None, "")}
        try:
            values = sorted(values, key=float)
        except ValueError:
            values = sorted(values, key=str.casefold)
        if len(values) > self.MAX_CHOICES:
            return []
        if column == "Account Code" and "Account Name" in self.columns:
            n = self.columns.index("Account Name")
            names = {}
            for r in self.rows:
                names.setdefault(str(r[c]), r[n])
            return [(f"{v} · {names.get(v) or ''}".strip(" ·"), v) for v in values]
        return [(v, v) for v in values]

    def default_field(self):
        for name in ("Account", "Account Name", "Buyer Name", "Customer", "Name / Counterparty", "Name"):
            if name in self.columns:
                return name
        return None

    def totals(self, idx):
        """[(label, value)] for the amount columns; split by currency when the table has one."""
        out = {}
        for i in idx:
            cur = self.rows[i][self.currency_col] if self.currency_col is not None else None
            for c in self.amount:
                v = self.rows[i][c]
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    label = f"{self.columns[c]}" + (f" {cur}" if cur else "")
                    out[label] = out.get(label, 0) + v
        return [(k, v) for k, v in out.items() if v]

    @staticmethod
    def money_text(v):
        return f"{v:,.0f}" if abs(v - round(v)) < 0.005 else f"{v:,.2f}"

    def summary(self, idx, compact=False):
        """Text for the totals bar: row count, each amount column total, and net when both sides exist.
        Bank and other tables with a Currency column are totalled per currency."""
        idx = list(idx)
        parts = [f"{len(idx):,} rows"]
        groups = defaultdict(lambda: defaultdict(float))
        for i in idx:
            cur = (self.rows[i][self.currency_col] if self.currency_col is not None else None) or ""
            for c in self.amount:
                v = self.rows[i][c]
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    groups[cur][self.columns[c]] += v
        for cur, cols in sorted(groups.items()):
            tag = f" {cur}" if cur else ""
            cols = dict(sorted(cols.items(), key=lambda kv: self.columns.index(kv[0]) if kv[0] in self.columns else 0))
            shown = [(n, v) for n, v in cols.items() if v] if compact else list(cols.items())
            items = [f"{name}{tag}: {self.money_text(v)}" for name, v in shown[:4 if compact else None]]
            if compact and len(shown) > 4:
                items.append(f"+{len(shown) - 4} more amount columns")
            pairs = (("Debit", "Credit", "Debit − Credit"), ("Money In", "Money Out", "In − Out"))
            for a, b, label in pairs:
                if a in cols and b in cols:
                    items.append(f"{label}{tag}: {self.money_text(cols[a] - cols[b])}")
            parts.extend(items)
        if len(parts) == 1 and not self.amount:
            parts.append("no amount column in this file")
        return "   |   ".join(parts)

    def export(self, path, idx):
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        ws = wb.active
        ws.title = "Filtered data"
        ws.append(self.columns)
        for c in ws[1]:
            c.font, c.fill = Font(bold=True), PatternFill("solid", fgColor="BDD7EE")
        n = 0
        for n, i in enumerate(idx, start=1):
            ws.append(self.rows[i])
        for c in self.money:
            for (cell,) in ws.iter_rows(min_row=2, min_col=c + 1, max_col=c + 1):
                cell.number_format = "#,##0.00;(#,##0.00);-"
        if self.columns:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(self.columns))}{max(1, n + 1)}"  # Excel filter buttons on the headings
            ws.freeze_panes = "A2"
            for c, name in enumerate(self.columns, start=1):
                ws.column_dimensions[get_column_letter(c)].width = max(10, min(45, len(str(name)) + 4))
        wb.save(path)


def _canonical(name, kind=""):
    """Bank files from different banks are aligned by canonical headings; other tables keep theirs."""
    return ALIASES.get(name.strip().lower(), name.strip()) if kind == "BANK" else name.strip()


def load_table(files, kind=""):
    """Combine one or more files into one Table. Bank files with different headings are aligned
    by canonical names (Date, Description, Money In, Money Out ...) and get Bank and Type columns."""
    notes, parts = [], []
    for p in files:
        try:
            cols, rows = read_rows(p)
        except Exception as exc:
            notes.append(f"{Path(p).name}: {str(exc).splitlines()[0]}")
            continue
        parts.append((Path(p), cols, rows))
    return table_from_parts(parts, kind, notes)


def table_from_parts(parts, kind="", notes=None):
    """parts: [(path, columns, rows)] already read with read_rows."""
    kind = (kind or "").upper()
    notes = list(notes or [])
    if not parts:
        return Table([], [], notes)
    multi = len(parts) > 1
    if kind == "BANK":
        from sales_reconciliation import classify_credit
    columns, rows_out = [], []
    for p, cols, rows in parts:
        names = [_canonical(c, kind) for c in cols]
        # (USD)/(KHR) in a heading is the currency when there is no Currency column
        cur = next((m.group(1).upper() for c in cols for m in [re.search(r"\((usd|khr)\)", c, re.I)] if m), None)
        extra = []
        if kind == "BANK":
            extra = ["Bank", "Account"] + (["Currency"] if "Currency" not in names and cur else []) + ["Type"]
        elif multi:
            extra = ["Source file"]
        for n in extra + names:
            if n not in columns:
                columns.append(n)
        for r in rows:
            rec = dict(zip(names, r))
            if kind == "BANK":
                rec["Bank"] = re.split(r"[_ ]", p.stem)[0]
                if "Currency" not in rec and cur:
                    rec["Currency"] = cur
                rec["Account"] = re.sub(r"\s+", " ", f"{rec['Bank']} {rec.get('Account') or ''} ({rec.get('Currency') or ''})").strip()
                cin = rec.get("Money In")
                if isinstance(cin, (int, float)) and cin > 0:
                    rec["Type"] = classify_credit(str(rec.get("Description") or ""), str(rec.get("Currency") or "USD").upper())
            elif multi:
                rec["Source file"] = p.name
            rows_out.append(rec)
    if kind == "BANK":
        columns = [c for c in BANK_ORDER if c in columns]  # compact, aligned view of every bank
    table_rows = [[rec.get(c) for c in columns] for rec in rows_out]
    if kind == "BANK":
        d = columns.index("Date") if "Date" in columns else None
        if d is not None:
            table_rows.sort(key=lambda r: (str(r[d] or ""),))
    return Table(columns, table_rows, notes)
