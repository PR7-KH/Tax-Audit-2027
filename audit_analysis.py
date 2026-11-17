"""Shared pieces of the Audit analysis menus (Revenue, Expense, TOS).

An Analysis is a list of named views; each view is a finance_view.Table, so the Control Center
shows it with the same Excel-style heading filters, sort, Find and SUM bar as GL / Sales / Bank.
save_workbook() writes every view to one formatted Excel file with AutoFilter buttons on the
headings and a SUBTOTAL row that follows the Excel filter. Nothing here touches the GUI.
"""
import re
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path

import finance_view as fv

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_WORDS = {w: i + 1 for i, w in enumerate(m.lower() for m in MONTHS)}
MONTH_WORDS.update({"january": 1, "february": 2, "march": 3, "april": 4, "june": 6, "july": 7, "august": 8,
                    "september": 9, "sept": 9, "october": 10, "november": 11, "december": 12})
FILLS = {"error": "F8CBAD", "warning": "FFE699", "total": "C6E0B4", "section": "DCE9F3"}


# ------------------------------------------------------------------ small helpers
def num(v):
    if isinstance(v, bool) or v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace(",", "").replace("៛", "").replace("$", "").strip()
    neg = t.startswith("(") and t.endswith(")")
    try:
        x = float(t.strip("()"))
    except ValueError:
        return 0.0
    return -x if neg else x


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and 20000 < v < 80000:   # Excel serial date
        from openpyxl.utils.datetime import from_excel
        d = from_excel(v)
        return d.date() if isinstance(d, datetime) else d
    t = str(v or "").strip()
    m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        return date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})", t)
    if m:
        try:
            return date(int(m[3]), int(m[2]), int(m[1]))
        except ValueError:
            return None
    return None


def file_month(name, year):
    """Month number from a monthly file name: TOS_03_2025, 2025-03, TOS Mar 2025, Salary_March_2025."""
    stem = Path(name).stem
    year = str(year)
    for pat in (rf"(?<!\d)(0?[1-9]|1[0-2])[ _\-.]+{year}(?!\d)", rf"(?<!\d){year}[ _\-.]+(0?[1-9]|1[0-2])(?!\d)",
                rf"(?<!\d)(0[1-9]|1[0-2]){year}(?!\d)", rf"(?<!\d){year}(0[1-9]|1[0-2])(?!\d)"):
        m = re.search(pat, stem)
        if m:
            return int(m.group(1))
    words = re.findall(r"[A-Za-z]+", stem.lower())
    for w in words:
        if w in MONTH_WORDS:
            return MONTH_WORDS[w]
    m = re.search(r"(?<!\d)(0[1-9]|1[0-2])(?!\d)", stem.replace(year, " "))
    return int(m.group(1)) if m else None


def month_key(ym):
    return f"{ym[0]:04d}-{ym[1]:02d}"


def month_label(ym):
    return f"{MONTHS[ym[1] - 1]} {ym[0]}"


def parse_label(text):
    m = re.match(r"([A-Za-z]{3}) (\d{4})", str(text or ""))
    return (int(m[2]), MONTH_WORDS[m[1].lower()]) if m else (0, 0)


def next_month(ym):
    return (ym[0] + 1, 1) if ym[1] == 12 else (ym[0], ym[1] + 1)


def prev_month(ym):
    return (ym[0] - 1, 12) if ym[1] == 1 else (ym[0], ym[1] - 1)


def month_range(a, b):
    out = []
    while a <= b:
        out.append(a)
        a = next_month(a)
    return out


def month_list(months):
    """'Mar 2025, Jun–Aug 2025': consecutive months are joined."""
    months = sorted(set(months))
    parts, i = [], 0
    while i < len(months):
        j = i
        while j + 1 < len(months) and months[j + 1] == next_month(months[j]):
            j += 1
        a, b = months[i], months[j]
        if a == b:
            parts.append(month_label(a))
        elif a[0] == b[0]:
            parts.append(f"{MONTHS[a[1] - 1]}–{MONTHS[b[1] - 1]} {a[0]}")
        else:
            parts.append(f"{month_label(a)} – {month_label(b)}")
        i = j + 1
    return ", ".join(parts)


def col_letter(i):
    from openpyxl.utils import get_column_letter
    return get_column_letter(i + 1)


def subfolder(root, names):
    root = Path(root)
    if not root.is_dir():
        return None
    for d in sorted(root.iterdir()):
        if d.is_dir() and d.name.upper() in names and not d.name.startswith("."):
            return d
    return None


def two_row_labels(headers, merges):
    """One label per column from the two heading rows (rows 2 and 3) of a monthly tax return,
    filling merged heading cells. merges: (top, left, bottom, right) 0-based, rows 1..2."""
    width = max((len(h) for h in headers), default=0)
    grid = [list(h) + [None] * (width - len(h)) for h in headers]
    for top, left, bottom, right in merges:
        value = grid[top - 1][left] if 0 <= top - 1 < len(grid) and left < width else None
        for r in range(top, bottom + 1):
            for c in range(left, min(right + 1, width)):
                if 0 <= r - 1 < len(grid) and grid[r - 1][c] in (None, ""):
                    grid[r - 1][c] = value
    labels = []
    for c in range(width):
        a = re.sub(r"\s+", " ", str(grid[0][c])).strip() if grid and grid[0][c] not in (None, "") else ""
        b = re.sub(r"\s+", " ", str(grid[1][c])).strip() if len(grid) > 1 and grid[1][c] not in (None, "") else ""
        labels.append(f"{a} / {b}" if a and b and a != b else a or b)
    return labels


# ------------------------------------------------------------------ the analysis object
class Analysis:
    def __init__(self, key, title):
        self.key, self.title = key, title
        self.views = OrderedDict()
        self.notes = []
        self.period = ""

    def add_note(self, text):
        self.notes.append(text)

    def view(self, name, columns, rows, tags=None, money=None, month_col=None, desc="", total=True, wide=False):
        """total=False: amounts are formatted but not added up (summary lines). wide=True: show actual widths."""
        table = fv.Table(list(columns), [list(r) for r in rows], [])
        table.money = list(money or [])
        table.amount = list(table.money) if total else []
        table.wide = wide
        table.date_col = month_col
        table.refresh_months()
        table.row_tags = list(tags) if tags else None
        table.description = desc
        self.views[name] = table
        return table

    def add_summary_only(self):
        rows = [["Not ready", n, None, ""] for n in self.notes] or [["Not ready", "No data", None, ""]]
        self.view("Summary", ["Area", "Result", "Count", "Suggested audit action"], rows, ["error"] * len(rows))

    def headline(self):
        flagged = sum(1 for t in self.views.values() for tag in (t.row_tags or []) if tag == "error")
        return (f"{self.title}{' · ' + self.period if self.period else ''}: {len(self.views)} views, "
                f"{flagged:,} high-risk line(s)")


def save_workbook(analysis, path, company=""):
    """Every view on its own sheet: title, description, SUBTOTAL row, headings with AutoFilter."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter as L

    wb = Workbook()
    wb.remove(wb.active)
    thin = Side(style="thin", color="BFBFBF")
    used = set()
    for name, t in analysis.views.items():
        title = re.sub(r"[\\/:?*\[\]]", "-", name)[:31]
        while title.casefold() in used:
            title = title[:28] + f"_{len(used)}"
        used.add(title.casefold())
        ws = wb.create_sheet(title)
        ncol = max(1, len(t.columns))
        ws.cell(1, 1, f"{analysis.title} – {name}").font = Font(bold=True, size=14, color="101F35")
        ws.cell(2, 1, " · ".join(x for x in (company, analysis.period, getattr(t, "description", "")) if x)).font = Font(italic=True, color="61748B")
        head = 4
        if t.amount:
            ws.cell(3, 1, "Total of visible rows").font = Font(bold=True)
        for c, col in enumerate(t.columns, start=1):
            cell = ws.cell(head, c, col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
        last = head
        for i, row in enumerate(t.rows):
            r = head + 1 + i
            last = r
            for c, v in enumerate(row, start=1):
                cell = ws.cell(r, c, v)
                if (c - 1) in t.money:
                    cell.number_format = "#,##0;(#,##0);-"
            tag = t.row_tags[i] if t.row_tags and i < len(t.row_tags) else None
            if tag in FILLS:
                fill = PatternFill("solid", fgColor=FILLS[tag])
                for c in range(1, ncol + 1):
                    ws.cell(r, c).fill = fill
        for c in t.amount:
            letter = L(c + 1)
            cell = ws.cell(3, c + 1, f"=SUBTOTAL(109,{letter}{head + 1}:{letter}{max(last, head + 1)})")
            cell.number_format = "#,##0;(#,##0);-"
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="E3F2EC")
        ws.auto_filter.ref = f"A{head}:{L(ncol)}{max(last, head)}"
        ws.freeze_panes = ws.cell(head + 1, min(3, ncol))
        for c, col in enumerate(t.columns, start=1):
            longest = max([len(str(col))] + [len(str(r[c - 1])) for r in t.rows[:300] if c - 1 < len(r) and r[c - 1] is not None])
            ws.column_dimensions[L(c)].width = max(9, min(60, longest + 2))
        ws.row_dimensions[head].height = 32
    if analysis.notes:
        ws = wb.create_sheet("Notes")
        ws.append(["Notes and skipped files"])
        ws["A1"].font = Font(bold=True)
        for n in analysis.notes:
            ws.append([n])
        ws.column_dimensions["A"].width = 120
    wb.save(path)
    return path
