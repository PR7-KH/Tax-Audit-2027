"""Small Excel formula calculator so the Control Center can show reports without Excel.

Supports what the engine writes: IF, AND, OR, SUM, N, TODAY, cell and range references
(including other sheets), + - * / & comparisons and percent. Empty cells behave like Excel:
equal to both "" and 0, and a formula that returns an empty reference shows 0.
"""
import re
from datetime import date

TOKEN = re.compile(r"""
    \s*(?:
      (?P<str>"(?:[^"]|"")*")
     |(?P<num>\d+(?:\.\d+)?)
     |(?P<ref>(?:(?:'[^']+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)
     |(?P<fn>[A-Z][A-Z0-9\.]*)\(
     |(?P<op><>|<=|>=|[-+*/&=<>^%(),])
    )""", re.X)


class XlError(Exception):
    def __init__(self, code):
        super().__init__(code)
        self.code = code


def tokenize(text):
    out, pos = [], 0
    while pos < len(text):
        m = TOKEN.match(text, pos)
        if not m or m.end() == pos:
            if text[pos:].strip() == "":
                break
            raise XlError("#NAME?")
        pos = m.end()
        kind = m.lastgroup
        out.append((kind, m.group(kind)))
    return out


def col_num(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


class Parser:
    def __init__(self, tokens):
        self.t, self.i = tokens, 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None)

    def take(self):
        tok = self.peek()
        self.i += 1
        return tok

    def expr(self, level=0):
        levels = (("=", "<>", "<", ">", "<=", ">="), ("&",), ("+", "-"), ("*", "/"), ("^",))
        if level == len(levels):
            return self.unary()
        left = self.expr(level + 1)
        while self.peek()[0] == "op" and self.peek()[1] in levels[level]:
            op = self.take()[1]
            left = ("bin", op, left, self.expr(level + 1))
        return left

    def unary(self):
        if self.peek() == ("op", "-"):
            self.take()
            return ("neg", self.unary())
        if self.peek() == ("op", "+"):
            self.take()
            return self.unary()
        node = self.atom()
        while self.peek() == ("op", "%"):
            self.take()
            node = ("bin", "/", node, ("num", 100.0))
        return node

    def atom(self):
        kind, val = self.take()
        if kind == "num":
            return ("num", float(val))
        if kind == "str":
            return ("str", val[1:-1].replace('""', '"'))
        if kind == "ref":
            return ("ref", val)
        if kind == "fn":
            args = []
            if self.peek() == ("op", ")"):
                self.take()
                return ("fn", val, args)
            while True:
                args.append(self.expr())
                k, v = self.take()
                if v == ")":
                    break
                if v != ",":
                    raise XlError("#NAME?")
            return ("fn", val, args)
        if (kind, val) == ("op", "("):
            node = self.expr()
            self.take()
            return node
        raise XlError("#NAME?")


class Calc:
    def __init__(self, wb):
        self.wb, self.cache, self.stack = wb, {}, set()

    def value(self, sheet, row, col):
        """Computed value of a cell (None when empty)."""
        key = (sheet, row, col)
        if key in self.cache:
            return self.cache[key]
        if key in self.stack:
            raise XlError("#REF!")
        v = self.wb[sheet].cell(row, col).value
        if isinstance(v, str) and v.startswith("="):
            self.stack.add(key)
            try:
                r = self.run(v[1:], sheet)
                if isinstance(r, list):
                    r = r[0][0] if r and r[0] else None
                if r is None:
                    r = 0
            except XlError as e:
                r = e.code
            except ZeroDivisionError:
                r = "#DIV/0!"
            finally:
                self.stack.discard(key)
            v = r
        self.cache[key] = v
        return v

    def run(self, formula, sheet):
        return self.ev(Parser(tokenize(formula)).expr(), sheet)

    def ref(self, text, sheet):
        sh = sheet
        if "!" in text:
            sh, text = text.rsplit("!", 1)
            sh = sh.strip("'")
        if sh not in self.wb.sheetnames:
            raise XlError("#REF!")
        text = text.replace("$", "")
        parts = text.split(":")
        cells = []
        for p in parts:
            m = re.fullmatch(r"([A-Z]+)(\d+)", p)
            cells.append((int(m.group(2)), col_num(m.group(1))))
        if len(cells) == 1:
            return self.value(sh, *cells[0])
        (r1, c1), (r2, c2) = cells
        return [[self.value(sh, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]

    @staticmethod
    def num(v):
        if v is None or v == "":
            return 0.0
        if isinstance(v, bool):
            return float(v)
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, str) and v.startswith("#"):
            raise XlError(v)
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            raise XlError("#VALUE!")

    @staticmethod
    def text(v):
        if v is None:
            return ""
        if isinstance(v, float) and v == int(v):
            return str(int(v))
        return str(v)

    @staticmethod
    def truth(v):
        if isinstance(v, str):
            if v.startswith("#"):
                raise XlError(v)
            raise XlError("#VALUE!")
        return bool(v) if v is not None else False

    def compare(self, op, a, b):
        for x in (a, b):
            if isinstance(x, str) and x.startswith("#"):
                raise XlError(x)
        if a is None:
            a = "" if isinstance(b, str) else 0
        if b is None:
            b = "" if isinstance(a, str) else 0
        if isinstance(a, str) != isinstance(b, str):
            # Excel: text is always greater than numbers
            ka, kb = (1 if isinstance(a, str) else 0), (1 if isinstance(b, str) else 0)
            a, b = ka, kb
        elif isinstance(a, str):
            a, b = a.lower(), b.lower()
        return {"=": a == b, "<>": a != b, "<": a < b, ">": a > b, "<=": a <= b, ">=": a >= b}[op]

    def ev(self, node, sheet):
        k = node[0]
        if k == "num" or k == "str":
            return node[1]
        if k == "ref":
            return self.ref(node[1], sheet)
        if k == "neg":
            return -self.num(self.ev(node[1], sheet))
        if k == "bin":
            op = node[1]
            a, b = self.ev(node[2], sheet), self.ev(node[3], sheet)
            if op in ("=", "<>", "<", ">", "<=", ">="):
                return self.compare(op, a, b)
            if op == "&":
                return self.text(a) + self.text(b)
            x, y = self.num(a), self.num(b)
            if op == "+":
                return x + y
            if op == "-":
                return x - y
            if op == "*":
                return x * y
            if op == "/":
                if y == 0:
                    raise XlError("#DIV/0!")
                return x / y
            return x ** y
        if k == "fn":
            return self.fn(node[1], node[2], sheet)
        raise XlError("#NAME?")

    def fn(self, name, args, sheet):
        if name == "IF":
            cond = self.truth(self.ev(args[0], sheet))
            branch = args[1] if cond else (args[2] if len(args) > 2 else None)
            return self.ev(branch, sheet) if branch is not None else False
        if name == "AND":
            return all(self.truth(self.scalar(self.ev(a, sheet))) for a in args)
        if name == "OR":
            return any(self.truth(self.scalar(self.ev(a, sheet))) for a in args)
        if name == "SUM":
            total = 0.0
            for a in args:
                v = self.ev(a, sheet)
                cells = [c for row in v for c in row] if isinstance(v, list) else [v]
                for c in cells:
                    if isinstance(c, str) and c.startswith("#"):
                        raise XlError(c)
                    if isinstance(c, (int, float)) and not isinstance(c, bool):
                        total += c
            return total
        if name == "N":
            v = self.scalar(self.ev(args[0], sheet))
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
        if name == "TODAY":
            return (date.today() - date(1899, 12, 30)).days
        if name == "IFERROR":
            try:
                v = self.ev(args[0], sheet)
                if isinstance(v, str) and v.startswith("#"):
                    raise XlError(v)
                return v
            except (XlError, ZeroDivisionError):
                return self.ev(args[1], sheet)
        raise XlError("#NAME?")

    @staticmethod
    def scalar(v):
        return v[0][0] if isinstance(v, list) else v


def display(value, number_format="General"):
    """Text as Excel would roughly show it."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        f = number_format or "General"
        if "yy" in f.lower():
            return (date(1899, 12, 30).fromordinal(date(1899, 12, 30).toordinal() + int(value))).isoformat()
        if "%" in f:
            return f"{value * 100:,.{1 if '0.0' in f else 0}f}%"
        if ".00" in f:
            return f"{value:,.2f}"
        if "#,##0" in f:
            return f"{value:,.0f}"
        if value == int(value):
            return str(int(value))
        return f"{value:,.4f}".rstrip("0").rstrip(".")
    return str(value)


def sheet_table(wb, name, calc=None, max_rows=5000):
    """All cells of a sheet as display text: (list of rows, number of columns)."""
    calc = calc or Calc(wb)
    ws = wb[name]
    rows = []
    for r in range(1, min(ws.max_row, max_rows) + 1):
        line = []
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            line.append(display(calc.value(name, r, c), cell.number_format))
        rows.append(line)
    while rows and not any(rows[-1]):
        rows.pop()
    return rows, ws.max_column
