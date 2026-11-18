"""Make an anonymous sample company folder from real files, safe to share or put in a public repository.

    python3 scripts/make_sample.py "<real company folder>" "<output folder>" [--month 01] [--year 2025]

For each target folder (PR, SR, TOS, WHT, TOI, GL, Info ...) one file is copied, with the same layout,
headings, type texts (បុគ្គលមិនជាប់អាករ ...) and amounts, but:
  * the company name and TIN are replaced by "Company XYZ" / L001-000000000 (also in file names);
  * supplier, buyer, payee, employee, shareholder and counterparty names become Supplier 001, Employee 0001 ...;
  * TIN, ID-card, passport and personal tax numbers become fake numbers of the same format;
  * government payees (customs offices, NSSF, GDT, EDC ...) are kept, so the WHT rules still work;
  * the GL is cut to one month and its Memo / Name columns are anonymised.
Review the output before sharing it: free-text descriptions are kept as they are.
"""
import argparse
import hashlib
import re
import shutil
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook

GOV = re.compile(r"គយ|ប្រតិវេទ|រដ្ឋាករ|សន្តិសុខសង្គម|\bNSSF\b|ពន្ធដារ|ក្រសួង|រតនាគារ|ministry|customs|treasury|\bGDT\b|"
                 r"អគ្គិសនីកម្ពុជា|\bEDC\b|ppwsa|ប្រាក់បំណាច់|ប្រាក់អតីតភាព|ប្រាក់អ្នកលាឈប់|^[A-Z]{3}\d{2}$|^[A-Z]+$", re.I)
NAME_HEAD = re.compile(r"ឈ្មោះ|name|counterparty|customer|supplier|vendor|payee|buyer", re.I)
ID_HEAD = re.compile(r"លេខសម្គាល់|អត្តសញ្ញាណ|លិខិតឆ្លងដែន|passport|\btin\b|id-card|\bid\b", re.I)
SAFE_MEMO = re.compile(r"^(cash withdrawal|bank service charge|bank charge|salary|wage|payroll|interest|transfer|deposit|"
                       r"opening balance|depreciation|accrual|nssf|tax|vat|exchange|fx|reversal)[\s\w\-/%.]*$", re.I)


class Fake:
    def __init__(self, company_names, company_tin):
        self.maps = {}
        self.company_names = [n for n in company_names if n]
        self.company_tin = company_tin

    def name(self, value, kind):
        v = str(value).strip()
        if not v or GOV.search(v):
            return value
        m = self.maps.setdefault(kind, {})
        if v not in m:
            m[v] = f"{kind} {len(m) + 1:0{4 if kind == 'Employee' else 3}d}"
        return m[v]

    @staticmethod
    def digits(value):
        """Same format, different digits (stable: the same real number always gets the same fake one)."""
        v = str(value)
        if GOV.search(v.strip()):
            return value
        h = hashlib.sha256(v.encode()).hexdigest()
        if not re.search(r"\d", v):          # a name typed in the ID column: replace it by a code
            return "ID-" + h[:6].upper() if v.strip() else value
        it = iter(int(c, 16) % 10 for c in h * 4)
        out = re.sub(r"\d", lambda m: str(next(it)), v)
        return int(out) if isinstance(value, int) else out

    def text(self, value):
        """Company name / TIN inside titles and file names."""
        if not isinstance(value, str):
            return value
        out = value
        if self.company_tin:
            out = out.replace(self.company_tin, "L001-000000000").replace(self.company_tin.split("-")[-1], "000000000")
        for n in self.company_names:
            out = re.sub(re.escape(n), "Company XYZ", out, flags=re.I)
        return out


def header_labels(ws, rows=(2, 3)):
    width = ws.max_column
    labels = []
    for c in range(1, width + 1):
        parts = [str(ws.cell(r, c).value or "") for r in rows]
        labels.append(" / ".join(p for p in parts if p))
    # merged heading cells: the label sits in the first column of the merge
    for m in ws.merged_cells.ranges:
        if m.min_row in rows:
            top = str(ws.cell(m.min_row, m.min_col).value or "")
            for c in range(m.min_col, m.max_col + 1):
                if top and top not in labels[c - 1]:
                    labels[c - 1] = top + " / " + labels[c - 1]
    return labels


def anonymise_return(path, out, fake, kind):
    wb = load_workbook(path)
    ws = wb.worksheets[0]
    ws.title = fake.text(ws.title)
    ws["A1"] = fake.text(ws["A1"].value)
    labels = header_labels(ws)
    party = {"PR": "Supplier", "SR": "Buyer", "WHT": "Payee", "TOS": "Employee"}[kind]
    name_cols = [i + 1 for i, l in enumerate(labels) if NAME_HEAD.search(l)]
    id_cols = [i + 1 for i, l in enumerate(labels) if ID_HEAD.search(l) and i + 1 not in name_cols]
    for row in ws.iter_rows(min_row=4):
        for c in row:
            if c.value in (None, ""):
                continue
            if c.column in name_cols and isinstance(c.value, str):
                c.value = fake.name(c.value, party)
            elif c.column in id_cols:
                c.value = fake.digits(c.value)
    wb.save(out)
    return dict(names=len(name_cols), ids=len(id_cols))


def anonymise_toi(path, out, fake):
    wb = load_workbook(path)
    for ws in wb.worksheets:
        step = ws.title.lower().replace(" ", "")
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str):
                    continue
                if step == "step1" and "៖" in v:
                    label, _, value = v.partition("៖")
                    if re.search(r"ឈ្មោះ|អាសយដ្ឋាន|លេខអត្តសញ្ញាណកម្ម", label):
                        c.value = label + "៖ " + ("L001-000000000" if "លេខអត្តសញ្ញាណកម្ម" in label else "XXX")
                        continue
                if step in ("step2", "step14") and c.row > 9 and c.column in (2, 3) and not v.startswith(("ក.", "ខ.", "គ.", "ឃ.", "ង.", "សរុប", "ល.រ")) \
                        and not re.search(r"ឈ្មោះ|ប្រទេស|អាសយដ្ឋាន|បរិយាយ", v):
                    c.value = fake.name(v, "Shareholder" if step == "step2" else "Related party") if c.column == 2 else "Address XXX"
                    continue
                c.value = fake.text(v)
    wb.save(out)


def anonymise_gl(path, out, fake, month):
    src = load_workbook(path, read_only=True, data_only=True)
    wb = Workbook()
    wb.remove(wb.active)
    for ws in src.worksheets:
        if re.search(r"about|check", ws.title, re.I):
            continue
        dst = wb.create_sheet(ws.title)
        rows = ws.iter_rows(values_only=True)
        head = next(rows, None)
        if head is None:
            continue
        dst.append(list(head))
        low = [str(h or "").lower() for h in head]
        date_col = next((i for i, h in enumerate(low) if h == "date" or h == "period"), None)
        name_cols = [i for i, h in enumerate(low) if re.search(r"name / counterparty|counterparty|^name$|customer|supplier|vendor", h)]
        memo_cols = [i for i, h in enumerate(low) if h in ("memo", "description", "narration")]
        for r in rows:
            r = list(r)
            if date_col is not None and "transaction" in ws.title.lower():
                d = r[date_col]
                key = f"{d.year}-{d.month:02d}" if hasattr(d, "year") else str(d or "")[:7]
                if not key.endswith(f"-{month}"):
                    continue
            for i in name_cols:
                if isinstance(r[i], str) and r[i].strip():
                    r[i] = fake.name(r[i], "Party")
            for i in memo_cols:
                if isinstance(r[i], str) and r[i].strip() and not SAFE_MEMO.match(r[i].strip()):
                    r[i] = fake.name(r[i], "Memo")
            dst.append([fake.text(v) for v in r])
    wb.save(out)


def anonymise_info(path, out, fake):
    wb = load_workbook(path)
    s1 = wb.worksheets[0]
    for row in s1.iter_rows():
        if row[0].value and len(row) > 1 and isinstance(row[1].value, str):
            label = str(row[0].value)
            if re.search(r"ឈ្មោះ|name|អាសយដ្ឋាន|address|លេខអត្តសញ្ញាណកម្ម|tin", label, re.I):
                row[1].value = "L001-000000000" if re.search(r"លេខអត្តសញ្ញាណកម្ម|tin", label, re.I) else "Company XYZ" if re.search(r"ឈ្មោះ|name", label, re.I) else "XXX"
    wb.save(out)


def pick(files, month, year):
    files = sorted(f for f in files if f.suffix.lower() in (".xlsx", ".xlsm") and not f.name.startswith((".", "~$")))
    if not files:
        return None
    best = [f for f in files if re.search(rf"(?<!\d){month}[_\-. ]+{year}", f.stem) and "(" not in f.stem]
    return (best or files)[0]


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    p.add_argument("source", type=Path)
    p.add_argument("output", type=Path)
    p.add_argument("--month", default="01")
    p.add_argument("--year", default="2025")
    a = p.parse_args(argv)
    src, out = a.source, a.output
    info_file = next((f for d in src.iterdir() if d.is_dir() and d.name.upper() in ("INFO", "INFORMATION") for f in d.glob("*.xlsx")), None)
    tin = None
    names = []
    for f in sorted(src.rglob("*.xlsx")):
        m = re.search(r"([A-Z]\d{3}-\d{9})", f.name)
        if m:
            tin = m.group(1)
            break
    toi = next((f for d in src.iterdir() if d.is_dir() and d.name.upper() == "TOI" for f in sorted(d.glob("*.xlsx"))), None)
    if toi is not None:
        ws = load_workbook(toi, read_only=True)["Step1"]
        for r in ws.iter_rows(values_only=True):
            for v in r:
                if isinstance(v, str) and v.startswith("ឈ្មោះសហគ្រាស"):
                    names.append(v.partition("៖")[2].strip())
    for f in sorted(src.rglob("*.xlsx")):
        if re.search(r"general.ledger|_gl_", f.name, re.I):
            first = f.stem.split("_")[0]
            if len(first) > 3:
                names.append(first)
    names += ["EASTEX GARMENT"] if any("អ៊ីស្ទេច" in n for n in names) else []
    fake = Fake(names, tin)
    out.mkdir(parents=True, exist_ok=True)
    done = []
    for d in sorted(x for x in src.iterdir() if x.is_dir() and not x.name.startswith(".")):
        key = d.name.upper()
        target = out / d.name
        if key in ("PR", "SR", "TOS", "WHT"):
            year_dir = d / a.year if (d / a.year).is_dir() else next((x for x in sorted(d.iterdir()) if x.is_dir()), d)
            f = pick(list(year_dir.glob("*")), a.month, a.year)
            if f is None:
                continue
            (target / year_dir.name).mkdir(parents=True, exist_ok=True)
            dest = target / year_dir.name / re.sub(r"\s*\(\d+\)", "", fake.text(f.name))
            info = anonymise_return(f, dest, fake, key)
            done.append((dest, info))
        elif key == "TOI":
            f = pick(list(d.glob("*")), a.month, a.year)
            if f:
                target.mkdir(parents=True, exist_ok=True)
                dest = target / fake.text(f.name)
                anonymise_toi(f, dest, fake)
                done.append((dest, {}))
        elif key in ("GL", "GENERAL LEDGER", "LEDGER"):
            f = pick(list(d.glob("*")), a.month, a.year)
            if f:
                target.mkdir(parents=True, exist_ok=True)
                dest = target / f"General_Ledger_{a.year}_{a.month}_sample.xlsx"
                anonymise_gl(f, dest, fake, a.month)
                done.append((dest, {}))
        elif key in ("INFO", "INFORMATION") and info_file:
            target.mkdir(parents=True, exist_ok=True)
            dest = target / "Company_Info.xlsx"
            anonymise_info(info_file, dest, fake)
            done.append((dest, {}))
    for dest, info in done:
        print(f"  {dest.relative_to(out)}" + (f"  (name columns {info['names']}, ID columns {info['ids']})" if info else ""))
    print(f"Saved {len(done)} anonymous sample file(s) in {out}. Review them before sharing.")


if __name__ == "__main__":
    sys.exit(main())
