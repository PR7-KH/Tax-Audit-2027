"""A small but complete TOI workbook for tests: balanced balance sheets for N and N-1, an income
statement, the cost-of-goods table and a step 10 asset table. The company makes a profit while its
cash falls, holds more stock and receivables, buys plant and pays tax below 1 % of turnover."""
from pathlib import Path

from openpyxl import Workbook

# code: (current N, prior N-1)
VALUES = {
    "A0": (5800, 4900), "A1": (3600, 3000), "A2": (1000, 1000), "A7": (2600, 2000),
    "A13": (2200, 1900), "A15": (900, 500), "A18": (800, 400), "A21": (50, 100), "A22": (450, 900),
    "A28": (5800, 4900), "A29": (1080, 700), "A30": (1000, 1000), "A35": (-300, -200), "A36": (380, -100),
    "A37": (3200, 3000), "A38": (3200, 3000), "A42": (1520, 1200), "A46": (1520, 1200),
    "B0": (5000, 4000), "B2": (5000, 4000), "B5": (3000, 2600), "B7": (2000, 1400), "B22": (1600, 1500),
    "B23": (600, 600), "B27": (200, 200), "B36": (300, 250), "B41": (500, 450), "B42": (400, -100),
    "B46": (400, -100), "B47": (20, 0), "B48": (380, -100),
    "D1": (500, 450), "D2": (3400, 2650), "D7": (3900, 3100), "D8": (900, 500), "D9": (3000, 2600),
}
STEPS = {4: [f"A{i}" for i in range(0, 53)], 5: [f"B{i}" for i in range(0, 49)],
         6: [f"C{i}" for i in range(1, 21)], 7: [f"D{i}" for i in range(1, 10)]}


SCALE = 1_000_000   # amounts in KHR: 5,000 above means 5,000,000,000 riel


def make_toi(root, values=None, step10=True, name="TOI2025.xlsx"):
    values = {k: (v[0] * SCALE if v[0] is not None else None, v[1] * SCALE if v[1] is not None else None)
              for k, v in dict(VALUES, **(values or {})).items()}
    d = Path(root) / "TOI"
    d.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); wb.remove(wb.active)
    for n in range(1, 18):
        ws = wb.create_sheet(f"step{n}")
        ws["A1"] = f"TOI step {n}"
        if n in STEPS:
            ws.append([None]); ws.append([None])
            ws.append(["បរិយាយ", "យោង", "ការិយបរិច្ឆេទនេះ (N)", "ការិយបរិច្ឆេទមុន (N-1)"])
            for code in STEPS[n]:
                cur, pri = values.get(code, (None, None))
                ws.append([f"Item {code}", code, cur, pri])
        if n == 10 and step10:
            ws.append(["ប្រភេទទ្រព្យសកម្ម", "តម្លៃដើមគ្រា", "ការទិញ ឬបន្ថែមក្នុងគ្រា", "ការលក់ចេញ", "តម្លៃចុងគ្រា"])
            for row in (["Plant and machinery", 2500, 900, 0, 3400], ["Land", 1000, 0, 0, 1000], ["សរុប", 3500, 900, 0, 4400]):
                ws.append([row[0]] + [v * SCALE for v in row[1:]])
    path = d / name
    wb.save(path)
    return path


def make_info(root):
    d = Path(root) / "Info"
    d.mkdir(parents=True, exist_ok=True)
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    for row in (["Company", "DEMO CO., LTD"], ["TIN", "K001-000000001"], ["Address", "Phnom Penh"], ["", ""],
                ["Tax year", "2025"], ["Activity", "Trading"]):
        ws.append(row)
    wb.save(d / "Info.xlsx")


def make_audit_report(root, revenue_khr_m=5000, name="Audit Report 2025.docx"):
    """Audited financial statements (Word) matching the TOI fixture, with US$ and KHR'000 columns.
    revenue_khr_m changes revenue (in KHR millions) to create a difference with the TOI."""
    import docx
    d = Path(root) / "Audit Report"
    d.mkdir(parents=True, exist_ok=True)
    doc = docx.Document()
    for p in ("DEMO CO., LTD", "INDEPENDENT AUDITOR'S REPORT",
              "In our opinion, the financial statements present fairly, in all material respects, the financial position "
              "of the Company as at 31 December 2025 and its financial performance for the year ended 31 December 2025.",
              "Material uncertainty related to going concern: the Company depends on its shareholder for funding.",
              "Translation into KHR is at the rate of US$1 = KHR 4,000."):
        doc.add_paragraph(p)

    def table(title, rows):
        doc.add_paragraph(title)
        doc.add_paragraph("for the year ended 31 December 2025")
        t = doc.add_table(rows=0, cols=6)
        for r in [["", "Note", "2025 US$", "2024 US$", "2025 KHR'000", "2024 KHR'000"]] + rows:
            cells = t.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = v

    def line(label, note, cur_m, pri_m, neg=False):
        f = (lambda v: f"({v:,.0f})") if neg else (lambda v: f"{v:,.0f}")
        return [label, note, f(cur_m * 1e6 / 4000), f(pri_m * 1e6 / 4000), f(cur_m * 1000), f(pri_m * 1000)]
    table("STATEMENT OF FINANCIAL POSITION", [
        line("Property, plant and equipment", "4", 3600, 3000), line("Total non-current assets", "", 3600, 3000),
        line("Inventories", "5", 900, 500), line("Trade receivables", "6", 800, 400), line("Cash and cash equivalents", "7", 500, 1000),
        line("Total current assets", "", 2200, 1900), line("Total assets", "", 5800, 4900),
        line("Share capital", "8", 1000, 1000), line("Accumulated losses", "", 80, -300), line("Total equity", "", 1080, 700),
        line("Amounts due to related parties", "9", 3200, 3000), line("Total non-current liabilities", "", 3200, 3000),
        line("Trade and other payables", "10", 1520, 1200), line("Total current liabilities", "", 1520, 1200),
        line("Total liabilities", "", 4720, 4200), line("Total equity and liabilities", "", 5800, 4900)])
    gp = revenue_khr_m - 3000
    table("STATEMENT OF PROFIT OR LOSS", [
        line("Revenue", "11", revenue_khr_m, 4000), line("Cost of sales", "12", 3000, 2600, neg=True), line("Gross profit", "", gp, 1400),
        line("General and administrative expenses", "13", 1600, 1500, neg=True), line("Profit before income tax", "", gp - 1600, -100),
        line("Income tax expense", "14", 20, 0, neg=True), line("Profit for the year", "", gp - 1620, -100)])
    table("STATEMENT OF CASH FLOWS", [
        line("Depreciation", "4", 300, 250), line("Net cash generated from operating activities", "", 200, 300),
        line("Purchase of property, plant and equipment", "4", 900, 400, neg=True),
        line("Net cash used in investing activities", "", -900, -400), line("Net cash from financing activities", "", 200, 0),
        line("Cash and cash equivalents at end of the year", "", 500, 1000)])
    doc.add_paragraph("NOTES TO THE FINANCIAL STATEMENTS")
    for n, title in ((15, "Related party transactions"), (16, "Contingent liabilities"), (17, "Subsequent events")):
        doc.add_paragraph(f"{n}. {title}")
        doc.add_paragraph("The Company has balances with related parties and a tax audit for 2023 is in progress.")
    path = d / name
    doc.save(path)
    return path


def make_ar_ap(root, ar_total_m=800, aging=True):
    """AR aging (buckets) and AP listing (invoice dates) matching TOI A18 = 800 and A46 = 1,520 (millions)."""
    for folder, rows in (("AR", [["Customer", "Current", "31-90 days", "91-365 days", "Over 365 days", "Balance"],
                                 ["ACME TRADING", 300, 100, 0, 0, 400], ["BETA CO", 100, 0, 100, 150 + (ar_total_m - 800), 350 + (ar_total_m - 800)],
                                 ["GAMMA LTD", 50, 0, 0, 0, 50], ["Total", 450, 100, 100, 150, 800]]),
                         ("AP", [["Supplier", "Invoice date", "Invoice", "Outstanding"],
                                 ["SOK LANDLORD", "2025-12-01", "R12", 120], ["ABC SUPPLY", "2025-11-15", "P77", 1000],
                                 ["OLD VENDOR", "2023-06-30", "X1", 400]])):
        d = Path(root) / folder
        d.mkdir(parents=True, exist_ok=True)
        wb = Workbook(); ws = wb.active
        for r in rows:
            ws.append([v * SCALE if isinstance(v, int) else v for v in r])
        wb.save(d / f"{folder}_listing_2025.xlsx")
