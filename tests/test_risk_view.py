import sys, tempfile, unittest
from datetime import date
from pathlib import Path
from unittest import mock
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import risk_view as rv


def write(path, sheets):
    wb = Workbook(); wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    wb.save(path)


GL_HEAD = ["Account Code", "Account Name", "Date", "Transaction Type", "Name / Counterparty", "Split Account", "Debit", "Credit", "Running Balance"]


def make_company(root, toi_sales=300, tb=True, ar_list=True):
    for d in ("Sales", "GL", "Bank", "TOI") + (("TB",) if tb else ()) + (("AR",) if ar_list else ()):
        (root / d).mkdir()
    write(root / "Sales" / "s.xlsx", {"2025 Sales": [["No", "Date", "Invoice", "Buyer Name", "Total Amount (KHR)"],
        [1, "2025-01-10", "E1", "ACME", 100], [2, "2025-02-10", "E2", "ACME", 200]]})
    write(root / "GL" / "gl.xlsx", {"GL Transactions": [GL_HEAD,
        ["11001", "Accounts Receivable", date(2025, 1, 10), "Invoice", "ACME", "Sales", 100, 0, 1100],
        ["40101", "Sales of Goods", date(2025, 1, 10), "Invoice", "ACME", "AR", 0, 100, -100],
        ["11001", "Accounts Receivable", date(2025, 2, 10), "Invoice", "ACME", "Sales", 200, 0, 1300],
        ["40101", "Sales of Goods", date(2025, 2, 10), "Invoice", "ACME", "AR", 0, 200, -300],
        ["11001", "Accounts Receivable", date(2025, 2, 20), "Payment", "OTHER", "Cash on Hand", 0, 50, 1250],
        ["10101", "Cash on Hand", date(2025, 2, 20), "Payment", "OTHER", "AR", 50, 0, 50],
        ["10201", "Vattanac Bank (USD)", date(2025, 3, 1), "Payment", "OTHER", "Accounts Receivable", 400000, 0, 400000],
        ["50100", "Cost of material", date(2025, 3, 1), "Bill", "SUP", "AP", 80, 0, 80],
        ["61200", "Other service expense", date(2025, 3, 1), "Bill", "SUP", "AP", 20, 0, 20],
        ["69200", "Income tax Expense", date(2025, 12, 31), "JE", "", "", 999, 0, 999]]})
    write(root / "Bank" / "Vattanac_bank.xlsx", {"Transactions": [["Date", "Account", "Currency", "Bank Description", "Money In", "Money Out"],
        [date(2025, 3, 1), "0900-1", "USD", "FT B/O OTHERCO LTD,HK", 100, 0]]})
    write(root / "TOI" / "toi.xlsx", {"Step5": [["ref", "desc", "N", "N-1"], ["", "", "", ""],
        ["revenue", "B0", toi_sales, 0], ["products", "B1", toi_sales, 0], ["cost of sales", "B4", 80, 0], ["operating", "B22", 20, 0]]})
    if tb:
        write(root / "TB" / "tb.xlsx", {"TB": [["Account Code", "Account Name", "Debit", "Credit"],
            ["40101", "Sales of Goods", 0, 300], ["50100", "Cost of material", 80, 0], ["61200", "Other service expense", 20, 0],
            ["69200", "Income tax Expense", 999, 0]]})
    if ar_list:
        write(root / "AR" / "ar.xlsx", {"Aging": [["Customer", "Invoice", "Balance"], ["ACME", "E1", 700], ["ACME", "E2", 550]]})


class RiskViewTest(unittest.TestCase):
    def build(self, **kw):
        with tempfile.TemporaryDirectory() as t:
            root = Path(t); make_company(root, **{k: v for k, v in kw.items() if k in ("toi_sales", "tb", "ar_list")})
            return rv.build_report(root, fx=4000, tol=1)

    def by_no(self, rep):
        return {r["no"]: r for r in rep["rows"]}

    def test_sales_section_agrees_and_flags_difference(self):
        r = self.by_no(self.build())
        self.assertEqual((r["1.3"]["amount"], r["1.3"]["status"]), (300, "Agrees"))
        self.assertEqual((r["1.2"]["amount"], r["1.2"]["status"]), (300, "Agrees"))
        self.assertEqual(r["1.4"]["status"], "Base")
        self.assertEqual((r["1.5"]["amount"], r["1.5"]["status"]), (300, "Agrees"))      # TB
        self.assertEqual(r["1.1"]["status"], "Not provided")                                # no SR folder
        r = self.by_no(self.build(toi_sales=350))
        self.assertEqual((r["1.2"]["diff"], r["1.2"]["status"]), (50, "Differs"))

    def test_cash_and_ar_section(self):
        r = self.by_no(self.build())
        self.assertEqual(r["2.1.1"]["amount"], 400000)                    # USD 100 * 4000
        self.assertEqual(r["2.1.1"]["base"], 400000)                      # matched GL 10201 debits
        self.assertEqual(r["2.1.1"]["status"], "Agrees")
        self.assertEqual(r["2.2.1"]["amount"], 1000)                      # opening = 1100 - 100
        self.assertEqual(r["2.2.4"]["amount"], 50)                        # cash collection
        self.assertEqual(r["2.2.5"]["amount"], 1250)
        self.assertEqual(r["2.2.5"]["status"], "Agrees")
        self.assertEqual((r["2.3"]["amount"], r["2.3"]["status"]), (1250, "Agrees"))     # A/R listing

    def test_expense_section_excludes_income_tax(self):
        r = self.by_no(self.build())
        self.assertEqual(r["3.3"]["amount"], 100)                          # 50100 + 61200, not 69200
        self.assertEqual((r["3.2"]["amount"], r["3.2"]["status"]), (100, "Agrees"))
        self.assertEqual((r["3.4"]["amount"], r["3.4"]["status"]), (100, "Agrees"))      # TB
        self.assertEqual(r["3.1"]["status"], "Not provided")               # no PR folder

    def test_findings_listed_and_missing_sources_say_so(self):
        rep = self.build(tb=False, ar_list=False)
        r = self.by_no(rep)
        self.assertEqual(r["1.5"]["status"], "Not provided"); self.assertEqual(r["2.3"]["status"], "Not provided")
        self.assertTrue(any(f["area"] == "Payer vs buyer" for f in rep["findings"]))
        self.assertTrue(r["4.1"]["status"] in ("High", "Medium", "Info", "OK"))

    def test_display_and_save(self):
        rep = self.build()
        rows, tags = rv.display_rows(rep)
        self.assertEqual(len(rows), len(tags)); self.assertEqual(tags[0], "section")
        with tempfile.TemporaryDirectory() as t:
            out = Path(t) / "o.xlsx"; rv.save_workbook(rep, out)
            wb = load_workbook(out)
            self.assertEqual(wb.sheetnames[0], "Tax risk sections"); self.assertIn("Findings", wb.sheetnames)

    def test_pr_purchase_report_reads_by_invoice_date(self):
        row = lambda d, v: [None, d, "P1", None, None, None, v]
        fake = mock.Mock(rows=[row("05-01-2025", 10), row("06-02-2025", 20), row("07-02-2024", 99)])
        with mock.patch("consolidate.files_in", return_value=[("2025", Path("x.xlsx"))]), \
             mock.patch("consolidate.read_source", return_value=fake):
            out = rv.read_pr(Path("."), 2025)
        self.assertEqual(out["total"], 30); self.assertEqual(out["monthly"], {1: 10, 2: 20})


if __name__ == "__main__":
    unittest.main()
