import sys, tempfile, unittest
from datetime import date
from pathlib import Path
from openpyxl import Workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import sales_reconciliation as sr


def make_company(root, listing_amounts, gl_sales_credit):
    for d in ("Sales", "GL", "Bank"):
        (root / d).mkdir()
    wb = Workbook(); ws = wb.active; ws.title = "2025 Sales"
    ws.append(["No", "Date", "Invoice / Customs Declaration", "Buyer Name", "Total Amount (KHR)", "Month"])
    for i, amt in enumerate(listing_amounts, 1):
        ws.append([i, f"2025-0{i}-15", f"E{i}", "ACME TRADING", amt, i])
    wb.save(root / "Sales" / "sales.xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "GL Transactions"
    ws.append(["Account Code", "Account Name", "Date", "Period", "Transaction Type", "Reference No.",
               "Name / Counterparty", "Memo", "Split Code", "Split Account", "Debit", "Credit"])
    for i, amt in enumerate(gl_sales_credit, 1):
        d = date(2025, i, 15)
        ws.append(["40101", "Sales of Goods", d, f"2025-0{i}", "Invoice", None, "ACME", None, "11001", "Accounts Receivable", 0, amt])
        ws.append(["11001", "Accounts Receivable", d, f"2025-0{i}", "Payment", None, "OTHERCO", None, "10101", "Cash on Hand", 0, amt])
    wb.save(root / "GL" / "gl.xlsx")
    wb = Workbook(); ws = wb.active; ws.title = "Transactions"
    ws.append(["Date", "Currency", "Bank Description", "Money In / Credit"])
    ws.append([date(2025, 1, 15), "USD", "FT B/O OTHERCO LTD,HK", 1000])
    wb.save(root / "Bank" / "Test_bank.xlsx")


class SalesReconciliationTest(unittest.TestCase):
    def run_case(self, listing, gl):
        with tempfile.TemporaryDirectory() as t:
            root = Path(t); make_company(root, listing, gl)
            return sr.run(root, root / "out.xlsx", fx=4000, tol=1)

    def test_agreeing_sources(self):
        a = self.run_case([100, 200], [100, 200])
        sev = {f["area"]: f["sev"] for f in a["findings"]}
        self.assertEqual(sev["Sales listing vs GL"], "OK")
        self.assertEqual(a["list_m"][2], 200)

    def test_difference_is_flagged_high(self):
        a = self.run_case([100, 250], [100, 200])
        f = [x for x in a["findings"] if x["area"] == "Sales listing vs GL"][0]
        self.assertEqual((f["sev"], f["amount"]), ("High", 50))

    def test_payer_not_buyer_and_cash(self):
        a = self.run_case([100, 200], [100, 200])
        areas = {x["area"]: x["sev"] for x in a["findings"]}
        self.assertEqual(areas["Payer vs buyer"], "High")
        self.assertEqual(areas["Cash collections"], "High")

    def test_bank_fx_conversion(self):
        a = self.run_case([100], [100])
        self.assertEqual(a["bank_m"][1], 4_000_000)

    def test_classify_credit(self):
        self.assertEqual(sr.classify_credit("FUND TRF FR NATIONAL BANK OF CAMBODIA", "KHR"), "NBC / government fund transfer")
        self.assertEqual(sr.classify_credit("IN HOUSE CHQ TRF FR CA 1 TO CA 2", "USD"), "In-house transfer received")
        self.assertEqual(sr.classify_credit("IN HOUSE CHQ TRF TO KHR (EXCHANGE)", "KHR"), "Own-account exchange transfer")


if __name__ == "__main__":
    unittest.main()
