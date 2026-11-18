"""TOI ratio analysis, key risk checks, cash flow, the ADT rules template and column fitting."""
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import toi_analysis as ta  # noqa: E402
from toi_fixture import SCALE, make_info, make_toi  # noqa: E402

M = SCALE


class ToiAnalysisTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def rows(self, an, view):
        t = an.views[view]
        return [dict(zip(t.columns, r)) for r in t.rows]

    def test_read_items_both_years(self):
        t, path, err = ta.load(make_toi(self.root).parent.parent)
        self.assertEqual(err, "")
        self.assertEqual((t.c("A21"), t.p("A21")), (50 * M, 100 * M))
        self.assertEqual(ta.year_of(path), 2025)

    def test_cash_flow_reconciles_with_step10(self):
        make_toi(self.root)
        an = ta.build_cashflow(self.root)
        lines = {r["Line"].strip(): r["Amount (KHR)"] for r in self.rows(an, "Indirect method")}
        self.assertEqual(lines["Net cash from operating activities"], 200 * M)
        self.assertEqual(lines["Purchase of long-term assets (TOI step 10)"], -900 * M)
        self.assertEqual(lines["Net cash from financing activities"], 200 * M)
        self.assertEqual(lines["Net increase / (decrease) in cash"], -500 * M)
        self.assertEqual(lines["Difference (should be 0)"], 0)
        self.assertNotIn("Step 10 assets", an.views)
        direct = {r["Line"].strip(): r["Amount (KHR)"] for r in self.rows(an, "Direct method")}
        self.assertEqual(direct["Cash received from customers"], (5000 - 400) * M)
        self.assertEqual(direct["Net cash from operating activities"], 200 * M)
        self.assertEqual(direct["Difference (should be 0)"], 0)
        self.assertTrue(any("cash decreased" in r["Finding"] for r in self.rows(an, "Findings")))

    def test_cash_flow_without_step10_derives_purchases(self):
        make_toi(self.root, step10=False)
        an = ta.build_cashflow(self.root)
        lines = {r["Line"].strip(): r["Amount (KHR)"] for r in self.rows(an, "Indirect method")}
        self.assertEqual(lines["Purchase of long-term assets (derived: change in A1 + depreciation)"], -900 * M)
        self.assertEqual(lines["Difference (should be 0)"], 0)

    def test_step10_disposals_balance(self):
        # step 10 says 900 bought, but the balance sheet only explains 700: 200 left via disposals
        make_toi(self.root, values={"A7": (2400, 2000), "A1": (3400, 3000), "A0": (5600, 4900), "A28": (5600, 4900),
                                    "A38": (3000, 3000), "A37": (3000, 3000)})
        an = ta.build_cashflow(self.root)
        lines = {r["Line"].strip(): r["Amount (KHR)"] for r in self.rows(an, "Indirect method")}
        self.assertEqual(lines["Disposals / other movements in long-term assets (balancing)"], 200 * M)
        self.assertEqual(lines["Difference (should be 0)"], 0)

    def test_ratios_and_key_checks(self):
        make_toi(self.root)
        an = ta.build_ratios(self.root)
        ratios = {r["Ratio"]: r for r in self.rows(an, "Ratios")}
        self.assertEqual(ratios["Gross profit margin"]["N"], 40.0)
        self.assertEqual(ratios["Gross profit margin"]["N-1"], 35.0)
        self.assertEqual(ratios["Debt to assets"]["Flag"], "High leverage (over 80 %)")
        self.assertEqual(ratios["Income tax to revenue (minimum tax 1 %)"]["Flag"], "Below 1 % of turnover")
        checks = {r["Check"]: r for r in self.rows(an, "Key risk checks")}
        self.assertEqual(checks["Profit made but cash decreased"]["Flag"], "Check")
        self.assertEqual(checks["Debt greater than assets"]["Flag"], "OK")
        self.assertEqual(checks["Gross profit negative"]["Flag"], "OK")
        self.assertEqual(checks["Opening stock D1 vs prior-year closing A15"]["Flag"], "OK")

    def test_debt_over_assets_negative_gp_and_expenses(self):
        make_toi(self.root, values={"A37": (7000, 3000), "A38": (7000, 3000), "A29": (-2720, 700), "B7": (-100, 1400),
                                    "B22": (4000, 1500)})
        checks = {r[0]: r for r in ta.key_checks(ta.load(self.root)[0])}
        self.assertEqual(checks["Debt greater than assets"][4], "Check")
        self.assertEqual(checks["Debt greater than assets"][1], "A37+A42 vs A0")
        self.assertEqual(checks["Negative equity"][4], "Check")
        self.assertEqual(checks["Gross profit negative"][4], "Check")
        self.assertEqual(checks["Operating expenses greater than cost of sales"][4], "Check")
        self.assertEqual(checks["Entertainment expense (B29)"][1], "B29")
        self.assertEqual(checks["Opening stock D1 vs prior-year closing A15"][1], "D1 vs A15")

    def test_no_toi(self):
        an = ta.build_ratios(self.root)
        self.assertIn("No TOI", an.views["Summary"].rows[0][1])


class AdtTemplateTest(unittest.TestCase):
    def test_template_keeps_rules_and_dashboard_calculates(self):
        sys.path.insert(0, str(ROOT / "scripts"))
        import build_adt_template as bt
        import calc
        import engine
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "ADT.xlsx"
            res = bt.build(bt.DEFAULT_SOURCE, out)
            self.assertGreaterEqual(res["kept"], 10)          # the hand-written A2..A12 rules
            ws = load_workbook(out)["ADT"]
            src = load_workbook(bt.DEFAULT_SOURCE)["ADT"]
            self.assertEqual(ws["K6"].value, src["K6"].value)  # A2 Khmer rule unchanged
            # the committed template is what the engine uses: build a company and evaluate it
            root = Path(tmp) / "co"
            root.mkdir()
            make_toi(root); make_info(root)
            from test_audit_analysis import make_tos
            make_tos(root)
            wb, _log = engine.build_workbook(root, 3, False)
            c = calc.Calc(wb)
            d = wb["Audit Dashboard"]
            values = [c.value("Audit Dashboard", r, col) for r in range(1, d.max_row + 1) for col in range(1, 13)]
            self.assertFalse([v for v in values if isinstance(v, str) and v.startswith("#")])
            adt = wb["ADT"]
            key = {adt.cell(r, 4).value: r for r in range(1, adt.max_row + 1)}
            self.assertEqual(c.value("ADT", key["KEY1"], 10), "Check")     # profit made but cash fell
            self.assertEqual(c.value("ADT", key["KEY3"], 10), "OK")        # debt not over assets
            self.assertEqual(c.value("ADT", key["A22"], 10), "Check")
            hidden = {r for r in range(1, d.max_row + 1) if d.row_dimensions[r].hidden}
            self.assertNotIn(key["A22"] + engine.ADT_OFFSET, hidden)       # flagged: shown on the dashboard
            self.assertIn(key["KEY3"] + engine.ADT_OFFSET, hidden)          # OK: hidden there, still in ADT


class FitWidthsTest(unittest.TestCase):
    def test_fit(self):
        sys.path.insert(0, str(ROOT))
        from control_center import fit_widths
        self.assertEqual(fit_widths([100, 100], 300), [150, 150])            # spare space shared
        w = fit_widths([60, 80, 900, 700], 1000)
        self.assertEqual(w[:2], [60, 80])                                      # short columns untouched
        self.assertEqual(w[2], w[3])
        self.assertLessEqual(sum(w), 1000)


if __name__ == "__main__":
    unittest.main()


class AuditReportTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.root = Path(self.tmp.name)
        make_toi(self.root)

    def tearDown(self):
        self.tmp.cleanup()

    def test_word_report_compared_with_toi(self):
        from toi_fixture import make_audit_report
        import audit_report as ar
        make_audit_report(self.root, revenue_khr_m=5200)
        an = ar.build(self.root)
        cmp_ = {r[1]: r for r in an.views["Compare with TOI"].rows}
        self.assertEqual(cmp_["Total assets"][6], "Agrees")
        self.assertEqual(cmp_["Cash and bank"][3], 500 * M)                 # KHR'000 column x 1,000
        self.assertEqual(cmp_["Cost of sales"][6], "Agrees")                 # (3,000) compared as positive
        self.assertEqual(cmp_["Purchase of property, plant and equipment"][6], "Agrees")   # step 10
        self.assertEqual(cmp_["Net cash from operating activities"][6], "Agrees")
        self.assertEqual(cmp_["Revenue"][5], 200 * M)
        self.assertTrue(cmp_["Revenue"][6].startswith("Differs"))
        summary = {(r[0], r[1]) for r in an.views["Summary"].rows}
        self.assertIn(("Audit opinion", "Material uncertainty – going concern"), summary)
        self.assertEqual([n[1] for n in an.views["Notes index"].rows],
                         ["Related party transactions", "Contingent liabilities", "Subsequent events"])

    def test_parse_line_notes_and_brackets(self):
        import audit_report as ar
        self.assertEqual(ar.parse_line("Cost of sales 12 (750,000) (650,000) (3,000,000) (2,600,000)"),
                         ("Cost of sales", "12", [-750000, -650000, -3000000, -2600000]))
        self.assertEqual(ar.parse_line("Borrowings - 1,000")[2], [0.0, 1000.0])
        self.assertIsNone(ar.parse_line("for the year ended 31 December 2025"))

    def test_scanned_or_missing(self):
        import audit_report as ar
        an = ar.build(self.root)
        self.assertIn("No Audit Report folder", an.views["Summary"].rows[0][1])
        (self.root / "Audit Report").mkdir()
        from openpyxl import Workbook  # noqa: F401
        an = ar.build(self.root)
        self.assertIn("no PDF or Word", an.views["Summary"].rows[0][1])


class ArApTest(unittest.TestCase):
    def test_listing_vs_toi_and_aging(self):
        from toi_fixture import make_ar_ap
        import ar_ap
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            make_toi(root); make_ar_ap(root, ar_total_m=850)
            an = ar_ap.build(root)
            s = {(r[0], r[1]): r for r in an.views["Summary"].rows}
            self.assertEqual(s[("A/R", "TOI A18")][3], 50 * M)               # listing 850 vs TOI 800
            self.assertEqual(s[("A/R", "TOI A18")][5], "Differs")
            self.assertEqual(s[("A/P", "TOI A46")][5], "Agrees")
            self.assertEqual(s[("A/P", "Aging")][4], 400 * M)                 # 2023 invoice is over a year old
            ap = {r[0]: r for r in an.views["A/P by supplier"].rows}
            self.assertIn("Over 1 year", ap["OLD VENDOR"][-1])
            ar = {r[0]: r for r in an.views["A/R by customer"].rows}
            self.assertEqual(ar["BETA CO"][-3], 200 * M)                      # 'Over 365 days' bucket


class DownloadRoutingTest(unittest.TestCase):
    def test_new_targets(self):
        import input_tools as it
        self.assertEqual(it.classify_download("Audit Report 2025.pdf")[0], "Audit Report")
        self.assertEqual(it.classify_download("Audited financial statements.docx")[0], "Audit Report")
        self.assertEqual(it.classify_download("AP aging 2025.xlsx")[0], "AP")
        self.assertIsNone(it.classify_download("memo.docx")[0])


class TaxPaidTest(unittest.TestCase):
    PDF_ROWS = [  # what pdfplumber returns for the GDT Payment History table (garbled Khmer, English kept)
        ["ល.រ\nNo.", "Type of Tax", "January"] + [""] * 11 + ["Total"],
        ["1", "2025", "28,222,356", "20,317,574"] + ["1,000"] * 10 + ["48,559,930"],
        ["", " ប  ំ   ំ2025", "", None] + [None] * 12,
        ["1", "\x00\x00 ពន  ៉តង់\nPatent Tax", None, "5,000,000"] + ["-"] * 11 + ["5,000,000"],
        ["2", "ពន េល\nSalary Tax", None, "14,443,236", "11,578,354"] + ["0"] * 10 + ["26,021,590"],
        ["13", "កៃ\x00ម\ne-Filing for Windows Desktop Application\nLicense Fee"] + ["-"] * 7 + ["400,000"] + ["-"] * 4 + ["400,000"],
    ]

    def test_pdf_rows(self):
        import tax_paid as tp
        rows = tp.parse_rows(self.PDF_ROWS, "x.pdf")
        self.assertEqual([r[1] for r in rows], ["Patent Tax", "Salary Tax", "e-Filing for Windows Desktop Application License Fee"])
        self.assertEqual(rows[1][3][:2], [14443236.0, 11578354.0])
        self.assertEqual(rows[2][3][7], 400000.0)                      # August
        self.assertEqual(rows[0][0], 2025)

    def test_info_sheets_and_compare_with_tos(self):
        import tax_paid as tp
        from openpyxl import Workbook
        from test_audit_analysis import make_tos
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            make_tos(root)                                  # TOS tax 10 % of salary each month of 2025
            wb = Workbook(); wb.active.title = "Sheet1"
            ws = wb.create_sheet("Sheet3")
            ws.append(["ល.រ", "ឆ្នាំបង់ពន្ធ"] + ["m"] * 12 + ["សរុប"])
            ws.append([1, 2025] + [1] * 12 + [12])
            ws.append([1, "ពន្ធលើប្រាក់បៀវត្ស"] + [2270000] * 12 + [27240000])
            ws.append([None, "Salary Tax"])
            (root / "Info").mkdir()
            wb.save(root / "Info" / "info.xlsx")
            an = tp.build(root)
            detail = an.views["Detail"].rows
            self.assertEqual(detail[0][0], "Year 2025")
            self.assertEqual(detail[1][1:3], ["ពន្ធលើប្រាក់បៀវត្ស", "Salary Tax"])
            self.assertEqual(list(an.views), ["Detail", "Summary", "Paid vs declared"])
            cmp_ = an.views["Paid vs declared"].rows
            self.assertTrue(any(r[6] == "Agrees" for r in cmp_))
            self.assertTrue(any(r[6] != "Agrees" for r in cmp_))   # later months declared less (staff left)


class DashboardTest(unittest.TestCase):
    def test_collects_flags_from_every_area(self):
        import risk_dashboard as rd
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            make_toi(root)
            an = {"an_RATIO": ta.build_ratios(root), "an_CASHFLOW": ta.build_cashflow(root)}
            report = {"findings": [{"sev": "High", "area": "Sales", "title": "SR below GL", "amount": 5, "evidence": "x", "action": "y"}]}
            dash = rd.build(root, an, report)
            rows = dash.views["Tax Risk Overview"].rows
            areas = {r[0] for r in rows}
            self.assertIn("1. TOI key risk checks", areas)
            self.assertIn("6. Sales tie-out (tax risk finding)", areas)
            key = [r for r in rows if r[1] == "Profit made but cash decreased"][0]
            self.assertEqual((key[2], key[5]), ("B48 vs A21+A22", "High"))
            overview = {r[1]: r[3] for r in rows if r[0] == "Overview"}
            self.assertEqual(overview["9. Salary tax (TOS)"], "Not run: source files missing")


class ToiStep1Test(unittest.TestCase):
    def test_year_from_step1_not_file_date_and_proper_books(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as tmp:
            path = make_toi(Path(tmp), name="TOI-L001-000000000_2026-04-20.xlsx")
            wb = load_workbook(path)
            ws = wb["step1"]
            ws["A14"] = "ការិយបរិច្ឆេទសារពើពន្ធ (ចំនួនខែ) ៖ ១២ខែ (ចាប់ពី ខែមករា ២០២៥ ដល់ ធ្នូ ២០២៥)"
            ws["A15"] = "ការកាន់កាប់បញ្ជិកាគណនេយ្យ ៖ កាន់កាប់បញ្ជិកាគណនេយ្យត្រឹមត្រូវ"
            wb.save(path)
            ta._YEAR_CACHE.clear()
            self.assertEqual(ta.year_of(path), 2025)
            t, _p, _e = ta.load(Path(tmp))
            self.assertTrue(t.proper_books)
            checks = {r[0]: r for r in ta.key_checks(t)}
            self.assertEqual(checks["Income tax below 1 % of turnover (minimum tax)"][4], "OK")
