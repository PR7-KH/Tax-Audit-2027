"""Excel-style heading filters and the Audit analysis menus (Revenue, Expense, TOS)."""
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))
import audit_analysis as aa  # noqa: E402
import finance_view as fv  # noqa: E402
import revenue_expense as rx  # noqa: E402
import tos_analysis as ta  # noqa: E402

HEAD = ["ល.រ", "លេខអត្តសញ្ញាណប័ណ្ណ/លិខិតឆ្លងដែន", "ឈ្មោះ", "ភេទ", "សញ្ជាតិ", "មុខតំណែង", "G", "H", "I",
        "ប្រាក់បៀវត្ស", "ពន្ធលើប្រាក់បៀវត្ស", "អត្ថប្រយោជន៍បន្ថែម", "ពន្ធលើអត្ថប្រយោជន៍បន្ថែម", "N"]
STAFF = {  # ID: (name, nationality, salary, months declared)
    "P100": ("WANG LEI", "Chinese", 3_000_000, range(1, 13)),          # foreign, below 4M every month
    "P200": ("MARIA CRUZ", "Filipino", 8_000_000, [1, 2, 4, 5, 6, 7, 8, 9, 10, 11, 12]),  # missing Mar
    "P300": ("LI MING", "Chinese", 9_000_000, range(1, 9)),           # leaves after Aug
    "P400": ("JOHN DOE", "Philippines", 4_000_000, range(3, 13)),     # joins Mar, exactly 4M
    "K001": ("SOK DARA", "ខ្មែរ", 1_500_000, [1, 2, 3, 5, 6, 7, 8, 9, 10, 11, 12]),  # Khmer missing Apr
    "K002": ("CHAN SINA", "Khmer", 1_200_000, range(1, 9)),           # Khmer leaver: not reported
}


def make_tos(root, skip_month=None):
    d = root / "TOS" / "2025"
    d.mkdir(parents=True, exist_ok=True)
    for m in range(1, 13):
        if m == skip_month:
            continue
        wb = Workbook(); ws = wb.active
        ws.append(["Salary tax"]); ws.append(HEAD); ws.append([None] * 14)
        n = 0
        for ident, (name, nat, sal, months) in STAFF.items():
            if m in months:
                n += 1
                ws.append([n, ident, name, "M", nat, "Staff", None, None, None, sal, sal * 0.1, 100, 20, 0])
        ws.append(["សរុប", None, None, None, None, None, None, None, None, 1])
        wb.save(d / f"TOS_{m:02d}_2025.xlsx")


class HeadingFilterTest(unittest.TestCase):
    def table(self):
        t = fv.Table(["Date", "Name", "Amount"], [["2025-01-02", "ACME", 1000], ["2025-02-03", "Beta", 2500],
                                                  ["2025-02-04", "ACME", 400], ["2025-03-01", None, 50]])
        return t

    def test_values_text_and_numbers(self):
        t = self.table()
        self.assertEqual(t.view(filters={1: {"values": {"ACME"}}}), [0, 2])
        self.assertEqual(t.view(filters={1: {"values": {"(Blanks)"}}}), [3])
        self.assertEqual(t.view(filters={2: {"values": None, "text": ">500"}}), [0, 1])
        self.assertEqual(t.view(filters={2: {"text": "100..1000"}, 1: {"values": {"ACME"}}}), [0, 2])
        self.assertEqual(t.view(filters={1: {"text": "be"}}), [1])
        self.assertIn("Name: ACME", t.filter_text({1: {"values": {"ACME"}}}))

    def test_distinct_counts_follow_other_filters(self):
        t = self.table()
        values, complete = t.distinct(1, range(4))
        self.assertTrue(complete)
        self.assertEqual(values, [("ACME", 2), ("Beta", 1), ("(Blanks)", 1)])
        others = t.apply_filters(range(4), {2: {"text": ">500"}, 1: {"values": {"Beta"}}}, skip=1)
        self.assertEqual(t.distinct(1, others)[0], [("ACME", 1), ("Beta", 1)])
        self.assertEqual(t.distinct(2, range(4))[0][0], ("50", 1))      # numbers sorted as numbers

    def test_popup_search_rules(self):
        from heading_filter import matches
        self.assertTrue(matches("4,000,000", "<=4000000"))
        self.assertFalse(matches("4,000,001", "<=4000000"))
        self.assertTrue(matches("Wang Lei", "lei wang"))
        self.assertTrue(matches("1,500", "1000..2000"))

    def test_export_has_excel_filter_buttons(self):
        t = self.table()
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / "x.xlsx"
            t.export(p, [0, 2])
            ws = load_workbook(p).active
            self.assertEqual(ws.auto_filter.ref, "A1:C3")
            self.assertEqual(ws.max_row, 3)


class MonthHelpersTest(unittest.TestCase):
    def test_file_month(self):
        for name, m in (("TOS_03_2025.xlsx", 3), ("TOS 2025-11.xlsx", 11), ("Salary March 2025.xlsx", 3),
                        ("TOS_Jan_2025.xlsx", 1), ("TOS_122025.xlsx", 12)):
            self.assertEqual(aa.file_month(name, 2025), m, name)

    def test_month_list_groups_runs(self):
        self.assertEqual(aa.month_list([(2025, 3), (2025, 4), (2025, 5), (2025, 8)]), "Mar–May 2025, Aug 2025")
        self.assertEqual(aa.month_list([(2024, 12), (2025, 1)]), "Dec 2024 – Jan 2025")


class TosAnalysisTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.root = Path(cls.tmp.name)
        make_tos(cls.root)
        cls.an = ta.build(cls.root)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def rows(self, view):
        t = self.an.views[view]
        return [dict(zip(t.columns, r)) for r in t.rows]

    def test_columns_found_by_heading(self):
        m = self.rows("Column mapping")[0]
        self.assertTrue(m["ID"].startswith("B:"))
        self.assertTrue(m["Nationality"].startswith("E:"))
        self.assertTrue(m["Tax on salary"].startswith("K:"))
        self.assertTrue(m["Fringe benefit"].startswith("L:"))
        self.assertTrue(m["Tax on FB"].startswith("M:"))

    def test_1_foreign_at_or_below_4m(self):
        rows = self.rows("1 Foreign ≤ 4M monthly")
        self.assertEqual({r["ID"] for r in rows}, {"P100", "P400"})      # Khmer staff never listed
        self.assertEqual(len(rows), 12 + 10)
        self.assertEqual(rows[0]["Month"], "2025-01")
        by_person = {r["ID"]: r for r in self.rows("1 by person")}
        self.assertEqual(by_person["P400"]["Months ≤ 4M (list)"], "Mar–Dec 2025")

    def test_2_foreign_missing_month(self):
        rows = self.rows("2 Foreign missing months")
        self.assertEqual([(r["ID"], r["Missing months"]) for r in rows], [("P200", "Mar 2025")])

    def test_3_foreign_left_and_joined(self):
        self.assertEqual([(r["ID"], r["Last declared"], r["Assumed left from"]) for r in self.rows("3 Foreign left")],
                         [("P300", "Aug 2025", "Sep 2025")])
        self.assertEqual([r["ID"] for r in self.rows("3 Foreign joined")], ["P400"])
        head = {r["Month"]: r for r in self.rows("Monthly headcount")}
        self.assertEqual(head["2025-09"]["Foreign left"], 1)
        self.assertEqual(head["2025-09"]["Foreign"], 3)
        self.assertEqual(head["2025-03"]["Foreign joined"], 1)

    def test_4_khmer_missing_month_only(self):
        rows = self.rows("4 Khmer missing months")
        self.assertEqual([(r["ID"], r["Missing months"]) for r in rows], [("K001", "Apr 2025")])
        self.assertNotIn("K002", [r["ID"] for r in self.rows("3 Foreign left")])

    def test_matrix_and_totals_rows_skipped(self):
        matrix = {r["ID"]: r for r in self.rows("Staff × month")}
        self.assertEqual(matrix["P200"]["2025-03"], "MISSING")
        self.assertEqual(matrix["P300"]["2025-10"], "")
        self.assertEqual(len(self.rows("All declarations")), sum(len(list(v[3])) for v in STAFF.values()))

    def test_month_without_file_is_not_a_staff_gap(self):
        with tempfile.TemporaryDirectory() as tmp:
            make_tos(Path(tmp), skip_month=3)
            an = ta.build(tmp)
            summary = {r[0]: r for r in an.views["Summary"].rows}
            self.assertEqual(summary["Months with no TOS file"][1], "Mar 2025")
            self.assertEqual(an.views["2 Foreign missing months"].rows, [])

    def test_workbook_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            p = aa.save_workbook(self.an, Path(tmp) / "tos.xlsx", "Demo")
            wb = load_workbook(p)
            self.assertEqual(wb.sheetnames[0], "Summary")
            ws = wb["1 Foreign ≤ 4M monthly"]
            self.assertEqual(ws.auto_filter.ref, "A4:J26")
            self.assertEqual(ws["F3"].value, "=SUBTOTAL(109,F5:F26)")

    def test_no_folder(self):
        with tempfile.TemporaryDirectory() as tmp:
            an = ta.build(tmp)
            self.assertIn("No TOS folder", an.views["Summary"].rows[0][1])


class RevenueExpenseTest(unittest.TestCase):
    def test_revenue_and_expense_views(self):
        from test_risk_view import make_company
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            make_company(root)
            make_tos(root)
            rev = rx.build_revenue(root)
            s = {r[0]: r for r in rev.views["Summary"].rows}
            self.assertEqual(s["General ledger (sales accounts)"][4], "Base")
            self.assertEqual(s["Sales listing"][4], "Agrees")
            self.assertEqual(s["SR monthly returns"][4], "Not provided")
            self.assertEqual(rev.views["Monthly comparison"].rows[1][3], 200)
            exp = rx.build_expense(root)
            s = {r[0]: r for r in exp.views["Summary"].rows}
            self.assertEqual(s["TOI cost of sales (B4+B5+B6) vs GL 5xxxx"][4], "Agrees")
            self.assertEqual(s["Salary per TOS returns vs GL salary accounts"][1], 262_100_000)
            self.assertIn("By GL account", exp.views)
            self.assertEqual(exp.views["Summary"].amount, [])          # summary lines are not added up


if __name__ == "__main__":
    unittest.main()


# ---------------------------------------------------------------- WHT
PR_HEAD = ["ល.រ", "កាលបរិច្ឆេទ", "លេខវិក្កយបត្រ", "ប្រភេទអ្នកផ្គត់ផ្គង់", "លេខអត្តសញ្ញាណកម្មសារពើពន្ធ", "ឈ្មោះអ្នកផ្គត់ផ្គង់",
           "តម្លៃសរុបលើវិក្កយបត្រ", "ការទិញមិនជាប់អាករ", "តម្លៃនាំចូលជាប់អាករ", "អាករនាំចូល", "ការទិញក្នុងស្រុកជាប់អាករ",
           "អាករលើការទិញក្នុងស្រុក", "M", "N", "O", "បរិយាយ"]
WHT_HEAD = ["ល.រ", "កាលបរិច្ឆេទ", "លេខអត្តសញ្ញាណកម្ម", "ឈ្មោះអ្នកទទួល", "ប្រភេទប្រាក់ចំណូល", "អត្រា", "G", "H", "I",
            "ទឹកប្រាក់ត្រូវបើក", "ប្រាក់ពន្ធកាត់ទុក"]


def make_pr_wht(root):
    for m in range(1, 13):
        lines = [("បុគ្គលមិនជាប់អាករ", "", "SOK LANDLORD", 4_000_000, 0, 0, 0, "Office rental"),
                 ("បុគ្គលមិនជាប់អាករ", "", "CHEA CLEANING", 800_000, 0, 0, 0, "Cleaning service"),
                 ("អ្នកជាប់អាករ", "K001-100123456", "ABC SUPPLY CO., LTD", 11_000_000, 0, 10_000_000, 1_000_000, "Raw material"),
                 ("បុគ្គលមិនជាប់អាករ", "", "MARKET SELLER", 500_000, 0, 0, 0, "Goods: fabric"),
                 ("Overseas", "", "CHINA MACHINE", 20_000_000, 20_000_000, 0, 0, "Import machine")]
        if m == 4:
            lines.append(("បុគ្គលមិនជាប់អាករ", "", "DARA", 2_000_000, 0, 0, 0, ""))
        for kind in ("PR", "WHT"):
            d = root / kind / "2025"
            d.mkdir(parents=True, exist_ok=True)
            wb = Workbook(); ws = wb.active
            ws.append([kind]); ws.append(PR_HEAD if kind == "PR" else WHT_HEAD); ws.append([None]); ws.append([None])
            if kind == "PR":
                for n, (t, tin, name, tot, imp, lv, lvat, desc) in enumerate(lines, 1):
                    ws.append([n, f"2025-{m:02d}-15", f"INV{m}{n}", t, tin, name, tot, tot if not (imp or lv) else 0,
                               imp, 0, lv, lvat, 0, 0, 0, desc])
            else:
                if m != 6:
                    ws.append([1, f"2025-{m:02d}-28", "", "SOK LANDLORD", "Rental", "10%", None, None, None, 4_000_000, 400_000])
                if m == 2:
                    ws.append([2, "2025-02-28", "", "UNKNOWN PAYEE", "Service", "15%", None, None, None, 1_000_000, 150_000])
            wb.save(d / f"{kind}_{m:02d}_2025.xlsx")


class WhtAnalysisTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        import wht_analysis
        cls.tmp = tempfile.TemporaryDirectory()
        make_pr_wht(Path(cls.tmp.name))
        cls.an = wht_analysis.build(cls.tmp.name)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def rows(self, view):
        t = self.an.views[view]
        return [dict(zip(t.columns, r)) for r in t.rows]

    def test_classification(self):
        nature = {r["Supplier"]: r["Nature"] for r in self.rows("PR lines classified")}
        self.assertEqual(nature["SOK LANDLORD"], "Rental")
        self.assertEqual(nature["CHEA CLEANING"], "Service / fee")
        self.assertEqual(nature["ABC SUPPLY CO., LTD"], "Excluded: registered taxpayer (has TIN)")
        self.assertEqual(nature["MARKET SELLER"], "Excluded: goods / material")
        self.assertEqual(nature["CHINA MACHINE"], "Excluded: import")
        self.assertEqual(nature["DARA"], "Unclear")

    def test_both_criteria_required(self):
        import wht_analysis as w
        self.assertEqual(w.classify("Office rental", "SOK", "បុគ្គលមិនជាប់អាករ", 0, False, "")[0], "Rental")
        self.assertTrue(w.classify("Office rental", "SOK", "បុគ្គលមិនជាប់អាករ", 0, False, "L001-100163327")[0].startswith("Excluded"))
        self.assertTrue(w.classify("Office rental", "SOK", "អ្នកជាប់អាករ", 0, False, "")[0].startswith("Excluded: supplier type"))

    def test_declared_months_are_ok_and_only_gaps_are_raised(self):
        ok = {(r["Month"], r["Supplier"]) for r in self.rows("Declared OK")}
        self.assertEqual(len(ok), 11)
        self.assertNotIn(("2025-06", "SOK LANDLORD"), ok)          # July's WHT line is not used twice
        risk = self.rows("WHT risk")
        rent = [r for r in risk if r["Supplier"] == "SOK LANDLORD"]
        self.assertEqual([(r["Month"], r["Estimated WHT at risk"]) for r in rent], [("2025-06", 400_000)])
        self.assertEqual(sum(r["Supplier"] == "CHEA CLEANING" for r in risk), 12)
        unclear = [r for r in risk if r["Supplier"] == "DARA"][0]
        self.assertIn("Check the invoice", unclear["Suggested action"])
        self.assertEqual(self.an.views["WHT risk"].row_tags[-1], "warning")   # unclear sorted last, amber

    def test_recurring_and_orphans(self):
        rec = {r["Supplier"]: r for r in self.rows("Recurring not declared")}
        self.assertEqual(rec["CHEA CLEANING"]["Months"], 12)
        self.assertEqual([r["Payee"] for r in self.rows("WHT not in PR")], ["UNKNOWN PAYEE"])
        total = [r for r in self.rows("Summary") if r["Area"] == "Total estimated WHT at risk"][0]
        self.assertEqual(total["Amount (KHR)"], 400_000 + 12 * 120_000 + 300_000)


class WhtInvoiceMatchTest(unittest.TestCase):
    WHT_HEAD_INV = ["ល.រ", "កាលបរិច្ឆេទ", "លេខវិក្កយបត្រ", "លេខអត្តសញ្ញាណកម្ម", "ឈ្មោះអ្នកទទួល", "ប្រភេទប្រាក់ចំណូល", "អត្រា",
                    "H", "I", "ទឹកប្រាក់ត្រូវបើក", "ប្រាក់ពន្ធកាត់ទុក"]

    def test_invoice_matches_other_name_and_later_month(self):
        import wht_analysis
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            for m in (1, 2, 3):
                for kind in ("PR", "WHT"):
                    d = root / kind / "2025"; d.mkdir(parents=True, exist_ok=True)
                    wb = Workbook(); ws = wb.active
                    ws.append([kind]); ws.append(PR_HEAD if kind == "PR" else self.WHT_HEAD_INV); ws.append([None]); ws.append([None])
                    if kind == "PR" and m == 1:
                        ws.append([1, "2025-01-10", "INV-001", "បុគ្គលមិនជាប់អាករ", "", "MR. A", 3_000_000, 3_000_000, 0, 0, 0, 0, 0, 0, 0, "House rental"])
                        ws.append([2, "2025-01-12", "INV-002", "បុគ្គលមិនជាប់អាករ", "", "MR. B", 1_000_000, 1_000_000, 0, 0, 0, 0, 0, 0, 0, "Repair service"])
                    if kind == "WHT" and m == 3:      # rental declared in March, under another spelling of the name
                        ws.append([1, "2025-03-25", "001", "", "LOK A (LANDLORD)", "Rental", "10%", None, None, 3_000_000, 300_000])
                    wb.save(d / f"{kind}_{m:02d}_2025.xlsx")
            an = wht_analysis.build(root)
            ok = {r[1]: r for r in an.views["Declared OK"].rows}
            self.assertIn("MR. A", ok)
            self.assertIn("invoice", ok["MR. A"][14])
            risk = [r[1] for r in an.views["WHT risk"].rows]
            self.assertEqual(risk, ["MR. B"])                      # only the undeclared purchase is raised
            self.assertTrue(wht_analysis.same_invoice("INV-001", "001"))
            self.assertFalse(wht_analysis.same_invoice("1001", "001"))
