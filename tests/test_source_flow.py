"""Source inspect (check + validate in one pass), PDF standardisation, and 'Build all data'
leaving every menu ready without a background job."""
import queue
import sys
import tempfile
import types
import unittest
from collections import OrderedDict
from datetime import date
from pathlib import Path
from unittest import mock

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import source_check as sc  # noqa: E402

SR_HEAD = "តម្លៃសរុបលើវិក្កយបត្រ"


def write(path, sheets):
    wb = Workbook(); wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def by_source(res, source, scope=None):
    return [r for r in res["rows"] if r["values"][1] == source and (scope is None or r["values"][2] == scope)]


class FakeCheck:
    """engine.check_sources replacement: every monthly .xlsx reads 5 rows."""
    def __init__(self, root):
        self.root = root

    def __call__(self, folder):
        out = []
        for tax in ("PR", "SR", "TOS", "WHT"):
            d = Path(folder) / tax
            for p in sorted(d.rglob("*.xlsx")) if d.exists() else []:
                out.append((tax, p.parent.name, p.name, 0 if "empty" in p.name else 5, "READY"))
        return out


class InspectTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.root = Path(self.tmp.name)
        for m in range(1, 13):
            (self.root / "SR" / "2024").mkdir(parents=True, exist_ok=True)
            if m != 3:
                (self.root / "SR" / "2024" / f"SALE_{m:02d}_2024.xlsx").touch()
            (self.root / "SR" / "2025").mkdir(exist_ok=True)
            (self.root / "SR" / "2025" / f"SALE_{m:02d}_2025.xlsx").touch()
        for m in range(1, 4):
            (self.root / "PR" / "2026").mkdir(parents=True, exist_ok=True)
            (self.root / "PR" / "2026" / f"PURCHASE_{m:02d}_2026.xlsx").touch()
        (self.root / "WHT" / "2025").mkdir(parents=True)
        (self.root / "WHT" / "2025" / "WHT_01_2025.pdf").write_bytes(b"%PDF")
        write(self.root / "TOI" / "toi.xlsx", {"Step5": [["a"]]})
        write(self.root / "GL" / "gl.xlsx", {"GL Transactions": [["Account Code", "Account Name", "Date", "Debit", "Credit"],
                                                                 ["40101", "Sales", date(2025, 1, 5), 0, 500]]})
        write(self.root / "Sales" / "sales.xlsx", {"Sales": [["Date", "Invoice", "Buyer", "Amount"], ["2025-01-02", "E1", "A", 10]]})
        write(self.root / "Sales" / "sales copy.xlsx", {"Sales": [["Date", "Invoice", "Buyer", "Amount"], ["2025-01-02", "E1", "A", 10]]})
        (self.root / "Bank").mkdir(); (self.root / "Bank" / "stmt.pdf").write_bytes(b"%PDF")
        write(self.root / "TB" / "tb.xlsx", {"TB": [["Code", "Name", "Note"], ["1", "Cash", "x"]]})
        (self.root / "Download").mkdir(); (self.root / "Download" / "mystery.xlsx").touch()
        patcher = mock.patch.object(sc.engine, "check_sources", FakeCheck(self.root))
        patcher.start(); self.addCleanup(patcher.stop)
        self.res = sc.inspect(self.root, [["Download/mystery.xlsx", "", "SKIPPED: Unrecognized filename"]])

    def tearDown(self):
        self.tmp.cleanup()

    def test_tax_years_ready_missing_months_and_pdf(self):
        sr25 = by_source(self.res, "SR", "2025")[0]
        self.assertEqual((sr25["mark"], sr25["values"][3]), ("✓", 12))
        self.assertIn("12 months (Jan–Dec)", sr25["values"][5])
        sr24 = by_source(self.res, "SR", "2024")[0]
        self.assertEqual(sr24["mark"], "⚠"); self.assertIn("missing Mar", sr24["values"][5])
        pr26 = by_source(self.res, "PR", "2026")[0]
        self.assertEqual(pr26["mark"], "✓"); self.assertIn("3 months (Jan–Mar), newest year", pr26["values"][5])
        wht = by_source(self.res, "WHT", "2025")[0]
        self.assertEqual(wht["mark"], "✗"); self.assertIn("must be the Excel", wht["values"][5])
        self.assertEqual(by_source(self.res, "TOS")[0]["mark"], "✗")            # missing folder
        self.assertEqual(by_source(self.res, "Info")[0]["mark"], "✗")

    def test_finance_ready_duplicate_pdf_and_unusable(self):
        self.assertEqual(by_source(self.res, "General ledger")[0]["mark"], "✓")
        self.assertIn("GL", self.res["tables"])
        sales = by_source(self.res, "Sales")[0]
        self.assertEqual(sales["mark"], "⚠"); self.assertIn("keep one", sales["values"][5])
        self.assertEqual(len(self.res["tables"]["SALES"].rows), 1)           # duplicate file not loaded twice
        bank = by_source(self.res, "Bank")[0]
        self.assertEqual(bank["mark"], "✗"); self.assertIn("Convert PDFs", bank["values"][5])
        self.assertEqual([p.name for p in self.res["pending_pdfs"]], ["stmt.pdf"])
        tb = by_source(self.res, "Trial balance")[0]
        self.assertEqual(tb["mark"], "✗"); self.assertIn("No amount column", tb["values"][5])
        self.assertEqual(by_source(self.res, "A/R")[0]["mark"], "—")

    def test_download_rows_carry_the_import_reason(self):
        self.assertEqual(by_source(self.res, "Download")[0]["mark"], "⚠")
        f = [r for r in self.res["files"] if r["values"][1] == "Download"][0]
        self.assertEqual(f["mark"], "✗"); self.assertIn("Unrecognized filename", f["values"][5])

    def test_converted_pdf_with_same_name_excel_becomes_ready(self):
        write(self.root / "Bank" / "stmt.xlsx", {"Transactions": [["Date", "Description", "Money In"], [date(2025, 2, 1), "x", 5]]})
        res = sc.inspect(self.root)
        self.assertEqual(by_source(res, "Bank")[0]["mark"], "✓")
        pdf = [r for r in res["files"] if r["values"][3] == "stmt.pdf"][0]
        self.assertEqual(pdf["mark"], "✓"); self.assertIn("stmt.xlsx is used", pdf["values"][5])
        self.assertEqual(res["pending_pdfs"], [])

    def test_banner_and_display(self):
        self.assertIn("✗", sc.banner(self.res))
        rows, tags, paths = sc.display(self.res)
        self.assertEqual(len(rows), len(tags)); self.assertEqual(len(rows), len(paths))
        rows, tags, paths = sc.display(self.res, per_file=True)
        self.assertTrue(all(isinstance(p, Path) for p in paths))

    def test_file_month(self):
        self.assertEqual(sc.file_month("L001_SALE_01_2025.xlsx", "2025"), 1)
        self.assertEqual(sc.file_month("PR 2025-11.xlsx", "2025"), 11)
        self.assertIsNone(sc.file_month("random.xlsx", "2025"))


class StandardizeTest(unittest.TestCase):
    def extraction(self, path, tables, text_only=False):
        sheets = {"Conversion Review": [["Status", "REVIEW"]]}
        if not text_only:
            for i, t in enumerate(tables, 1):
                sheets[f"P{i} Table 1"] = t
        sheets["P1 Text"] = [["line"]]
        write(path, sheets)

    def test_tables_join_skip_repeated_heading_and_total(self):
        with tempfile.TemporaryDirectory() as t:
            ext, out = Path(t) / "e.xlsx", Path(t) / "stmt.xlsx"
            head = ["Date", "Description", "Money In", "Money Out"]
            self.extraction(ext, [[head, ["01/03/2025", "a", "1,200.50", ""], ["02/03/2025", "fee", "", "(5.00)"]],
                                  [head, ["03/04/2025", "b", "100", ""], ["Total", "", "1,300.50", "5.00"]]])
            ok, why = sc.standardize_extraction(ext, out, "BANK")
            self.assertTrue(ok, why); self.assertIn("3 rows", why)
            mark, remark, part = sc.check_data_file(out, "BANK")
            self.assertEqual(mark, "✓"); self.assertEqual(part[2][1][3], -5)
            self.assertFalse(sc.standardize_extraction(ext, out, "BANK")[0])          # never overwrites

    def test_text_only_and_no_amount_are_refused(self):
        with tempfile.TemporaryDirectory() as t:
            ext = Path(t) / "e.xlsx"
            self.extraction(ext, [], text_only=True)
            self.assertEqual(sc.standardize_extraction(ext), (False, "the PDF gave text lines only, no table"))
            self.extraction(ext, [[["Name", "Address", "Phone"], ["a", "b", "c"]]])
            ok, why = sc.standardize_extraction(ext)
            self.assertFalse(ok); self.assertIn("no amount column", why)

    def test_raw_extraction_file_is_not_ready(self):
        with tempfile.TemporaryDirectory() as t:
            p = Path(t) / "raw.xlsx"
            write(p, {"Conversion Review": [["Status", "REVIEW REQUIRED"], ["Source PDF", "x.pdf"]]})
            mark, remark, _ = sc.check_data_file(p, "GL")
            self.assertEqual(mark, "✗")


def control_center():
    try:
        import tkinter  # noqa: F401
    except ImportError:
        import test_finance_view
        return test_finance_view.import_control_center()
    import control_center as cc
    return cc


class Var:
    def __init__(self, v=""): self.v = v
    def get(self): return self.v
    def set(self, v): self.v = v


class BuildAllMakesMenusInstantTest(unittest.TestCase):
    def setUp(self):
        self.cc = control_center()
        a = self.cc.App.__new__(self.cc.App)
        a.busy, a.q = False, queue.Queue()
        a.folder, a.months, a.link_all = Var("/co"), Var(3), Var(False)
        a.audit_panel = None
        a.preview_card, a.status_lbl = mock.Mock(), mock.Mock()
        for name in ("title_lbl", "sub_lbl", "search_controls", "cons_controls", "source_controls", "data_controls",
                     "search_label", "page_label", "tax_options", "toi_pdf_options", "dashboard_options", "btn_export", "picker"):
            setattr(a, name, mock.Mock())
        a.mark_active = mock.Mock()
        a.need_folder = lambda: True
        a.run_bg = mock.Mock()
        a.status = mock.Mock()
        a.import_results = []
        a.data_tables, a.risk_cache, a.cons_tables, a.cons_state = {}, {}, {}, {}
        a.inspection = None
        a.partial_build = False
        a.built = None
        a.page_cache, a.sheet_pages, a.sheet_dimensions, a.page_size = OrderedDict(), {}, {}, 200
        self.a = a

    def test_before_build_reports_use_background_job(self):
        self.a.open_report("consolidate")
        self.a.run_bg.assert_called_once()

    def test_after_build_tax_reports_open_without_background_job(self):
        wb = Workbook(); wb.active.title = "PR"
        self.a.built = (("/co", 3, False), wb, [], mock.Mock())
        self.a.open_report("consolidate")
        self.a.run_bg.assert_not_called()
        msg = self.a.q.get_nowait()
        self.assertEqual(msg[:2], ("sheets", "consolidate"))
        self.a.months.set(4)                             # changed options must rebuild
        self.a.open_report("consolidate")
        self.a.run_bg.assert_called_once()

    def test_finance_and_risk_menus_use_the_build_cache(self):
        self.a.data_tables[("/co", "fin_GL")] = object()
        self.a.show_data = mock.Mock()
        self.a.open_report("fin_GL")
        self.a.show_data.assert_called_once_with("fin_GL")
        self.a.risk_cache["/co"] = ("table", ["No"], [["1"]], "note", False, ["even"])
        self.a.open_report("tax_risk")
        self.assertEqual(self.a.q.get_nowait()[0], "table")
        self.a.inspection = {"folder": "/co"}
        self.a.show_inspection = mock.Mock()
        self.a.open_report("sources")
        self.a.show_inspection.assert_called_once()
        self.a.run_bg.assert_not_called()

    def test_menu_has_no_separate_source_folder_check(self):
        keys = [k for kind, k, _ in self.cc.MENU]
        self.assertNotIn("check_fin", keys)
        self.assertEqual(self.cc.ACTION_TITLES["sources"], "Source inspect")

    def test_consolidated_sheet_becomes_a_filterable_table(self):
        wb = Workbook(); ws = wb.active; ws.title = "SR"
        width = 19
        ws.append(["Grand Total"] + [None] * 6 + [30] + [None] * 11)
        ws.append(["Year 2025"] + [None] * 6 + [30] + [None] * 11)
        ws.append(["No", "Date", "Invoice"] + [None] * 4 + [SR_HEAD] + [None] * 11)
        ws.append([None] * width)
        ws.append([1, "05-01-2025", "E1", None, None, None, "ACME"] + [10] + [0] * 11)
        ws.append([2, "07-02-2025", "E2", None, None, None, "BETA"] + [20] + [0] * 11)
        self.a.built = (("/co", 3, False), wb, [], mock.Mock())
        t = self.a.cons_table("SR")
        self.assertEqual(t.columns[:4], ["Row", "No", "Date", "Invoice"])
        self.assertEqual(len(t.rows), 2); self.assertEqual(t.months, ["2025-01", "2025-02"])
        idx = t.view(month="2025-02")
        self.assertEqual([t.rows[i][0] for i in idx], [6])                   # Excel row number kept
        self.assertIn(f"{SR_HEAD}: 20", t.summary(idx))


if __name__ == "__main__":
    unittest.main()
