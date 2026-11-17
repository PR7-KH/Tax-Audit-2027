import sys, tempfile, types, unittest
from datetime import date
from pathlib import Path
from unittest import mock
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import finance_view as fv


def write(path, sheets):
    wb = Workbook(); wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    wb.save(path)


def make_folder(root):
    (root / "Sales").mkdir(); (root / "Bank").mkdir(); (root / "GL").mkdir(); (root / "Sales" / ".pdf-extractions").mkdir()
    write(root / "Sales" / "s.xlsx", {
        "2025 Sales": [["No", "Date", "Invoice", "Buyer Name", "Total Amount (KHR)", "Month"],
                       [1, "2025-01-31", "E1", "ACME Trading", 1000000, 1],
                       [2, "2025-02-10", "E2", "Beta Garment", 250000, 2],
                       [3, "2025-02-11", "E3", "ACME Trading", 4000000, 2]],
        "Summary": [["Month", "Sales", "Invoices"], ["Jan", 1, 1]]})
    write(root / "Sales" / ".pdf-extractions" / "r.xlsx", {"x": [["a", "b", "c"], [1, 2, 3]]})
    write(root / "GL" / "gl.xlsx", {"GL Transactions": [
        ["Account Code", "Account Name", "Date", "Debit", "Credit"],
        ["40101", "Sales of Goods", date(2025, 1, 5), 0, 500],
        ["10101", "Cash", date(2025, 1, 5), 500, 0]]})
    write(root / "Bank" / "A_bank.xlsx", {"Transactions": [
        ["BANK A | 2025"], [], ["Date", "Currency", "Bank Description", "Money In / Credit", "Money Out / Debit"],
        [date(2025, 3, 1), "USD", "FT B/O GLOBEX LTD,HK", 100, 0],
        [date(2025, 3, 2), "USD", "service charge", 0, 5]]})
    write(root / "Bank" / "B_bank.xlsx", {"Transactions": [
        ["Posting Date", "Description", "Money In (USD)", "Money Out (USD)"],
        [date(2025, 1, 4), "Credit interest", 3.5, 0]]})


class ModelTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(); self.root = Path(self.tmp.name); make_folder(self.root)
        self.sales = fv.load_table(fv.table_files(self.root / "Sales"), "SALES")
        self.gl = fv.load_table(fv.table_files(self.root / "GL"), "GL")
        self.bank = fv.load_table(fv.table_files(self.root / "Bank"), "BANK")

    def tearDown(self):
        self.tmp.cleanup()

    def test_hidden_folders_and_sheet_choice(self):
        self.assertEqual(len(self.sales.rows), 3)            # not the Summary sheet, not .pdf-extractions
        self.assertEqual(self.sales.months, ["2025-01", "2025-02"])

    def test_find_month_and_numbers(self):
        s = self.sales
        self.assertEqual(len(s.view(find="acme")), 2)
        self.assertEqual(len(s.view(find="acme trading", month="2025-02")), 1)
        col = s.columns.index("Total Amount (KHR)")
        self.assertEqual(len(s.view(find=">500,000")), 2)
        self.assertEqual(len(s.view(find="200000..300000", column="Total Amount (KHR)")), 1)
        self.assertEqual(s.view(find="1000000", column="Total Amount (KHR)"), [0])
        self.assertEqual(s.view(find="nothing"), [])
        self.assertEqual(col, 4)

    def test_sort_and_totals(self):
        s = self.sales; col = s.columns.index("Total Amount (KHR)")
        self.assertEqual(s.view(sort=(col, True)), [2, 0, 1])
        self.assertEqual(s.view(sort=(col, False)), [1, 0, 2])
        self.assertEqual(s.totals(s.view())[0][1], 5250000)
        self.assertEqual(s.display(2)[col], "4,000,000")

    def test_gl_keeps_debit_credit(self):
        self.assertEqual(self.gl.columns[-2:], ["Debit", "Credit"])
        idx = self.gl.view(find="sales", column="Account Name")
        self.assertEqual(self.gl.totals(idx), [("Credit", 500)])

    def test_bank_alignment_and_type(self):
        b = self.bank
        self.assertEqual(b.columns[:6], ["Bank", "Date", "Account", "Currency", "Description", "Money In"])
        self.assertEqual(len(b.rows), 3)
        self.assertEqual([r[b.columns.index("Date")] for r in b.rows], sorted(r[b.columns.index("Date")] for r in b.rows))
        types = {r[b.columns.index("Description")]: r[b.columns.index("Type")] for r in b.rows}
        self.assertEqual(types["FT B/O GLOBEX LTD,HK"], "Foreign remittance")
        self.assertEqual(types["Credit interest"], "Interest")
        self.assertEqual(dict(b.totals(b.view()))["Money In USD"], 103.5)

    def test_summary_text_after_filters(self):
        s = self.sales
        self.assertEqual(s.summary(s.view(month="2025-02")), "2 rows   |   Total Amount (KHR): 4,250,000")
        self.assertEqual(s.summary(s.view(field="Buyer Name", value="ACME Trading")), "2 rows   |   Total Amount (KHR): 5,000,000")
        self.assertEqual(s.summary([]), "0 rows")
        g = self.gl
        self.assertEqual(g.summary(g.view(field="Account Name", value="Cash")), "1 rows   |   Debit: 500   |   Credit: 0   |   Debit − Credit: 500")
        b = self.bank
        text = b.summary(b.view(month="2025-03"))
        self.assertIn("Money In USD: 100", text); self.assertIn("In − Out USD: 95", text)

    def test_ar_balance_is_summed_but_running_balance_is_not(self):
        t = fv.Table(["Customer", "Balance", "Running Balance"], [["A", 10, 5], ["B", 20, 25]])
        self.assertEqual(t.summary([0, 1]), "2 rows   |   Balance: 30")

    def test_export_filtered(self):
        out = self.root / "o.xlsx"
        self.sales.export(out, self.sales.view(find="acme"))
        ws = load_workbook(out).active
        self.assertEqual(ws.max_row, 3)


class FakeVar:
    def __init__(self, v=""): self.v = v
    def get(self): return self.v
    def set(self, v): self.v = v


class FakeTree:
    def __init__(self): self.cols, self.heads, self.items = (), {}, []
    def __getitem__(self, k): return self.cols
    def __setitem__(self, k, v): self.cols = tuple(v)
    def delete(self, *a): self.items = []
    def get_children(self): return []
    def winfo_width(self): return 1000
    def heading(self, cid, **kw): self.heads.setdefault(cid, {}).update(kw)
    def column(self, *a, **kw): pass
    def insert(self, parent, pos, values, tags): self.items.append(list(values))


def import_control_center():
    try:
        import tkinter  # noqa: F401
    except ImportError:
        class Any:
            def __init__(self, *a, **k): pass
        mod = lambda name: types.ModuleType(name)
        tk = mod("tkinter"); tk.__getattr__ = lambda n: Any
        for n in ("Tk", "Label", "Frame", "Canvas", "Entry", "Spinbox", "Checkbutton", "StringVar", "IntVar", "BooleanVar", "TclError"):
            setattr(tk, n, Any)
        ttk = mod("tkinter.ttk"); fd = mod("tkinter.filedialog"); mb = mod("tkinter.messagebox")
        for n in ("Combobox", "Treeview", "Scrollbar", "Label", "Style"):
            setattr(ttk, n, Any)
        tk.ttk, tk.filedialog, tk.messagebox = ttk, fd, mb
        sys.modules.update({"tkinter": tk, "tkinter.ttk": ttk, "tkinter.filedialog": fd, "tkinter.messagebox": mb})
    import control_center
    return control_center


class ControlCenterDataViewTest(unittest.TestCase):
    def setUp(self):
        self.cc = import_control_center()
        self.tmp = tempfile.TemporaryDirectory(); self.root = Path(self.tmp.name); make_folder(self.root)
        a = self.cc.App.__new__(self.cc.App)
        a.tree, a.fit_columns, a.zoom, a.page_size, a.busy = FakeTree(), True, 90, 2, False
        a.folder, a.data_find, a.data_column, a.data_month = FakeVar(str(Path(str(self.root)))), FakeVar(), FakeVar(), FakeVar()
        a.data_value, a.data_choices, a.data_value_box, a.data_totals = FakeVar("All"), {}, mock.Mock(), mock.Mock()
        a.data_tables, a.data_state, a.last_table, a.current = {}, {}, None, {"key": "fin_SALES"}
        a.job_settings = (str(self.root), 3, False)
        a.company_label, a.page_label = mock.Mock(), mock.Mock()
        a.data_column_box, a.data_month_box, a.source_controls, a.data_controls = (mock.Mock() for _ in range(4))
        a.picker, a.btn_export, a.status_text = mock.Mock(), mock.Mock(), []
        a.status = a.status_text.append
        a.set_picker = lambda *x: None
        a.run_bg = lambda fn, *args: setattr(a, "ran", fn(*args))
        self.app = a

    def tearDown(self):
        self.tmp.cleanup()

    def load(self, key="fin_SALES"):
        import source_check
        a = self.app
        res = source_check.inspect(self.root)
        a.inspection = res
        for fkey, table in res["tables"].items():
            a.data_tables[(res["folder"], "fin_" + fkey)] = table
        a.show_data(key)

    def test_show_filter_sort_page(self):
        a = self.app; self.load()
        self.assertEqual(a.tree.items[0][3], "ACME Trading")
        self.assertEqual(len(a.tree.items), 2)                     # page size 2 of 3 rows
        self.assertIn("rows 1–2 of 3", a.status_text[-1]); self.assertIn("5,250,000", a.data_totals.config.call_args.kwargs["text"])
        a.change_page(1); self.assertEqual(len(a.tree.items), 1)
        self.assertIn("5,250,000", a.data_totals.config.call_args.kwargs["text"])
        a.data_find.set("acme"); a.apply_data_filter()
        self.assertIn("2 rows", a.data_totals.config.call_args.kwargs["text"])
        self.assertIn("filtered from 3", a.status_text[-1]); self.assertEqual(a.data_state["page"], 0)
        a.sort_data(4); self.assertEqual(a.tree.items[0][4], "1,000,000"); self.assertIn("▲", a.tree.heads["c4"]["text"])
        a.sort_data(4); self.assertEqual(a.tree.items[0][4], "4,000,000"); self.assertIn("▼", a.tree.heads["c4"]["text"])
        a.data_month.set("2025-01"); a.apply_data_filter(); self.assertEqual(len(a.tree.items), 1)
        a.clear_data_filter(); self.assertEqual(a.data_find.get(), ""); self.assertEqual(len(a.data_state["indices"]), 3)

    def test_value_picker_lists_values_and_filters(self):
        a = self.app; self.load()
        self.assertEqual(a.data_column.get(), "Buyer Name")                      # default column for Sales
        labels = a.data_value_box.config.call_args.kwargs["values"]
        self.assertEqual(labels, ["All", "ACME Trading", "Beta Garment"])
        a.data_value.set("ACME Trading"); a.apply_data_filter()
        self.assertEqual(len(a.data_state["indices"]), 2)
        a.data_column.set("Month"); a.data_column_changed()                      # switching column resets the value
        self.assertEqual(a.data_value.get(), "All"); self.assertEqual(len(a.data_state["indices"]), 3)
        a.data_value.set("2"); a.apply_data_filter(); self.assertEqual(len(a.data_state["indices"]), 2)

    def test_gl_account_code_and_name_lists(self):
        a = self.app; self.load("fin_GL")
        self.assertEqual(a.data_column.get(), "Account Name")
        self.assertEqual(a.data_value_box.config.call_args.kwargs["values"], ["All", "Cash", "Sales of Goods"])
        a.data_column.set("Account Code"); a.data_column_changed()
        self.assertIn("40101 · Sales of Goods", a.data_value_box.config.call_args.kwargs["values"])
        a.data_value.set("40101 · Sales of Goods"); a.apply_data_filter()
        self.assertEqual(len(a.data_state["indices"]), 1)

    def test_bank_one_account_only(self):
        a = self.app; self.load("fin_BANK")
        self.assertEqual(a.data_column.get(), "Account")
        labels = a.data_value_box.config.call_args.kwargs["values"]
        self.assertEqual(labels[0], "All"); self.assertEqual(len(labels), 3)       # two bank accounts in the fixture
        a.data_value.set(labels[1]); a.apply_data_filter()
        accounts = {a.current_table().rows[i][a.current_table().columns.index("Account")] for i in a.data_state["indices"]}
        self.assertEqual(accounts, {labels[1]})

    def test_gl_and_bank_load(self):
        self.load("fin_GL"); self.assertEqual(self.app.tree.cols and len(self.app.tree.cols), 5)
        self.load("fin_BANK"); self.assertIn("Type", self.app.current_table().columns)

    def test_pdf_only_folder_shows_its_inspect_lines(self):
        import source_check
        (self.root / "TB").mkdir(); (self.root / "TB" / "tb.pdf").write_bytes(b"%PDF")
        a = self.app
        a.inspection = source_check.inspect(self.root)
        a.draw = mock.Mock()
        a.show_not_ready("fin_TB")
        rows = a.draw.call_args.args[1]
        self.assertEqual(rows[0][1], "Trial balance"); self.assertEqual(rows[0][0], "✗")
        self.assertIn("tb.pdf", [r[3] for r in rows])


if __name__ == "__main__":
    unittest.main()
