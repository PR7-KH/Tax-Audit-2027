"""Headless regression checks for inventory, lazy pages and workbook reuse."""
import sys
import tempfile
import time
import unittest
from collections import OrderedDict
from pathlib import Path
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook
import calc
import control_center as ui


class PreviewTests(unittest.TestCase):
    def app(self, wb):
        app = Mock()
        app.built = (('company', 3, False), wb, [], calc.Calc(wb))
        app.job_settings = app.built[0]
        app.page_size = 200
        app.page_cache = OrderedDict()
        app.sheet_pages = {}
        app.sheet_dimensions = {}
        app.busy = False
        return app

    def test_inventory_counts_without_opening_workbooks(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            for year, count in [('2025', 12), ('2026', 13)]:
                directory = root / 'PR' / year
                directory.mkdir(parents=True)
                for n in range(count):
                    (directory / f'{n}.xlsx').touch()
                (directory / '~$lock.xlsx').touch()
            (root / 'GL').mkdir()
            (root / 'GL' / 'ledger.xlsx').touch()
            with patch.object(ui.engine, 'load_workbook', side_effect=AssertionError('must not open')):
                rows = ui.source_inventory(root)
            self.assertIn(['PR', '2025', 12, 'Found; content not validated'], [r[:4] for r in rows])
            self.assertIn(['PR', '2026', 13, 'Found; content not validated'], [r[:4] for r in rows])
            self.assertIn(['General ledger', 'All', 1, 'Found; content not validated'], [r[:4] for r in rows])

    def test_all_rows_available_and_formula_values_preserved(self):
        wb = Workbook(); ws = wb.active; ws.title = 'PR'
        for r in range(1, 8202):
            ws.append([r, f'=A{r}*2'])
        app = self.app(wb)
        started = time.perf_counter()
        first = ui.App.job_page(app, 'PR', 0)
        last = ui.App.job_page(app, 'PR', 41)
        print(f'8201-row workbook: first and final page {time.perf_counter()-started:.3f}s')
        self.assertEqual(len(first[5]), 200)
        self.assertEqual(first[5][0], [1, '1', '2'])
        self.assertEqual(last[5], [[8201, '8201', '16402']])
        self.assertEqual(ws['B8201'].value, '=A8201*2')
        # Dimensions are cached, not repeatedly scanned per row or page.
        self.assertEqual(app.sheet_dimensions['PR'], (2, 8201))

    def test_revisiting_page_uses_cache(self):
        wb = Workbook(); wb.active.title = 'PR'
        app = self.app(wb)
        app.current = {'tables': {'PR': None}}
        msg = ui.App.job_page(app, 'PR', 0)
        app.page_cache[('PR', 0)] = msg
        ui.App.show_sheet(app, 'PR')
        app.present_page.assert_called_once_with(msg)
        app.run_bg.assert_not_called()

    def test_report_switch_does_not_rebuild(self):
        wb = Workbook(); wb.active.title = 'PR'
        app = self.app(wb)
        app.ensure_built.side_effect = lambda **kwargs: ui.App.ensure_built(app, **kwargs)
        with patch.object(ui.engine, 'build_workbook') as build:
            self.assertEqual(ui.App.job_report(app, 'consolidate'), ('sheets', 'consolidate', {'PR': None}))
            ui.App.job_report(app, 'summary')
            ui.App.job_report(app, 'consolidate')
            build.assert_not_called()

    def test_option_change_invalidates_cached_pages(self):
        wb = Workbook(); app = self.app(wb)
        app.page_cache[('PR', 0)] = 'old'
        app.job_settings = ('company', 6, False)
        with patch.object(ui.engine, 'build_workbook') as build, patch.object(ui.engine, 'find_dir', return_value=None):
            ui.App.ensure_built(app)
            build.assert_not_called()
        self.assertFalse(app.page_cache)
        self.assertEqual(app.built[0], app.job_settings)

    def test_maximize_keeps_normal_window(self):
        app = Mock()
        ui.App.maximize_window(app)
        app.state.assert_called_once_with('zoomed')
        app.attributes.assert_not_called()

    def test_inventory_automatically_on_folder_selection(self):
        app = Mock(); app.busy = False
        with patch.object(ui.filedialog, 'askdirectory', return_value='/company'):
            ui.App.pick_folder(app)
        app.folder.set.assert_called_once_with('/company')
        app.run_bg.assert_called_once_with(app.job_prepare_sources)

    def test_busy_export_does_not_start_another_worker(self):
        app = Mock(); app.busy = True
        ui.App.export(app, None, 'report.xlsx')
        app.run_bg.assert_not_called()


if __name__ == '__main__':
    unittest.main()
