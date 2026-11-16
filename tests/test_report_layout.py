import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import engine
import control_center as ui


class ReportLayoutTests(unittest.TestCase):
    def info(self, reverse=False):
        wb = Workbook(); wb.active.title = 'Sheet1'
        wb.active['A1'] = 'Company'; wb.active['B1'] = 'Example'
        for title, year in [('Sheet2', 2026 if reverse else 2025), ('Sheet3', 2025 if reverse else 2026)]:
            ws = wb.create_sheet(title); ws['B3'] = year
            ws['A4'] = 1; ws['B4'] = 'VAT'
            for c in range(3, 15): ws.cell(4, c, 10 if year == 2025 else 20)
        return wb

    def test_tax_paid_chronological_years_and_month_selection(self):
        for reverse in (False, True):
            for months in (1, 3, 12):
                with self.subTest(reverse=reverse, months=months):
                    wb = Workbook()
                    engine.tax_paid_summary(wb, self.info(reverse), months, engine.Log())
                    ws = wb['Tax Paid Summary']
                    self.assertEqual(ws['C9'].value, 'Year 2025')
                    self.assertEqual(ws['D9'].value, f'2026 / {months} months')
                    self.assertEqual(ws['C10'].value, 120)
                    self.assertEqual(ws['D10'].value, 20 * months)
                    self.assertEqual(ws['E10'].value, 120 + 20 * months)
                    engine.tax_paid_summary(wb, self.info(reverse), 6, engine.Log())
                    self.assertEqual(wb['Tax Paid Summary']['D10'].value, 120)
                    self.assertNotIn('_Tax Paid Summary', wb.sheetnames)

    def test_years_not_guessed(self):
        for year in (None, 'unknown', 2025):
            info = self.info(); info['Sheet3']['B3'] = year
            with self.assertRaises(ValueError):
                engine.tax_paid_summary(Workbook(), info, 3, engine.Log())

    def test_compact_export_preserves_transactions_and_mtr_detail(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)/'company'
            for tax, layout in engine.cons.LAYOUTS.items():
                for year, value in [(2025, 10), (2026, 20)]:
                    directory = root/tax/str(year); directory.mkdir(parents=True)
                    for month in (1, 2):
                        wb = Workbook(); ws = wb.active
                        ws.cell(2, layout.first_col, layout.first_heading)
                        ws.cell(layout.first_row, 1, f'{year} transaction {month}')
                        for c in range(layout.first_col, layout.last_col+1): ws.cell(layout.first_row, c, value)
                        wb.save(directory/f'{month:02}.xlsx')
            output = Path(folder)/'audit.xlsx'
            engine.build(root, output)
            wb = load_workbook(output)
            for tax, layout in engine.cons.LAYOUTS.items():
                ws = wb[tax]; c = layout.first_col
                self.assertEqual([ws.cell(r,1).value for r in (1,2,3)], ['Grand Total','Year 2025','Year 2026'])
                self.assertEqual([ws.cell(r,c).value for r in (1,2,3)], [60,20,40])
                self.assertEqual(ws.max_row, 9)  # 3 totals + 2 headings + 4 transactions
                self.assertEqual(ws.cell(6,1).value, '2025 transaction 1')
                self.assertEqual(ws.cell(9,1).value, '2026 transaction 2')
                self.assertEqual(ws.freeze_panes, 'A6')
                mtr = wb[f'MTR_{tax}']
                names = [cell.value for row in mtr.iter_rows() for cell in row]
                self.assertEqual(names.count('01.xlsx'), 2)
                self.assertIn('GRAND TOTAL', names)
                self.assertTrue(any(cell.data_type == 'f' for row in mtr.iter_rows() for cell in row))

    def test_month_update_reuses_tax_workbook(self):
        import calc
        from collections import OrderedDict
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as folder:
            info_dir = Path(folder)/'Info'; info_dir.mkdir()
            self.info().save(info_dir/'info.xlsx')
            wb = Workbook(); wb.active.title='PR'; wb['PR']['A1']='Keep transaction'
            engine.tax_paid_summary(wb, self.info(), 3, engine.Log())
            app=Mock(); app.built=((folder,3,False),wb,engine.Log(),calc.Calc(wb))
            app.job_settings=(folder,6,False)
            app.page_cache=OrderedDict(); app.sheet_pages={}; app.sheet_dimensions={}
            with patch.object(ui.engine, 'build_workbook') as rebuild:
                ui.App.ensure_built(app)
                rebuild.assert_not_called()
            self.assertIs(app.built[1],wb)
            self.assertEqual(wb['PR']['A1'].value,'Keep transaction')
            self.assertEqual(wb['Tax Paid Summary']['D10'].value,120)

    def test_summary_can_open_without_tax_consolidation(self):
        from collections import OrderedDict
        from unittest.mock import patch
        with tempfile.TemporaryDirectory() as folder:
            info_dir = Path(folder)/'Info'; info_dir.mkdir()
            self.info().save(info_dir/'info.xlsx')
            app=Mock(); app.built=None; app.partial_build=False
            app.job_settings=(folder,3,False)
            app.page_cache=OrderedDict(); app.sheet_pages={}; app.sheet_dimensions={}
            with patch.object(ui.engine,'build_workbook') as build:
                ui.App.ensure_built(app, summary_only=True)
                build.assert_not_called()
            self.assertTrue(app.partial_build)
            self.assertEqual(app.built[1]['Tax Paid Summary']['D10'].value,60)
            full=Workbook()
            with patch.object(ui.engine,'build_workbook',return_value=(full,[])) as build:
                ui.App.ensure_built(app)
                build.assert_called_once()
            self.assertFalse(app.partial_build)

    def test_font_cache_preserves_previous_formatting(self):
        wb = Workbook(); ws=wb.active
        ws['A1']='Title'; ws['A1'].font=Font(bold=True, italic=True, color='FF123456')
        ws['A2']=10
        engine.fmt_all(ws)
        self.assertEqual(ws['A1'].font.name, engine.FONT)
        self.assertTrue(ws['A1'].font.bold)
        self.assertTrue(ws['A1'].font.italic)
        self.assertEqual(ws['A1'].font.color.rgb, 'FF123456')
        self.assertFalse(ws['A2'].font.bold)

    def test_fit_and_zoom_redraw_non_sheet_views(self):
        app=Mock(); app.fit_columns=False
        ui.App.toggle_fit(app)
        self.assertTrue(app.fit_columns)
        app.redraw_table.assert_called_once()
        app.last_table=(['File'], [['statement.pdf']], True, None)
        ui.App.redraw_table(app)
        app.draw.assert_called_once_with(['File'], [['statement.pdf']], True, tags=None)

    def test_source_file_names_and_types(self):
        with tempfile.TemporaryDirectory() as folder:
            directory=Path(folder)/'GL'; directory.mkdir()
            for name in ('ledger.xlsx','bank.pdf','sales.csv','~$lock.xlsx'):
                (directory/name).touch()
            app=Mock(); app.job_settings=(folder,3,False)
            msg=ui.App.job_source_files(app)
            self.assertEqual({r[3] for r in msg[2]}, {'Excel','PDF','CSV'})
            self.assertEqual(len(msg[2]),3)


if __name__ == '__main__': unittest.main()
