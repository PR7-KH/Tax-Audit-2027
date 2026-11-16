import sys
import tempfile
import types
import unittest
from contextlib import contextmanager
from pathlib import Path
from unittest.mock import Mock, MagicMock, patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font
import toi_pdf
from input_tools import classify_download,organize_downloads,file_hash


class ExcelPdfTests(unittest.TestCase):
    def source(self,path):
        wb=Workbook();ws=wb.active;ws.title='step1'
        ws['A1']='ក្រុមហ៊ុន';ws['A1'].font=Font(name='Khmer OS Siemreap',size=12)
        ws['A2']='Total';ws['B2']='=1+2';wb.save(path)

    def test_toi_year_suffixes_and_real_filename_routing(self):
        for name in ('L001-100121101_TOI25.xlsx','TOI26.xlsx','TOI2025.xlsx','toi_2026.xlsx','TOI.xlsx'):
            self.assertEqual(classify_download(name)[:2],('TOI',None),name)
        self.assertIsNone(classify_download('STOIC25.xlsx')[0])
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'Download').mkdir();(root/'Info').mkdir()
            name='L001-100121101_TOI25.xlsx';self.source(root/'Download'/name)
            before=file_hash(root/'Download'/name)
            report=organize_downloads(root)
            self.assertEqual(report[0][2],'MOVED')
            self.assertEqual(file_hash(root/'TOI'/name),before)

    def test_renderer_defaults_and_explicit_alternative(self):
        for platform in ('win32','darwin'):
            with patch.object(toi_pdf.sys,'platform',platform):
                self.assertEqual(toi_pdf.select_renderer(),'excel')
                self.assertEqual(toi_pdf.select_renderer('libreoffice'),'libreoffice')
        with patch.object(toi_pdf.sys,'platform','linux'):
            self.assertEqual(toi_pdf.select_renderer(),'libreoffice')
        with self.assertRaises(ValueError):toi_pdf.select_renderer('unknown')

    def test_windows_worker_initializes_com_and_owns_separate_app(self):
        app=Mock();factory=Mock(return_value=app);com=Mock()
        with patch.object(toi_pdf.sys,'platform','win32'), patch.dict(sys.modules,{'xlwings':types.SimpleNamespace(App=factory),'pythoncom':com}):
            with toi_pdf.excel_application() as running:
                self.assertIs(running,app)
                com.CoInitialize.assert_called_once()
                self.assertFalse(app.display_alerts)
            com.CoUninitialize.assert_called_once();app.quit.assert_called_once()
            factory.assert_called_once_with(visible=False,add_book=False)

    def test_mac_app_cleanup_and_launch_failure_does_not_use_libreoffice(self):
        app=Mock()
        with patch.object(toi_pdf.sys,'platform','darwin'),patch.dict(sys.modules,{'xlwings':types.SimpleNamespace(App=Mock(return_value=app))}):
            with self.assertRaisesRegex(RuntimeError,'export failure'):
                with toi_pdf.excel_application():raise RuntimeError('export failure')
            app.quit.assert_called_once()
        with patch.object(toi_pdf.sys,'platform','darwin'),patch.dict(sys.modules,{'xlwings':types.SimpleNamespace(App=Mock(side_effect=OSError('Excel not installed')))}),patch.object(toi_pdf,'find_soffice') as fallback:
            with self.assertRaisesRegex(RuntimeError,'Cannot start Microsoft Excel'):
                with toi_pdf.excel_application():pass
            fallback.assert_not_called()

    def test_cleanup_error_keeps_the_original_export_error(self):
        app=Mock();app.quit.side_effect=RuntimeError('cleanup denied')
        with patch.object(toi_pdf.sys,'platform','darwin'),patch.dict(sys.modules,{'xlwings':types.SimpleNamespace(App=Mock(return_value=app))}):
            with self.assertWarns(RuntimeWarning),self.assertRaisesRegex(RuntimeError,'original export error'):
                with toi_pdf.excel_application():raise RuntimeError('original export error')

    def test_native_renderer_opens_only_temp_copy_and_closes_on_error(self):
        app=Mock();book=MagicMock();app.books.open.return_value=book
        book.to_pdf.side_effect=RuntimeError('Native PDF export failed')
        with self.assertRaisesRegex(RuntimeError,'Native PDF export failed'):
            toi_pdf.render_excel(app,Path('/tmp/prepared.xlsx'),Path('/tmp/output.pdf'),['step1'])
        app.books.open.assert_called_once_with('/tmp/prepared.xlsx',update_links=False,read_only=True)
        book.to_pdf.assert_called_once_with(path='/tmp/output.pdf',include=['step1'],quality='standard',show=False)
        book.close.assert_called_once();app.calculate.assert_called_once()

    def test_native_batch_reuses_app_preserves_sources_and_retains_fonts(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'TOI').mkdir()
            for name in ('TOI25.xlsx','TOI26.xlsx'):self.source(root/'TOI'/name)
            before={p.name:p.read_bytes() for p in (root/'TOI').glob('*.xlsx')}
            app=Mock();book=MagicMock();app.books.open.return_value=book
            def export(**options):Path(options['path']).write_bytes(b'%PDF-1.4\nmock native API response')
            book.to_pdf.side_effect=export
            @contextmanager
            def session():yield app
            with patch.object(toi_pdf,'excel_application',side_effect=session) as opened,patch.object(toi_pdf,'find_soffice') as lo:
                result=toi_pdf.export_toi_pdfs(root,'1-1',renderer='excel')
                opened.assert_called_once();lo.assert_not_called()
            self.assertTrue(all(row[2]=='EXPORTED' for row in result),result)
            self.assertEqual(book.close.call_count,2)
            self.assertTrue(all(Path(row[3]).is_file() for row in result))
            self.assertEqual(before,{p.name:p.read_bytes() for p in (root/'TOI').glob('*.xlsx')})
            prepared=root/'prepared.xlsx';toi_pdf.prepare_workbook(root/'TOI'/'TOI25.xlsx',prepared,1,1)
            wb=load_workbook(prepared)
            self.assertEqual(wb['step1']['A1'].font.name,'Khmer OS Siemreap')
            self.assertEqual(wb['step1']['A1'].value,'ក្រុមហ៊ុន')
            self.assertEqual(wb['step1']['B2'].value,'=1+2')
            self.assertTrue(all('Khmer OS Siemreap' in row[1] for row in toi_pdf.toi_workbook_fonts(root)))


if __name__=='__main__':unittest.main()
