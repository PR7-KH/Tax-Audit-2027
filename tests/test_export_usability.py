import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook,load_workbook
import control_center as ui
import toi_pdf
import setup_windows
from input_tools import classify_download,organize_downloads
import test_import_pdf_export as fixtures


class ExportUsabilityTests(unittest.TestCase):
    def test_future_toi_names_route_without_false_substring_matches(self):
        for name in ('L001-100121101_TOI25.xlsx','TOI2X.xlsx','TOI30.xlsx','TOI2099.xlsx','TOI2500.xlsx','TOIX.xlsx'):
            self.assertEqual(classify_download(name)[:2],('TOI',None),name)
        for name in ('STOIC.xlsx','TOILET.xlsx'):
            self.assertIsNone(classify_download(name)[0])

    def test_windows_runtime_rejects_the_uploaded_failure_configuration(self):
        self.assertFalse(setup_windows.compatible_runtime({'version':[3,15],'platform':'win-arm64'}))
        self.assertFalse(setup_windows.compatible_runtime({'version':[3,12],'platform':'win-arm64'}))
        self.assertTrue(setup_windows.compatible_runtime({'version':[3,12],'platform':'win-amd64'}))
        with patch.object(setup_windows,'runtime_info',return_value={'version':[3,15],'platform':'win-arm64'}),patch.object(setup_windows.subprocess,'run',return_value=Mock(stdout='')):
            with self.assertRaisesRegex(RuntimeError,'Python 3.12 Windows x86-64'):
                setup_windows.find_runtime()

    def test_windows_setup_preserves_mac_and_backs_up_wrong_runtime(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            mac=root/'.venv';mac.mkdir();(mac/'keep.txt').write_text('Mac environment')
            old=root/'.venv-win'/'Scripts';old.mkdir(parents=True);(old/'python.exe').write_bytes(b'old ARM runtime')
            commands=[]
            def run(command,**kwargs):
                commands.append(command)
                if command[1:3]==['-m','venv']:
                    target=Path(command[3])/'Scripts';target.mkdir(parents=True);(target/'python.exe').write_bytes(b'new x64 runtime')
                return Mock(returncode=0)
            with patch.object(setup_windows,'ROOT',root),patch.object(setup_windows.sys,'platform','win32'),patch.object(setup_windows,'find_runtime',return_value=Path('/fake/python312.exe')),patch.object(setup_windows,'runtime_info',return_value={'version':[3,15],'platform':'win-arm64'}),patch.object(setup_windows.subprocess,'run',side_effect=run):
                setup_windows.main()
            backup=list(root.glob('.venv-win.backup-*'))
            self.assertEqual(len(backup),1)
            self.assertEqual((backup[0]/'Scripts'/'python.exe').read_bytes(),b'old ARM runtime')
            self.assertEqual((mac/'keep.txt').read_text(),'Mac environment')
            self.assertEqual((root/'.venv-win'/'Scripts'/'python.exe').read_bytes(),b'new x64 runtime')
            self.assertTrue(any('--only-binary=:all:' in c for c in commands))

    def test_windows_setup_missing_runtime_does_not_change_existing_files(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'keep.txt').write_text('keep')
            with patch.object(setup_windows,'ROOT',root),patch.object(setup_windows.sys,'platform','win32'),patch.object(setup_windows,'find_runtime',side_effect=RuntimeError('Install Python 3.12 x64')):
                with self.assertRaisesRegex(RuntimeError,'Install Python 3.12 x64'):setup_windows.main()
            self.assertEqual([p.name for p in root.iterdir()],['keep.txt'])

    def test_mac_workspace_reuses_prepared_file_and_preserves_original(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);source=root/'TOI25.xlsx';fixtures.ToiPdfTests().make_book(source)
            original=source.read_bytes();dest=root/'TOI_PDFs';dest.mkdir()
            with patch.object(toi_pdf.sys,'platform','darwin'):
                with toi_pdf.export_workspace(dest,source,'excel') as (prepared,pdf,_):
                    toi_pdf.prepare_workbook(source,prepared,1,1)
                    first=prepared;inode=prepared.stat().st_ino
                    pdf.write_bytes(b'%PDF-previous-work-file')
                with toi_pdf.export_workspace(dest,source,'excel') as (prepared,pdf,_):
                    self.assertEqual(prepared,first)
                    self.assertFalse(pdf.exists()) # stale stage cannot be published as a new export
                    toi_pdf.prepare_workbook(source,prepared,4,4)
                    self.assertEqual(prepared.stat().st_ino,inode)
                self.assertTrue(first.is_file())
            self.assertEqual(source.read_bytes(),original)

    def test_mac_workspace_protects_unowned_folder_and_concurrent_export(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);source=root/'TOI25.xlsx';fixtures.ToiPdfTests().make_book(source)
            dest=root/'TOI_PDFs';dest.mkdir();private=dest/'.excel-work';private.mkdir()
            (private/'keep.txt').write_text('user file')
            with patch.object(toi_pdf.sys,'platform','darwin'):
                with self.assertRaisesRegex(RuntimeError,'not a recognized app workspace'):
                    with toi_pdf.export_workspace(dest,source,'excel'):pass
                self.assertEqual((private/'keep.txt').read_text(),'user file')
                (private/'keep.txt').unlink();private.rmdir()
                with toi_pdf.export_workspace(dest,source,'excel'):
                    with self.assertRaisesRegex(RuntimeError,'Another TOI export'):
                        with toi_pdf.export_workspace(dest,source,'excel'):pass
                with toi_pdf.export_workspace(dest,source,'excel'):pass # lock released

    def test_mtr_file_link_decodes_spaces_and_unicode(self):
        with tempfile.TemporaryDirectory() as folder:
            path=Path(folder)/'Purchase ក្រុមហ៊ុន 01.xlsx';path.write_bytes(b'raw file')
            wb=Workbook();wb.active.title='MTR_PR';wb['MTR_PR']['B5']='Purchase file'
            wb['MTR_PR']['B5'].hyperlink=path.as_uri()
            self.assertEqual(ui.mtr_source_path(wb,'MTR_PR',5),path)
            self.assertIsNone(ui.mtr_source_path(wb,'MTR_PR',6))
            self.assertIsNone(ui.mtr_source_path(wb,'Summary MTR',5))
            wb['MTR_PR']['B5'].hyperlink='https://example.com'
            self.assertIsNone(ui.mtr_source_path(wb,'MTR_PR',5))

    def test_mtr_preview_click_opens_linked_file_only(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.xlsx';source.touch()
            wb=Workbook();wb.active.title='MTR_PR';wb['MTR_PR']['B5']='source.xlsx';wb['MTR_PR']['B5'].hyperlink=source.as_uri()
            app=Mock();app.busy=False;app.built=(None,wb,[],None);app.current={'sheet':'MTR_PR'}
            app.tree.identify_column.return_value='#3';app.tree.identify_row.return_value='row'
            app.tree.item.return_value=('5','1','source.xlsx','100')
            app.mtr_link_at.side_effect=lambda event:ui.App.mtr_link_at(app,event)
            event=Mock(x=100,y=150)
            with patch.object(ui,'open_path') as opened:
                ui.App.open_mtr_source(app,event);opened.assert_called_once_with(source)
            app.tree.identify_column.return_value='#4'
            with patch.object(ui,'open_path') as opened:
                ui.App.open_mtr_source(app,event);opened.assert_not_called()

    def test_mtr_export_link_uses_actual_folder_case(self):
        import engine
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);year=root/'pr'/'2025';year.mkdir(parents=True)
            source=year/'Purchase 01.xlsx';source.touch()
            wb=Workbook();ws=wb.active;ws.title='PR';ws.cell(4,1,source.name);ws.cell(4,2,2025);ws.cell(4,7,100)
            idx=wb.create_sheet(engine.cons.MTR_INDEX);idx.append(['Tax','Row']);idx.append(['PR',4])
            engine.build_mtr_detail(wb,root,engine.Log())
            link=wb['MTR_PR']['B5'].hyperlink
            self.assertIsNotNone(link)
            self.assertEqual(link.target,source.as_uri())


if __name__=='__main__':unittest.main()
