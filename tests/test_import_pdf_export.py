import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook,load_workbook
import pdfplumber
import control_center as ui
from input_tools import organize_downloads,classify_download,file_hash
from toi_pdf import prepare_workbook,export_toi_pdfs,parse_steps,find_soffice,STEP_LAYOUTS


class DownloadSetupTests(unittest.TestCase):
    def test_download_routing_and_journal(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);download=root/'Download';download.mkdir()
            names={'L001_Purchase_01_2025.xlsx':'PR/2025',
                   'L001_Sale_01_2026.xlsx':'SR/2026',
                   'L001_Salary_Tax_03_2026.xlsx':'TOS/2026',
                   'L001_WITHHOLDING_TAX_03_2025.xlsx':'WHT/2025',
                   'GL_2026.xlsx':'GL', 'TB_2026.xlsx':'TB','Bank_statement.pdf':'BANK',
                   'TOI_2025_2026.xlsx':'TOI','company_Info.xlsx':'Info'}
            for name in names:(download/name).write_bytes(('contents '+name).encode())
            (download/'unknown.xlsx').write_bytes(b'keep')
            (download/'Purchase_2025_2026.xlsx').write_bytes(b'ambiguous')
            results=organize_downloads(root)
            self.assertEqual(sum(r[2]=='MOVED' for r in results),len(names))
            for name,destination in names.items():
                self.assertEqual((root/destination/name).read_bytes(),('contents '+name).encode())
                self.assertFalse((download/name).exists())
            self.assertEqual((download/'unknown.xlsx').read_bytes(),b'keep')
            self.assertTrue((download/'Purchase_2025_2026.xlsx').exists())
            self.assertTrue((root/'.tax-audit-imports.jsonl').is_file())
            self.assertIn('Automatic setup skipped',organize_downloads(root)[0][2])

    def test_existing_targets_mean_no_changes(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'Bank').mkdir();(root/'Download').mkdir()
            source=root/'Download'/'Purchase_2025.xlsx';source.write_bytes(b'untouched')
            self.assertIn('Automatic setup skipped',organize_downloads(root)[0][2])
            self.assertTrue(source.exists());self.assertFalse((root/'PR').exists())

    def test_collisions_ambiguity_and_missing_year_stay_in_download(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);download=root/'Download';download.mkdir()
            for name in ('a','b'):
                directory=download/name;directory.mkdir();(directory/'Purchase_2025.xlsx').write_bytes(name.encode())
            (download/'Sale.xlsx').write_bytes(b'no year')
            (download/'Purchase_GL_2025.xlsx').write_bytes(b'ambiguous')
            results=organize_downloads(root)
            self.assertEqual(sum(r[2]=='MOVED' for r in results),1)
            self.assertEqual((root/'PR'/'2025'/'Purchase_2025.xlsx').read_bytes(),b'a')
            self.assertEqual((download/'b'/'Purchase_2025.xlsx').read_bytes(),b'b')
            self.assertTrue((download/'Sale.xlsx').exists())
            self.assertTrue((download/'Purchase_GL_2025.xlsx').exists())

    def test_info_and_information_do_not_block_initial_setup(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            for name in ('Info','Information','TOI','Download'):(root/name).mkdir()
            (root/'Info'/'keep.xlsx').write_bytes(b'existing info')
            source=root/'Download'/'Purchase_01_2025.xlsx';source.write_bytes(b'purchase')
            report=organize_downloads(root)
            self.assertEqual(sum(r[2]=='MOVED' for r in report),1)
            self.assertEqual((root/'PR'/'2025'/source.name).read_bytes(),b'purchase')
            self.assertEqual((root/'Info'/'keep.xlsx').read_bytes(),b'existing info')

    def test_explicit_import_handles_existing_targets_and_preserves_conflicts(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'PR'/'2025').mkdir(parents=True);(root/'Download').mkdir()
            name='Purchase_01_2025.xlsx'
            (root/'PR'/'2025'/name).write_bytes(b'old')
            (root/'Download'/name).write_bytes(b'new')
            (root/'Download'/'Sale_2026.xlsx').write_bytes(b'sales')
            report=organize_downloads(root,allow_existing=True)
            self.assertTrue(any('destination already exists' in r[2] for r in report))
            self.assertEqual((root/'PR'/'2025'/name).read_bytes(),b'old')
            self.assertEqual((root/'Download'/name).read_bytes(),b'new')
            self.assertEqual((root/'SR'/'2026'/'Sale_2026.xlsx').read_bytes(),b'sales')

    def test_no_download_and_empty_download_explain_no_work(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            self.assertIn('No Download child',organize_downloads(root)[0][2])
            (root/'Download').mkdir()
            self.assertIn('no visible files',organize_downloads(root)[0][2])
            self.assertIn('not Download itself',organize_downloads(root/'Download')[0][2])

    def test_explicit_financial_names(self):
        self.assertEqual(classify_download('purchase ledger.xlsx')[:2],('EXPENSE',None))
        self.assertEqual(classify_download('salary_tax_2026.xlsx')[:2],('TOS','2026'))
        self.assertIsNone(classify_download('supplier_123.xlsx')[0])


class ToiPdfTests(unittest.TestCase):
    def make_book(self,path):
        wb=Workbook();wb.active.title='Controls';wb.active['A1']='MUST NOT PRINT'
        for n in (1,4,15):
            ws=wb.create_sheet(f'step{n}');ws['A1']=f'TOI STEP {n}';ws['B2']='=1+2';ws['A3']='End'
            ws['E1']='OUTSIDE STEP ONE PRINT AREA'
        wb.save(path)

    def test_vba_layouts_and_hidden_dependencies(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.xlsx';prepared=Path(folder)/'prepared.xlsx';self.make_book(source)
            before=source.read_bytes();selected=prepare_workbook(source,prepared,4,15)
            self.assertEqual(selected,['step4','step15'])
            wb=load_workbook(prepared)
            self.assertEqual(wb['step1'].sheet_state,'hidden')
            self.assertEqual(wb['Controls'].sheet_state,'hidden')
            self.assertEqual(wb['step4'].page_setup.orientation,'portrait')
            self.assertEqual(wb['step4'].page_setup.fitToHeight,2)
            self.assertEqual(wb['step15'].page_setup.fitToHeight,6)
            self.assertEqual(wb['step4'].page_margins.left,.5)
            self.assertEqual(wb['step4'].row_dimensions[1].height,22)
            self.assertAlmostEqual(wb['step4'].column_dimensions['A'].width,10/.2135)
            self.assertIn('$D$3',str(wb['step4'].print_area))
            self.assertEqual(wb['step4']['B2'].value,'=1+2')
            self.assertEqual(source.read_bytes(),before)
            self.assertEqual(len(STEP_LAYOUTS),17)

    def test_ranges_and_missing_steps(self):
        for invalid in ('1','0-2','1-18','4-1','x-y'):
            with self.assertRaises(ValueError):parse_steps(invalid)
        self.assertEqual(parse_steps(' 4 - 4 '),(4,4))
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'source.xlsx';self.make_book(source)
            with self.assertRaises(ValueError):prepare_workbook(source,Path(folder)/'out.xlsx',2,3)

    def test_real_pdf_export_and_no_source_changes(self):
        try: find_soffice()
        except RuntimeError:self.skipTest('LibreOffice not installed')
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);toi=root/'TOI';toi.mkdir();source=toi/'company.xlsx';self.make_book(source)
            before=file_hash(source)
            results=export_toi_pdfs(root,'1-4')
            self.assertEqual(results[0][2],'EXPORTED',results)
            with pdfplumber.open(results[0][3]) as pdf:
                self.assertEqual(len(pdf.pages),2)
                text=' '.join(p.extract_text() or '' for p in pdf.pages)
                self.assertIn('TOI STEP 1',text);self.assertIn('TOI STEP 4',text)
                self.assertNotIn('MUST NOT PRINT',text)
                self.assertNotIn('OUTSIDE STEP ONE PRINT AREA',pdf.pages[0].extract_text() or '')
                self.assertGreater(pdf.pages[0].width,pdf.pages[0].height)
                self.assertLess(pdf.pages[1].width,pdf.pages[1].height)
            old=Path(results[0][3]).read_bytes()
            again=export_toi_pdfs(root,'1-4')
            self.assertEqual(again[0][2],'EXPORTED')
            self.assertNotEqual(again[0][3],results[0][3])
            self.assertEqual(Path(results[0][3]).read_bytes(),old)
            self.assertEqual(file_hash(source),before)

    def test_dashboard_preview_hides_toi_dependency(self):
        wb=Workbook();wb.remove(wb.active)
        for name in ('Audit Dashboard','ADT','TOI','Company Info'):wb.create_sheet(name)
        app=Mock();app.ensure_built.return_value=(None,wb,[],None)
        self.assertNotIn('TOI',ui.App.job_report(app,'dashboard')[2])
        self.assertIn('TOI',ui.engine.SECTIONS['dashboard'][1])  # export formulas retain dependency


if __name__=='__main__':unittest.main()
