import sys
import tempfile
import unittest
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook, Workbook
from input_tools import convert_pdf, source_inventory, validate_sources
import control_center as ui
from unittest.mock import Mock


def write_pdf(path, blank=False, ruled=True):
    """Small real PDF fixture: ruled table with text, dates and literal formula-like text."""
    text = b'' if blank else b'''0.5 w
50 700 m 450 700 l S 50 680 m 450 680 l S 50 660 m 450 660 l S 50 640 m 450 640 l S
50 700 m 50 640 l S 250 700 m 250 640 l S 450 700 m 450 640 l S
BT /F1 10 Tf 60 685 Td (Date) Tj ET
BT /F1 10 Tf 260 685 Td (Amount) Tj ET
BT /F1 10 Tf 60 665 Td (2026-01-01) Tj ET
BT /F1 10 Tf 260 665 Td (1,000.00) Tj ET
BT /F1 10 Tf 60 645 Td (000123) Tj ET
BT /F1 10 Tf 260 645 Td (=1+1) Tj ET
'''
    if not blank and not ruled:
        text = b'BT ' + text.split(b'BT ', 1)[1]
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>',
               b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
               b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 800] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
               b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
               b'<< /Length '+str(len(text)).encode()+b' >>\nstream\n'+text+b'\nendstream']
    data = b'%PDF-1.4\n'; offsets = [0]
    for index, obj in enumerate(objects, 1):
        offsets.append(len(data)); data += str(index).encode()+b' 0 obj\n'+obj+b'\nendobj\n'
    start=len(data); data += b'xref\n0 6\n0000000000 65535 f \n'
    for offset in offsets[1:]: data += f'{offset:010d} 00000 n \n'.encode()
    data += b'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n'+str(start).encode()+b'\n%%EOF\n'
    path.write_bytes(data)


class InputToolsTests(unittest.TestCase):
    def test_summary_counts_formats_without_filenames(self):
        with tempfile.TemporaryDirectory() as folder:
            gl=Path(folder)/'GL'; gl.mkdir()
            for n in range(6): (gl/f'{n}.xlsx').touch()
            for n in range(3): (gl/f'{n}.pdf').touch()
            (gl/'~$lock.xlsx').touch()
            row = next(r for r in source_inventory(folder) if r[0]=='General ledger')
            self.assertEqual(row[2],9)
            self.assertEqual(row[4],'6 Excel, 3 PDF')
            self.assertIn('Not ready',row[3])
            self.assertNotIn('.xlsx',str(row))

    def test_tax_pdf_is_counted_and_flagged(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder); year=root/'PR'/'2025'; year.mkdir(parents=True)
            (year/'return.pdf').touch()
            row=next(r for r in source_inventory(root) if r[:2]==['PR','2025'])
            self.assertEqual(row[2],1)
            self.assertIn('need conversion',row[3])
            self.assertTrue(any(r[2]=='return.pdf' and 'NOT READY' in r[4] for r in validate_sources(root)))

    def test_text_pdf_extracts_table_and_preserves_literal_values(self):
        with tempfile.TemporaryDirectory() as folder:
            bank=Path(folder)/'Bank'; bank.mkdir()
            source=bank/'bank.pdf'; output=bank/'bank.xlsx'
            write_pdf(source); original=source.read_bytes()
            result=convert_pdf(source,output)
            self.assertEqual(result['tables'],1)
            wb=load_workbook(output)
            ws=wb['P1 Table 1']
            self.assertEqual(ws['A1'].value,'Date')
            self.assertEqual(ws['B2'].value,'1,000.00')
            self.assertEqual(ws['A3'].value,'000123')
            self.assertEqual(ws['B3'].value,'=1+1')
            self.assertEqual(ws['B3'].data_type,'s')
            self.assertIn('Conversion Review',wb.sheetnames)
            self.assertEqual(source.read_bytes(),original)
            self.assertTrue(any('REVIEW: PDF extraction' in r[4] for r in validate_sources(folder)))

    def test_text_only_fallback_and_scanned_pdf(self):
        from PIL import Image
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'text.pdf'; output=Path(folder)/'text.xlsx'
            write_pdf(source,ruled=False)
            result=convert_pdf(source,output)
            self.assertEqual(result['tables'],0)
            self.assertEqual(result['text_pages'],1)
            wb=load_workbook(output)
            self.assertIn('P1 Text',wb.sheetnames)
            scan=Path(folder)/'scan.pdf'
            Image.new('RGB',(50,50),'white').save(scan,'PDF')
            with self.assertRaisesRegex(ValueError,'OCR/manual review'):
                convert_pdf(scan,Path(folder)/'scan.xlsx')
            self.assertFalse((Path(folder)/'scan.xlsx').exists())

    def test_no_overwrite_or_blank_pdf_output(self):
        with tempfile.TemporaryDirectory() as folder:
            source=Path(folder)/'blank.pdf';output=Path(folder)/'existing.xlsx'
            write_pdf(source,blank=True); output.write_bytes(b'preserve')
            with self.assertRaises(FileExistsError): convert_pdf(source,output)
            self.assertEqual(output.read_bytes(),b'preserve')
            output.unlink()
            with self.assertRaisesRegex(ValueError,'OCR/manual review'): convert_pdf(source,output)
            self.assertFalse(output.exists())

    def test_search_reaches_rows_beyond_preview_and_other_tax_sheets(self):
        wb=Workbook();wb.active.title='PR';wb['PR']['A6001']='Target supplier'
        wb.create_sheet('SR')['B3001']='target SUPPLIER'
        app=Mock();app.built=(None,wb,[],None)
        result=ui.App.job_search(app,'target supplier',['PR','SR'])
        self.assertEqual(result,('search',[('PR',6001),('SR',3001)]))
        app.busy=False;app.search_hits=result[1];app.search_index=-1;app.page_size=200;app.sheet_pages={}
        ui.App.next_match(app)
        self.assertEqual(app.sheet_pages['PR'],30)
        self.assertEqual(app.search_target,('PR',6001))
        app.show_sheet.assert_called_once_with('PR')

    def test_grouped_tabs_and_no_company_tab(self):
        import engine
        wb=Workbook();wb.remove(wb.active)
        for name in ('Summary MTR','MTR_PR','MTR_SR','MTR_TOS','MTR_WHT','TOI','FULL TOI','Tax Paid Summary','Company Info'):
            wb.create_sheet(name)
        app=Mock();app.ensure_built.return_value=(None,wb,[],None)
        self.assertEqual(list(ui.App.job_report(app,'mtr')[2]), ['MTR_PR','MTR_SR','MTR_TOS','MTR_WHT','Summary MTR'])
        self.assertEqual(list(ui.App.job_report(app,'tois')[2]),['TOI','FULL TOI'])
        self.assertEqual(list(ui.App.job_report(app,'tps')[2]),['Tax Paid Summary'])
        self.assertNotIn('Company Info',engine.SECTIONS['tps'][1])
        self.assertFalse(any(key in ('sources','build_all','summary','toi','fulltoi') for _,key,_ in ui.MENU))


if __name__=='__main__': unittest.main()
