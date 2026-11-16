import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock,patch
from zipfile import ZipFile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,Border,Side
from docx import Document
from docx.oxml.ns import qn
import audit_documents as docs
import control_center as ui
import engine
from toi_pdf import export_toi_pdfs,find_soffice,is_toi_output
from input_tools import source_inventory,validate_sources
import test_import_pdf_export as fixtures


def company(root):
    info=root/'Info';info.mkdir()
    wb=Workbook();ws=wb.active;ws.title='Sheet1'
    for row in range(1,7):ws.cell(row,1,f'Label{row}');ws.cell(row,2,f'Value{row}')
    wb.save(info/'ZZ_Company.xlsx')
    template=Workbook();ws=template.active;ws.title='Sheet1';ws.merge_cells('A1:B1');ws['A1']='Document Request  ________   '
    ws['A1'].font=Font(underline='singleAccounting')
    ws['A31']='Closing text';ws.print_area='A1:B32';ws.page_setup.paperSize=ws.PAPERSIZE_A4
    ws['A32']='________';ws['B31'].border=Border(bottom=Side(style='thin'))
    ws['A33']='Signature: ________'
    choices=template.create_sheet('Sheet2')
    for i in range(20):choices.cell(i+7,1,f'Request {i+1}')
    template.save(info/'Template.xlsx')


class AuditDocumentTests(unittest.TestCase):
    def test_vba_module11_mapping_order_and_input_integrity(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);company(root)
            template=root/'Info'/'Template.xlsx';before=template.read_bytes()
            self.assertEqual(engine.first_xlsx(root/'Info').name,'ZZ_Company.xlsx')
            self.assertEqual(docs.company_info(root),[(f'Label{i}',f'Value{i}') for i in (1,2,5,6,3)])
            output=docs.build_request(root,[0,4,2])
            ws=load_workbook(output).active
            self.assertEqual(ws.title,'Doc Requirement')
            self.assertEqual(ws['B6'].value,'Value1');self.assertEqual(ws['B8'].value,'Value5')
            self.assertEqual(ws['A14'].value.strip(),'1. Request 1')
            self.assertEqual(ws['A15'].value.strip(),'2. Request 3')
            self.assertEqual(ws['A16'].value.strip(),'3. Request 5')
            self.assertFalse(ws.row_dimensions[16].hidden);self.assertTrue(ws.row_dimensions[17].hidden)
            self.assertEqual(ws['A31'].value,'Closing text');self.assertIn('A1:B1',[str(r) for r in ws.merged_cells.ranges])
            self.assertEqual(ws['A1'].value,'Document Request');self.assertEqual(ws['A1'].font.underline,'single')
            self.assertEqual(ws['A32'].value,'________');self.assertEqual(ws['B31'].border.bottom.style,'thin')
            self.assertEqual(ws['A33'].value,'Signature: ________')
            self.assertEqual(template.read_bytes(),before)
            self.assertEqual(docs.build_request(root,[0]),output)
            current=load_workbook(output)
            self.assertIsNone(current.active['A15'].value);self.assertTrue(current.active.row_dimensions[16].hidden)
            current.close()
            self.assertEqual(list(output.parent.glob('Doc_Requirement*.xlsx')),[output])
            with self.assertRaisesRegex(ValueError,'17 slots'):docs.build_request(root,list(range(18)))
            with self.assertRaises(ValueError):docs.build_request(root,[])

    def test_request_overwrite_preserves_last_save_on_failure(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            output=docs.publish_request(root,'.xlsx',lambda p:p.write_bytes(b'last complete request'))
            def failed_save(path):
                path.write_bytes(b'incomplete');raise RuntimeError('save failed')
            with self.assertRaisesRegex(RuntimeError,'save failed'):docs.publish_request(root,'.xlsx',failed_save)
            self.assertEqual(output.read_bytes(),b'last complete request')
            with patch.object(docs.os,'replace',side_effect=PermissionError('file locked')):
                with self.assertRaisesRegex(PermissionError,'Close Doc_Requirement.xlsx'):
                    docs.publish_request(root,'.xlsx',lambda p:p.write_bytes(b'new request'))
            self.assertEqual(output.read_bytes(),b'last complete request')
            self.assertFalse(list(root.glob('.doc-request-*')))
            self.assertEqual(docs.publish_request(root,'.xlsx',lambda p:p.write_bytes(b'new request')),output)
            self.assertEqual(output.read_bytes(),b'new request')

    def test_meeting_word_uses_current_company_and_edited_fields(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);company(root)
            values=docs.default_meeting(docs.company_info(root))
            values.update(meeting_date='2026-11-16',meeting_details='10:00 at company office',reference='Legal reference',
                          notice='Current notice',purpose='Current agenda',closing='Meeting closed',
                          discussion='Current meeting findings\nកំណត់ហេតុថ្មី',requested_documents='Request A\nRequest B',
                          auditors='Current auditor',actions='Follow up next week')
            output=docs.save_meeting(root,values);document=Document(output)
            text='\n'.join(n.text or '' for n in document._element.iter(qn('w:t')))
            self.assertIn('Value1',text);self.assertIn('កំណត់ហេតុថ្មី',text);self.assertIn('Current auditor',text)
            self.assertIn('Follow up next week',text);self.assertNotIn('{{discussion}}',text)
            ordered=(values['title'],'Value1','2026-11-16','10:00 at company office','Legal reference',
                     'Current notice','Current auditor','Current agenda','Current meeting findings',
                     'Request A','Follow up next week','Meeting closed')
            positions=[text.index(value) for value in ordered];self.assertEqual(positions,sorted(positions))
            self.assertEqual(output.name,'Meeting_Note_2026-11-16_Note_1.docx')
            note=next(p for p in document.paragraphs if 'Current meeting findings' in p.text)
            self.assertEqual(note.paragraph_format.left_indent,0);self.assertEqual(note.paragraph_format.right_indent,0)
            for old in ('Hoinco','L001-902401242','AU202504070115'):self.assertNotIn(old,text)
            with ZipFile(output) as archive:
                self.assertIn('word/media/image1.png',archive.namelist())
            self.assertGreater(len(list(document._element.iter(qn('w:txbxContent')))),0)
            second=docs.save_meeting(root,values)
            self.assertEqual(second.name,'Meeting_Note_2026-11-16_Note_2.docx')
            values['meeting_date']='2026-11-17'
            self.assertEqual(docs.save_meeting(root,values).name,'Meeting_Note_2026-11-17_Note_1.docx')

    def test_meeting_date_validation_and_empty_optional_fields(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);values=docs.default_meeting([('Company:', 'Current company')])
            self.assertEqual(values['company_info'],'Company: Current company')
            for invalid in ('','2026-02-30','not a date','../x'):
                with self.subTest(date=invalid),self.assertRaisesRegex(ValueError,'YYYY-MM-DD'):
                    docs.save_meeting(root,dict(values,meeting_date=invalid))
            self.assertFalse((root/'Audit_Tools').exists())
            output=docs.save_meeting(root,values)
            text='\n'.join(p.text for p in Document(output).paragraphs)
            self.assertNotIn('{{',text);self.assertNotIn('សវនករស្នើសុំឯកសារ',text)

    def test_missing_template_and_unknown_info_never_read_template_as_company(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);(root/'Info').mkdir()
            wb=Workbook();wb.save(root/'Info'/'Template.xlsx')
            with self.assertRaisesRegex(ValueError,'company Info workbook'):docs.company_info(root)
            with self.assertRaisesRegex(ValueError,'Sheet1.*Sheet2'):docs.request_items(root)

    def test_doc_pdf_is_rendered_and_original_template_stays_unchanged(self):
        try:find_soffice()
        except RuntimeError:self.skipTest('LibreOffice unavailable')
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);company(root);source=root/'Info'/'Template.xlsx';before=source.read_bytes()
            output=docs.request_pdf(root,[0,1],renderer='libreoffice')
            import pdfplumber
            with pdfplumber.open(output) as pdf:
                text=' '.join(p.extract_text() or '' for p in pdf.pages)
                self.assertIn('Request 1',text);self.assertIn('Value1',text)
            self.assertEqual(docs.request_pdf(root,[1],renderer='libreoffice'),output)
            with pdfplumber.open(output) as pdf:
                text=' '.join(p.extract_text() or '' for p in pdf.pages)
                self.assertIn('Request 2',text);self.assertNotIn('Request 1',text)
            self.assertEqual(list(output.parent.glob('Doc_Requirement*.pdf')),[output])
            self.assertEqual(len(list(output.parent.glob('Doc_Requirement*.xlsx'))),1)
            self.assertEqual(source.read_bytes(),before)

    def test_pdf_can_save_beside_toi_without_polluting_source_inventory(self):
        try:find_soffice()
        except RuntimeError:self.skipTest('LibreOffice unavailable')
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);toi=root/'TOI';toi.mkdir();fixtures.ToiPdfTests().make_book(toi/'TOI25.xlsx')
            rows=export_toi_pdfs(root,'1-1',renderer='libreoffice',output_folder='TOI')
            self.assertEqual(rows[0][2],'EXPORTED',rows)
            output=Path(rows[0][3]);self.assertEqual(output.parent,toi);self.assertTrue(is_toi_output(output))
            self.assertEqual(next(r for r in source_inventory(root) if r[0]=='TOI')[2],1)
            self.assertFalse(any(r[2]==output.name for r in validate_sources(root)))

    def test_generated_output_folders_are_not_consolidated(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);pr=root/'PR'/'2025';pr.mkdir(parents=True)
            wb=Workbook();ws=wb.active;layout=engine.cons.LAYOUTS['PR']
            ws.cell(2,layout.first_col,layout.first_heading);ws.cell(5,1,'Transaction');ws.cell(5,7,100)
            wb.save(pr/'source.xlsx')
            for name in ('Audit_Tools','TOI_PDFs'):
                directory=root/name/'2025';directory.mkdir(parents=True);(directory/'not-tax.xlsx').write_bytes(b'not input')
            built,_=engine.build_workbook(root)
            self.assertEqual(built['PR']['G1'].value,100)
            self.assertNotIn('Audit_Tools',built.sheetnames);self.assertNotIn('TOI_PDFs',built.sheetnames)

    def test_build_sound_uses_main_thread_notification(self):
        app=Mock()
        with patch.object(ui.sys,'platform','win32'),patch.dict(sys.modules,{'winsound':Mock(MB_OK=0)}) as modules:
            ui.App.notify_build_done(app);modules['winsound'].MessageBeep.assert_called_once_with(0)
        with patch.object(ui.sys,'platform','linux'):ui.App.notify_build_done(app);app.bell.assert_called_once()

    def test_busy_refresh_defers_reset_until_current_work_finishes(self):
        app=Mock();app.busy=True;ui.App.refresh(app)
        self.assertTrue(app.reset_pending);app.reset_session.assert_not_called()
        app.busy=False;ui.App.refresh(app);app.reset_session.assert_called_once()


if __name__=='__main__':unittest.main()
