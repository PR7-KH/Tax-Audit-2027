import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock,patch
from zipfile import ZipFile
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import Workbook,load_workbook
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
    template=Workbook();ws=template.active;ws.title='Sheet1';ws.merge_cells('A1:B1');ws['A1']='Document Request'
    ws['A31']='Closing text';ws.print_area='A1:B32';ws.page_setup.paperSize=ws.PAPERSIZE_A4
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
            self.assertEqual(template.read_bytes(),before)
            self.assertNotEqual(docs.build_request(root,[0]),output)
            with self.assertRaisesRegex(ValueError,'17 slots'):docs.build_request(root,list(range(18)))
            with self.assertRaises(ValueError):docs.build_request(root,[])

    def test_meeting_word_uses_current_company_and_edited_fields(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);company(root)
            values=docs.default_meeting(docs.company_info(root))
            values.update(discussion='Current meeting findings\nកំណត់ហេតុថ្មី',requested_documents='Request A\nRequest B',auditors='Current auditor',actions='Follow up next week')
            output=docs.save_meeting(root,values);document=Document(output)
            text='\n'.join(n.text or '' for n in document._element.iter(qn('w:t')))
            self.assertIn('Value1',text);self.assertIn('កំណត់ហេតុថ្មី',text);self.assertIn('Current auditor',text)
            self.assertIn('Follow up next week',text);self.assertNotIn('{{discussion}}',text)
            for old in ('Hoinco','L001-902401242','AU202504070115'):self.assertNotIn(old,text)
            with ZipFile(output) as archive:
                self.assertIn('word/media/image1.png',archive.namelist())
            self.assertGreater(len(list(document._element.iter(qn('w:txbxContent')))),0)
            self.assertNotEqual(docs.save_meeting(root,values),output)

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
