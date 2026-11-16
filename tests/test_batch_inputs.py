import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from test_input_tools import write_pdf
from input_tools import (batch_convert_pdfs, selected_pdfs, source_inventory,
                         validate_sources, conversion_records, converted_output)
import control_center as ui


class BatchInputsTests(unittest.TestCase):
    def test_target_only_scope_and_menu_order(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            for directory in ('Bank','GL','Unrelated'):
                (root/directory).mkdir(); write_pdf(root/directory/'statement.pdf')
            write_pdf(root/'outside.pdf')
            paths=selected_pdfs(root)
            self.assertEqual({p.parent.name for p in paths},{'Bank','GL'})
            rows=validate_sources(root)
            self.assertFalse(any('Unrelated' in str(r) or 'outside.pdf' in str(r) for r in rows))
            items=[(key,label) for kind,key,label in ui.MENU if kind=='item']
            self.assertEqual(items[:4],[('dashboard','Audit Dashboard'),('tois','TOIs Overview'),('mtr','MTR Overview'),('consolidate','Consolidated data')])

    def test_batch_refresh_idempotency_and_preservation(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder); bank=root/'Bank';bank.mkdir()
            for n in range(2): write_pdf(bank/f'{n}.pdf')
            (bank/'0_review.xlsx').write_bytes(b'user workbook')
            before={p.name:p.read_bytes() for p in bank.glob('*.pdf')}
            calls=[]
            result=batch_convert_pdfs(root,progress=lambda a,b:calls.append((a,b)))
            self.assertEqual(len(result),2)
            self.assertTrue(all(r[4].startswith('CONVERTED:') for r in result))
            self.assertEqual(calls,[(1,2),(2,2)])
            self.assertEqual({p.name:p.read_bytes() for p in bank.glob('*.pdf')},before)
            self.assertEqual((bank/'0_review.xlsx').read_bytes(),b'user workbook')
            outputs=list((bank/'.pdf-extractions').glob('*.xlsx'))
            self.assertEqual(len(outputs),2)
            again=batch_convert_pdfs(root)
            self.assertTrue(all(r[4].startswith('ALREADY CONVERTED:') for r in again))
            self.assertEqual(len(list((bank/'.pdf-extractions').glob('*.xlsx'))),2)
            row=next(r for r in source_inventory(root) if r[0]=='Bank')
            self.assertTrue(row[3].startswith('Converted:'))
            self.assertIn('review required',row[3])
            self.assertTrue(any(r[4].startswith('CONVERTED:') for r in validate_sources(root)))

    def test_scope_errors_continue_and_tax_drafts_are_not_imported(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder)
            for year in ('2025','2026'):
                directory=root/'PR'/year;directory.mkdir(parents=True)
                write_pdf(directory/'valid.pdf')
            bad=root/'PR'/'2025'/'scan.pdf';write_pdf(bad,blank=True)
            paths=selected_pdfs(root,['PR','2025'],overview=True)
            self.assertEqual(len(paths),2)
            results=batch_convert_pdfs(root,paths)
            self.assertEqual(sum(r[4].startswith('ERROR:') for r in results),1)
            self.assertEqual(sum(r[4].startswith('CONVERTED:') for r in results),1)
            self.assertFalse(list((root/'PR'/'2025').glob('*.xlsx')))
            self.assertFalse((root/'PR'/'2026'/'.pdf-extractions').exists())
            self.assertIn('Not ready',next(r for r in source_inventory(root) if r[:2]==['PR','2025'])[3])
            single=selected_pdfs(root,['PR','PR/2026','valid.pdf',0,'NOT READY'])
            self.assertEqual(single,[root/'PR'/'2026'/'valid.pdf'])
            with self.assertRaises(ValueError):batch_convert_pdfs(root,[root/'outside.pdf'])

    def test_changed_source_needs_reconversion_without_replacing_old_output(self):
        with tempfile.TemporaryDirectory() as folder:
            root=Path(folder);bank=root/'Bank';bank.mkdir();pdf=bank/'bank.pdf';write_pdf(pdf)
            batch_convert_pdfs(root)
            old=converted_output(root,pdf,conversion_records(root));content=old.read_bytes()
            pdf.write_bytes(pdf.read_bytes()+b'\n% changed\n')
            self.assertIsNone(converted_output(root,pdf,conversion_records(root)))
            self.assertIn('Not ready',next(r for r in source_inventory(root) if r[0]=='Bank')[3])
            result=batch_convert_pdfs(root)
            self.assertTrue(result[0][4].startswith('CONVERTED:'))
            self.assertEqual(old.read_bytes(),content)
            self.assertNotEqual(converted_output(root,pdf,conversion_records(root)),old)

    def test_worker_needs_no_file_or_destination_dialogs(self):
        import queue
        with tempfile.TemporaryDirectory() as folder:
            bank=Path(folder)/'Bank';bank.mkdir();write_pdf(bank/'bank.pdf')
            app=Mock();app.job_settings=(folder,3,False);app.q=queue.Queue()
            result=ui.App.job_batch_convert(app,['Bank','All',1,'Not ready','1 PDF'],'sources')
            self.assertEqual(result[0],'batch_done')
            self.assertEqual(len(result[1]),1)
            self.assertEqual(app.q.get_nowait(),('batch_progress',1,1))


if __name__=='__main__':unittest.main()
