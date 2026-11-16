import hashlib
import json
from pathlib import Path
import tempfile
import unittest
from openpyxl import Workbook
from tax_law import LawLibrary, import_workbook, load_library


def sample_data():
    return {'schema_version': 1, 'source': {'filename': 'test compilation.xlsm'},
            'khmer_characters': ['ឯ', 'ឫ'], 'entries': [
                {'row': 2, 'reference': 'LOT(2023)_Article1',
                 'text': 'ពន្ធលើប្រាក់ចំណូល\nពន្ធលើប្រាក់ចំណូល និងឯកសារ'},
                {'row': 3, 'reference': 'LOT(2023)_Article2-\u200bEN',
                 'text': 'Income tax; INCOME TAX.\nLiteral [tax] and a.b.\n<script>plain text</script>'},
                {'row': 5, 'reference': 'LOT(2023)_Article2-\u200bEN',
                 'text': 'Second entry with an identical reference. Income tax.'},
                {'row': 8, 'reference': 'VAT_Article9-EN', 'text': 'Value added tax documentation.'}]}


class TaxLawTests(unittest.TestCase):
    def test_literal_phrase_search_preserves_workbook_order_and_duplicates(self):
        library = LawLibrary(sample_data())
        self.assertEqual([e.row for e in library.search('  INCOME tax ')], [3, 5])
        self.assertEqual([e.row for e in library.search('LOT(2023)_Article')], [2, 3, 5])
        self.assertEqual([e.row for e in library.search('ពន្ធលើប្រាក់ចំណូល')], [2])
        self.assertEqual([e.row for e in library.search('[tax]')], [3])
        self.assertEqual([e.row for e in library.search('a.b')], [3])
        self.assertEqual(library.search('.*'), [])
        self.assertEqual(library.search('tax income'), [])
        self.assertEqual([e.row for e in library.search()], [2, 3, 5, 8])

    def test_filters_refine_results_without_changing_the_text(self):
        library = LawLibrary(sample_data())
        self.assertEqual(library.groups, ('LOT(2023)', 'VAT'))
        self.assertEqual([e.row for e in library.search('tax', group='LOT(2023)', language='English')], [3, 5])
        self.assertEqual([e.row for e in library.search(group='LOT(2023)', language='Khmer')], [2])
        self.assertEqual([e.row for e in library.search(group='VAT')], [8])
        self.assertEqual(library.entries[1].text, sample_data()['entries'][1]['text'])

    def test_bundled_catalog_has_all_entries_and_the_original_helper(self):
        library = load_library()
        self.assertEqual(len(library.entries), 741)
        self.assertEqual(len(library.khmer_characters), 11)
        self.assertEqual(library.entries[0].reference, 'LOT(16/05-2023)_Article1')
        self.assertIn('មាត្រា១', library.entries[0].text)
        self.assertEqual(sum(e.language == 'English' for e in library.entries), 363)
        self.assertTrue(library.search('income tax'))
        self.assertEqual(len(library.cell_limit_entries), 1)
        self.assertEqual(len(library.cell_limit_entries[0].text), 32767)
        self.assertRegex(library.source_sha256, r'^[0-9a-f]{64}$')

    def test_workbook_import_preserves_unicode_lines_rows_and_source_bytes(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'laws.xlsx'
            workbook = Workbook(); sheet = workbook.active; sheet.title = 'Regulations'
            sheet.append(['Regulation Code', 'Definition'])
            for entry in sample_data()['entries']:
                sheet.cell(entry['row'], 1, entry['reference'])
                sheet.cell(entry['row'], 2, entry['text'])
            helper = workbook.create_sheet('Khmer')
            for character in sample_data()['khmer_characters']:
                helper.append([character])
            workbook.save(path); workbook.close()
            before = path.read_bytes()
            data = import_workbook(path)
            self.assertEqual(data['entries'], sample_data()['entries'])
            self.assertEqual(data['khmer_characters'], sample_data()['khmer_characters'])
            self.assertEqual(data['source']['sha256'], hashlib.sha256(before).hexdigest())
            self.assertEqual(path.read_bytes(), before)
            # A JSON round trip preserves characters, tabs and literal markup.
            self.assertEqual(LawLibrary(json.loads(json.dumps(data))).entries[1].text,
                             sample_data()['entries'][1]['text'])

    def test_import_rejects_formulas_or_incomplete_rows_instead_of_losing_laws(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / 'laws.xlsx'
            workbook = Workbook(); sheet = workbook.active; sheet.title = 'Regulations'
            sheet.append(['Code', 'Text']); sheet.append(['A', '=1+1'])
            workbook.save(path)
            with self.assertRaisesRegex(ValueError, 'not formulas'):
                import_workbook(path)
            sheet['B2'] = None; workbook.save(path)
            with self.assertRaisesRegex(ValueError, 'both a reference and law text'):
                import_workbook(path)
            workbook.close()

    def test_invalid_library_reports_a_useful_error(self):
        for data in ([], {}, dict(sample_data(), entries=[]), dict(sample_data(), khmer_characters=[None]),
                     dict(sample_data(), entries=[{'row': 2, 'reference': 'A', 'text': None}])):
            with self.subTest(data=type(data).__name__), self.assertRaises(ValueError):
                LawLibrary(data)
        duplicate_row = sample_data()
        duplicate_row['entries'][1]['row'] = 2
        with self.assertRaisesRegex(ValueError, 'invalid entry'):
            LawLibrary(duplicate_row)


if __name__ == '__main__':
    unittest.main()
