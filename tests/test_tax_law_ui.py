import tkinter as tk
import unittest
from unittest.mock import patch
from tax_law import LawLibrary
from tax_law_ui import TaxLawPanel, ALL_GROUPS, ALL_LANGUAGES
from test_tax_law import sample_data


class TaxLawUiTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest('Tk display unavailable; run with DISPLAY or xvfb-run for live UI checks.')
        self.root.geometry('1100x740+0+0')
        self.errors = []
        self.root.report_callback_exception = lambda *args: self.errors.append(args)
        with patch('tax_law_ui.load_library', return_value=LawLibrary(sample_data())):
            self.panel = TaxLawPanel(self.root)
        self.panel.pack(fill='both', expand=True)
        self.root.update()

    def tearDown(self):
        if hasattr(self, 'panel'):
            self.panel.destroy()
        if hasattr(self, 'root'):
            self.root.destroy()
        if hasattr(self, 'errors'):
            self.assertFalse(self.errors)

    def test_matches_are_highlighted_navigation_wraps_and_duplicate_references_open_correctly(self):
        panel = self.panel
        panel.query.set('income tax'); panel.search(); self.root.update()
        self.assertEqual(panel.results.get_children(), ('3', '5'))
        self.assertEqual(len(panel.matches), 2)
        self.assertEqual(len(panel.body.tag_ranges('match')), 4)
        self.assertEqual(panel.match_index, 0)
        panel.jump(-1); self.assertEqual(panel.match_index, 1)
        panel.jump(1); self.assertEqual(panel.match_index, 0)
        panel.results.selection_set('5'); self.root.update()
        self.assertEqual(panel.body.get('1.0', 'end-1c'), sample_data()['entries'][2]['text'])
        panel.copy_article(False)
        self.assertEqual(panel.clipboard_get(), panel.current_entry.reference + '\n\n' + panel.current_entry.text)

    def test_khmer_helper_literal_punctuation_and_read_only_text(self):
        panel = self.panel
        panel.query.set('ពន្ធលើប្រាក់ចំណូល'); panel.search(); self.root.update()
        self.assertEqual(panel.results.get_children(), ('2',))
        self.assertEqual(len(panel.matches), 2)
        panel.query.set('[tax]'); panel.search(); self.root.update()
        self.assertEqual(len(panel.matches), 1)
        text = panel.body.get('1.0', 'end-1c')
        self.assertIn('<script>plain text</script>', text)
        panel.body.insert('1.0', 'Unwanted edit')
        self.assertEqual(panel.body.get('1.0', 'end-1c'), text)
        panel.query.set('AB'); panel.search_entry.icursor(1); panel.insert_character('ឯ')
        self.assertEqual(panel.query.get(), 'AឯB')

    def test_filters_browse_font_and_reset_clear_all_state_and_pending_search(self):
        panel = self.panel
        panel.language.set('English'); panel.group.set('VAT'); panel.search(browse=True); self.root.update()
        self.assertEqual(panel.results.get_children(), ('8',))
        panel.resize_font(1); self.assertEqual(panel.font_size, 13)
        panel.toggle_characters(); panel.query.set('tax')
        self.assertIsNotNone(panel._timer)
        panel.reset(); self.root.update()
        self.assertIsNone(panel._timer)
        self.assertEqual(panel.query.get(), '')
        self.assertEqual(panel.group.get(), ALL_GROUPS)
        self.assertEqual(panel.language.get(), ALL_LANGUAGES)
        self.assertEqual(panel.font_size, 12)
        self.assertFalse(panel.results.get_children())
        self.assertEqual(panel.body.get('1.0', 'end-1c'), '')
        self.assertIsNone(panel.current_entry)
        self.assertFalse(panel.helper_frame.winfo_manager())

    def test_missing_library_does_not_crash_the_company_tools(self):
        self.panel.destroy()
        with patch('tax_law_ui.load_library', side_effect=FileNotFoundError('missing library')):
            self.panel = TaxLawPanel(self.root)
        self.panel.pack(fill='both', expand=True)
        self.panel.query.set('tax'); self.panel.search(); self.root.update()
        self.assertIn('assets/tax_laws.json', self.panel.note.cget('text'))
        self.assertFalse(self.panel.results.get_children())


if __name__ == '__main__':
    unittest.main()
