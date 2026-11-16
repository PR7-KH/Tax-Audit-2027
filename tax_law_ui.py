"""In-app law lookup; all article text is displayed as plain, read-only text."""
import tkinter as tk
from tkinter import font as tkfont, messagebox, scrolledtext, ttk

from tax_law import load_library

TEXT, WHITE, LINE = '#172b43', '#ffffff', '#d7e1eb'
ALL_GROUPS, ALL_LANGUAGES = 'All regulations', 'All languages'


class TaxLawPanel(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent, bg='#f0f4f8')
        self.library = None
        self.load_error = ''
        try:
            self.library = load_library()
        except (OSError, ValueError) as exc:
            self.load_error = f'Law library unavailable: {exc}. Copy assets/tax_laws.json from the complete update.'
        self.query = tk.StringVar()
        self.group = tk.StringVar(value=ALL_GROUPS)
        self.language = tk.StringVar(value=ALL_LANGUAGES)
        self._timer = None
        self._updating = False
        self.current_entry = None
        self.current_query = ''
        self.matches = []
        self.match_index = -1
        self.entries_by_row = {str(e.row): e for e in self.library.entries} if self.library else {}
        self.font_size = 12
        families = set(tkfont.families(self))
        family = next((f for f in ('Khmer OS Siemreap', 'Noto Sans Khmer', 'Noto Serif Khmer') if f in families),
                      tkfont.nametofont('TkTextFont').actual('family'))
        self.reader_font = tkfont.Font(self, family=family, size=self.font_size)

        search_bar = ttk.Frame(self)
        search_bar.pack(fill='x', padx=8, pady=(8, 4))
        ttk.Label(search_bar, text='Search code or article text').pack(side='left', padx=(0, 8))
        self.search_entry = ttk.Entry(search_bar, textvariable=self.query, width=30)
        self.search_entry.pack(side='left', fill='x', expand=True)
        ttk.Button(search_bar, text='Search', command=self.search).pack(side='left', padx=4)
        ttk.Button(search_bar, text='Clear', command=self.reset).pack(side='left')
        self.search_entry.bind('<Return>', lambda event: self.search())
        self.search_entry.bind('<Escape>', lambda event: self.reset())

        filters = ttk.Frame(self)
        filters.pack(fill='x', padx=8, pady=(0, 6))
        groups = (ALL_GROUPS,) + (self.library.groups if self.library else ())
        selector = ttk.Combobox(filters, textvariable=self.group, values=groups, state='readonly', width=35)
        selector.pack(side='left', padx=(0, 6))
        selector.bind('<<ComboboxSelected>>', lambda event: self.search(browse=True))
        selector = ttk.Combobox(filters, textvariable=self.language,
                               values=(ALL_LANGUAGES, 'Khmer', 'English'), state='readonly', width=13)
        selector.pack(side='left', padx=(0, 6))
        selector.bind('<<ComboboxSelected>>', lambda event: self.search(browse=True))
        ttk.Button(filters, text='Browse all', command=lambda: self.search(browse=True)).pack(side='left', padx=3)
        ttk.Button(filters, text='Khmer characters', command=self.toggle_characters).pack(side='left', padx=3)
        ttk.Button(filters, text='Library info', command=self.library_info).pack(side='left', padx=3)

        self.helper_frame = ttk.Frame(self)
        ttk.Label(self.helper_frame, text='Insert:').pack(side='left', padx=8)
        for character in self.library.khmer_characters if self.library else ():
            ttk.Button(self.helper_frame, text=character, width=3,
                       command=lambda c=character: self.insert_character(c)).pack(side='left', padx=2, pady=4)

        panes = self.panes = ttk.Panedwindow(self, orient='horizontal')
        panes.pack(fill='both', expand=True, padx=8, pady=3)
        results_card = ttk.Frame(panes, width=320)
        reader = tk.Frame(panes, bg=WHITE, highlightthickness=1, highlightbackground=LINE)
        panes.add(results_card, weight=1)
        panes.add(reader, weight=3)
        self._sized = False
        panes.bind('<Configure>', self.size_panes)
        self.result_count = ttk.Label(results_card, text='Type a phrase or choose Browse all.')
        self.result_count.pack(fill='x', pady=(0, 5))
        table = ttk.Frame(results_card)
        table.pack(fill='both', expand=True)
        style = ttk.Style(self)
        style.configure('TaxLaw.Treeview', rowheight=30, font=('Helvetica', 10), foreground=TEXT, background=WHITE)
        self.results = ttk.Treeview(table, columns=('reference',), show='headings', selectmode='browse', style='TaxLaw.Treeview')
        self.results.heading('reference', text='Regulation / article reference')
        self.results.column('reference', width=310, minwidth=150, stretch=True)
        vertical = ttk.Scrollbar(table, orient='vertical', command=self.results.yview)
        horizontal = ttk.Scrollbar(table, orient='horizontal', command=self.results.xview)
        self.results.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        vertical.pack(side='right', fill='y')
        horizontal.pack(side='bottom', fill='x')
        self.results.pack(fill='both', expand=True)
        self.results.tag_configure('alternate', background='#edf3f8')
        self.results.bind('<<TreeviewSelect>>', self.select_entry)

        self.reference_label = tk.Label(reader, text='', bg=WHITE, fg=TEXT, anchor='w', justify='left',
                                        font=('Helvetica', 12, 'bold'), wraplength=600)
        self.reference_label.pack(fill='x', padx=12, pady=(10, 4))
        self.reference_label.bind('<Configure>', lambda event: self.reference_label.config(wraplength=max(180, event.width)))
        copies = ttk.Frame(reader)
        copies.pack(fill='x', padx=10, pady=(0, 4))
        self.language_label = ttk.Label(copies, text='')
        self.language_label.pack(side='left')
        ttk.Button(copies, text='Copy article', command=lambda: self.copy_article(False)).pack(side='right', padx=3)
        ttk.Button(copies, text='Copy reference', command=lambda: self.copy_article(True)).pack(side='right', padx=3)
        navigation = ttk.Frame(reader)
        navigation.pack(fill='x', padx=10, pady=(0, 4))
        ttk.Button(navigation, text='Previous', command=lambda: self.jump(-1)).pack(side='left', padx=(0, 3))
        ttk.Button(navigation, text='Next', command=lambda: self.jump(1)).pack(side='left', padx=3)
        self.match_label = ttk.Label(navigation, text='Text matches: 0 / 0')
        self.match_label.pack(side='left', padx=6)
        ttk.Button(navigation, text='+', width=3, command=lambda: self.resize_font(1)).pack(side='right')
        ttk.Button(navigation, text='-', width=3, command=lambda: self.resize_font(-1)).pack(side='right', padx=3)
        self.body = scrolledtext.ScrolledText(reader, wrap='word', font=self.reader_font, width=30, height=10,
                                             bg=WHITE, fg=TEXT, selectbackground='#087f8c', selectforeground=WHITE,
                                             padx=14, pady=10, spacing1=4, spacing3=6, state='disabled')
        self.body.pack(fill='both', expand=True, padx=1, pady=(0, 1))
        self.body.tag_configure('match', background='#cceede', foreground=TEXT)
        self.body.tag_configure('active_match', background='#ffe29a', foreground=TEXT)
        self.body.tag_raise('sel')
        for widget in (self.body, self.results, self.search_entry):
            widget.bind('<F3>', lambda event: self.jump(1))
            widget.bind('<Shift-F3>', lambda event: self.jump(-1))
            widget.bind('<Control-f>', self.focus_search)
            widget.bind('<Command-f>', self.focus_search)
        self.note = ttk.Label(self, text='', wraplength=1000)
        self.note.pack(side='bottom', fill='x', before=self.panes, padx=10, pady=(2, 6))
        self.query.trace_add('write', self.schedule_search)
        self.bind('<Destroy>', self.on_destroy)
        self.reset()

    def size_panes(self, event):
        if not self._sized and event.width > 100:
            self.panes.sashpos(0, max(240, int(event.width * .30)))
            self._sized = True

    def focus_search(self, event=None):
        self.search_entry.focus_set()
        self.search_entry.selection_range(0, 'end')
        return 'break'

    def cancel_search(self):
        if self._timer is not None:
            self.after_cancel(self._timer)
            self._timer = None

    def on_destroy(self, event):
        if event.widget is self:
            self.cancel_search()

    def schedule_search(self, *_):
        if not self._updating:
            self.cancel_search()
            self._timer = self.after(180, self.search)

    def reset(self):
        self.cancel_search()
        self._updating = True
        self.query.set('')
        self.group.set(ALL_GROUPS)
        self.language.set(ALL_LANGUAGES)
        self._updating = False
        self.results.delete(*self.results.get_children())
        self.result_count.config(text='Type a phrase or choose Browse all.')
        self.helper_frame.pack_forget()
        self.font_size = 12
        self.reader_font.configure(size=self.font_size)
        self.clear_reader()
        total = len(self.library.entries) if self.library else 0
        self.note.config(text=self.load_error or f'{total} entries · Local library from your supplied workbook · See Library info for source details.')

    def clear_reader(self):
        self.current_entry = None
        self.current_query = ''
        self.matches = []
        self.match_index = -1
        self.reference_label.config(text='Select a regulation to read its full text')
        self.language_label.config(text='')
        self.match_label.config(text='Text matches: 0 / 0')
        self.body.config(state='normal')
        self.body.delete('1.0', 'end')
        self.body.config(state='disabled')

    def search(self, browse=False):
        self.cancel_search()
        if self.library is None:
            self.note.config(text=self.load_error)
            return
        query = self.query.get().strip()
        group = None if self.group.get() == ALL_GROUPS else self.group.get()
        language = None if self.language.get() == ALL_LANGUAGES else self.language.get()
        entries = self.library.search(query, group, language) if query or browse else []
        self.results.delete(*self.results.get_children())
        self.clear_reader()
        for index, entry in enumerate(entries):
            self.results.insert('', 'end', iid=str(entry.row), values=(entry.reference,),
                                tags=('alternate',) if index % 2 else ())
        self.result_count.config(text=f'{len(entries)} / {len(self.library.entries)} entries')
        if entries:
            key = str(entries[0].row)
            self.results.selection_set(key)
            self.results.focus(key)
            self.show_entry(entries[0])
        else:
            self.reference_label.config(text='No matching regulations' if query or browse else 'Type a phrase or choose Browse all')
        self.note.config(text=f'{len(entries)} entries · Search matches a literal phrase in the reference or article text.')

    def select_entry(self, event=None):
        selected = self.results.selection()
        if selected:
            self.show_entry(self.entries_by_row[selected[0]])

    def show_entry(self, entry):
        keyword = self.query.get().strip()
        if self.current_entry is entry and self.current_query == keyword:
            return
        self.current_entry = entry
        self.current_query = keyword
        self.reference_label.config(text=entry.reference)
        self.language_label.config(text=entry.language)
        self.body.config(state='normal')
        self.body.delete('1.0', 'end')
        self.body.insert('1.0', entry.text)
        self.body.config(state='disabled')
        self.body.yview_moveto(0)
        self.matches = []
        self.match_index = -1
        if keyword:
            position = '1.0'
            count = tk.IntVar(self)
            while True:
                found = self.body.search(keyword, position, stopindex='end', nocase=True, regexp=False, count=count)
                if not found or count.get() <= 0:
                    break
                end = self.body.index(f'{found}+{count.get()}c')
                self.matches.append((found, end))
                self.body.tag_add('match', found, end)
                position = end
        if self.matches:
            self.jump(1)
        else:
            self.match_label.config(text='Text matches: 0 / 0 (reference match)' if keyword else 'Text matches: 0 / 0')

    def jump(self, direction):
        if not self.matches:
            return 'break'
        self.match_index = (self.match_index + direction) % len(self.matches)
        self.body.tag_remove('active_match', '1.0', 'end')
        start, end = self.matches[self.match_index]
        self.body.tag_add('active_match', start, end)
        self.body.see(start)
        self.match_label.config(text=f'Text matches: {self.match_index + 1} / {len(self.matches)}')
        return 'break'

    def resize_font(self, amount):
        self.font_size = min(22, max(9, self.font_size + amount))
        self.reader_font.configure(size=self.font_size)

    def toggle_characters(self):
        if self.helper_frame.winfo_manager():
            self.helper_frame.pack_forget()
        else:
            self.helper_frame.pack(fill='x', before=self.panes, padx=8)

    def insert_character(self, character):
        self.search_entry.focus_set()
        try:
            self.search_entry.delete('sel.first', 'sel.last')
        except tk.TclError:
            pass
        self.search_entry.insert('insert', character)

    def copy_article(self, reference_only):
        if self.current_entry is None:
            self.note.config(text='Select a regulation first.')
            return
        value = self.current_entry.reference
        if not reference_only:
            value += '\n\n' + self.current_entry.text
        self.clipboard_clear()
        self.clipboard_append(value)
        self.note.config(text='Reference copied.' if reference_only else 'Full article copied; paste it into your note or document.')

    def library_info(self):
        if self.library is None:
            messagebox.showinfo('Tax law library', self.load_error, parent=self)
            return
        text = (f'Source: {self.library.source_filename}\n'
                f'{len(self.library.entries)} entries from Regulations columns A/B.\n'
                'Codes and article wording are preserved as supplied. This is the uploaded compilation, not a live law feed.')
        if self.library.cell_limit_entries:
            references = '\n'.join(entry.reference for entry in self.library.cell_limit_entries)
            text += (f'\n\nSource-size note: {len(self.library.cell_limit_entries)} cell(s) reach Excel\'s '
                     f'32,767-character cell limit. Their stored text is preserved:\n{references}')
        messagebox.showinfo('Tax law library', text, parent=self)
