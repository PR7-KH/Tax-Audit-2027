"""Local regulation library and the workbook's literal, case-insensitive search."""
from dataclasses import dataclass
from functools import lru_cache
import hashlib
import json
from pathlib import Path

LIBRARY_PATH = Path(__file__).resolve().parent / 'assets' / 'tax_laws.json'


@dataclass(frozen=True)
class LawEntry:
    row: int
    reference: str
    text: str

    @property
    def group(self):
        return self.reference.split('_Article', 1)[0].strip()

    @property
    def language(self):
        return 'English' if self.reference.replace('\u200b', '').upper().endswith('-EN') else 'Khmer'


class LawLibrary:
    def __init__(self, data):
        if not isinstance(data, dict) or data.get('schema_version') != 1:
            raise ValueError('Unsupported tax law library format.')
        source = data.get('source')
        rows = data.get('entries')
        characters = data.get('khmer_characters')
        if not isinstance(source, dict) or not isinstance(source.get('filename'), str):
            raise ValueError('The tax law library is missing its source information.')
        if not isinstance(rows, list) or not rows:
            raise ValueError('The tax law library has no entries.')
        entries = []
        seen_rows = set()
        for item in rows:
            if (not isinstance(item, dict) or type(item.get('row')) is not int or item['row'] < 2
                    or item['row'] in seen_rows
                    or not isinstance(item.get('reference'), str) or not item['reference'].strip()
                    or not isinstance(item.get('text'), str) or not item['text'].strip()):
                raise ValueError('The tax law library contains an invalid entry.')
            seen_rows.add(item['row'])
            entries.append(LawEntry(item['row'], item['reference'], item['text']))
        if not isinstance(characters, list) or any(not isinstance(c, str) or not c for c in characters):
            raise ValueError('The Khmer character helper is invalid.')
        self.entries = tuple(entries)
        self.khmer_characters = tuple(characters)
        self.source_filename = source['filename']
        self.source_sha256 = source.get('sha256', '')
        self.groups = tuple(dict.fromkeys(entry.group for entry in self.entries))
        self.cell_limit_entries = tuple(entry for entry in self.entries if len(entry.text) >= 32767)
        self._index = tuple((entry.reference.casefold(), entry.text.casefold()) for entry in self.entries)

    def search(self, query='', group=None, language=None):
        """Match a whole literal phrase in either code or text, in source order.

        Empty queries browse the library; the UI separately implements Clear.
        Repeated references are retained and selected by source row, not name.
        """
        needle = query.strip().casefold()
        return [entry for entry, (code, text) in zip(self.entries, self._index)
                if (not group or entry.group == group)
                and (not language or entry.language == language)
                and (not needle or needle in code or needle in text)]


@lru_cache(maxsize=1)
def load_library(path=LIBRARY_PATH):
    return LawLibrary(json.loads(Path(path).read_text(encoding='utf-8')))


def import_workbook(source):
    """Read cells without running VBA, Excel, external links or browser controls."""
    from openpyxl import load_workbook

    source = Path(source)
    workbook = load_workbook(source, read_only=True, data_only=False, keep_links=False)
    try:
        if 'Regulations' not in workbook.sheetnames:
            raise ValueError('The law workbook needs a Regulations sheet with code/text in A/B.')
        entries = []
        for number, cells in enumerate(workbook['Regulations'].iter_rows(min_row=2, max_col=2), 2):
            if all(cell.value is None for cell in cells):
                continue
            if any(cell.data_type == 'f' for cell in cells):
                raise ValueError(f'Regulations row {number}: store the code and law text as values, not formulas.')
            code, text = (cell.value for cell in cells)
            if not isinstance(code, str) or not code.strip() or not isinstance(text, str) or not text.strip():
                raise ValueError(f'Regulations row {number}: both a reference and law text are required.')
            entries.append({'row': number, 'reference': code, 'text': text})
        characters = []
        if 'Khmer' in workbook.sheetnames:
            for row in workbook['Khmer'].iter_rows(max_col=1):
                cell = row[0]
                if cell.value is None:
                    continue
                if cell.data_type == 'f' or not isinstance(cell.value, str):
                    raise ValueError('The Khmer helper must contain text characters, not formulas.')
                if cell.value:
                    characters.append(cell.value)
        data = {'schema_version': 1,
                'source': {'filename': source.name, 'sha256': hashlib.sha256(source.read_bytes()).hexdigest(),
                           'sheet': 'Regulations'},
                'khmer_characters': characters, 'entries': entries}
        LawLibrary(data)  # Validate before a caller replaces the bundled library.
        return data
    finally:
        workbook.close()
