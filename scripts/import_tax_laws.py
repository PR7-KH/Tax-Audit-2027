"""Rebuild the local law library from the user's compilation, without executing VBA."""
import argparse
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tax_law import LIBRARY_PATH, LawLibrary, import_workbook


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('workbook', type=Path)
    parser.add_argument('--output', type=Path, default=LIBRARY_PATH)
    args = parser.parse_args()
    if args.output.resolve() == args.workbook.resolve():
        parser.error('The output must not replace the source workbook.')
    data = import_workbook(args.workbook)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(data, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    library = LawLibrary(data)
    print(f'Imported {len(library.entries)} entries and {len(library.khmer_characters)} helper characters.')
    for entry in library.cell_limit_entries:
        print(f'Source-size note: row {entry.row} ({entry.reference}) reaches Excel\'s 32,767-character cell limit; text preserved as stored.')


if __name__ == '__main__':
    main()
