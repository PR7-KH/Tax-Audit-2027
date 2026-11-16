"""Source inventory, format validation and local PDF extraction helpers."""
from collections import Counter
from pathlib import Path
import os
import tempfile
from openpyxl import Workbook, load_workbook
import engine


def visible_files(directory):
    if directory is None:
        return []
    return sorted(p for p in directory.rglob('*') if p.is_file() and not any(
        part.startswith(('.', '~$')) for part in p.relative_to(directory).parts))


def format_counts(files):
    labels = {'.xlsx': 'Excel', '.xlsm': 'Excel macro', '.xls': 'Excel legacy', '.pdf': 'PDF', '.csv': 'CSV'}
    counts = Counter(labels.get(p.suffix.lower(), 'Other') for p in files)
    return ', '.join(f'{counts[label]} {label}' for label in sorted(counts)) or '—'


def format_status(files, tax=False):
    if not files:
        return 'No files'
    suffixes = {p.suffix.lower() for p in files}
    if '.pdf' in suffixes:
        return 'Not ready: convert PDF, review extracted data and map columns'
    if '.xls' in suffixes:
        return 'Not ready: save legacy Excel as .xlsx'
    if tax and suffixes != {'.xlsx'}:
        return 'Not ready: tax import requires V5.1 .xlsx layout'
    if not suffixes <= engine.FIN_EXT:
        return 'Not ready: unsupported format present'
    return 'Found; content not validated'


def source_inventory(folder):
    root = Path(folder)
    rows = []
    for tax in engine.cons.LAYOUTS:
        directory = engine.find_dir(root, tax)
        files = visible_files(directory)
        groups = {}
        for p in files:
            rel = p.relative_to(directory)
            groups.setdefault(rel.parts[0] if len(rel.parts) > 1 else 'Root (move to year folder)', []).append(p)
        for year, paths in sorted(groups.items()):
            status = format_status(paths, tax=True)
            if year.startswith('Root'):
                status = 'Not ready: monthly files must be under year folders'
            rows.append([tax, year, len(paths), status, format_counts(paths)])
        if not groups:
            rows.append([tax, '—', 0, 'Missing' if directory is None else 'Empty', '—'])
    for name in ('TOI', 'Info'):
        directory = engine.find_dir(root, name)
        files = visible_files(directory)
        rows.append([name, 'All', len(files), format_status(files, tax=True), format_counts(files)])
    for key, (label, _) in engine.FINANCE_FOLDERS.items():
        files = visible_files(engine.finance_folder(root, key))
        rows.append([label, 'All', len(files), format_status(files), format_counts(files)])
    return rows


def validate_sources(folder):
    """Keep the existing tax reader and supplement it with all-folder format checks."""
    rows = [list(r) for r in engine.check_sources(folder)
            if r[0] in engine.cons.LAYOUTS or r[4] == 'MISSING']
    for row in rows:
        if row[4] == 'READY' and row[3] == 0:
            row[4] = 'NOT READY: workbook contains no transaction rows' 
    root = Path(folder)
    tax_roots = {engine.find_dir(root, name) for name in engine.cons.LAYOUTS}
    for path in visible_files(root):
        rel = path.relative_to(root)
        suffix = path.suffix.lower()
        if suffix == '.pdf':
            result = 'NOT READY: PDF needs conversion, review and column mapping; scanned PDFs need OCR'
        elif suffix == '.xls':
            result = 'NOT READY: save as .xlsx and check the required layout'
        elif suffix == '.csv':
            result = 'REVIEW: CSV needs column mapping; tax input requires V5.1 .xlsx'
        elif suffix in ('.xlsx', '.xlsm'):
            if path.parent.parent in tax_roots and suffix == '.xlsx':
                continue  # already checked by the authoritative monthly reader
            try:
                wb = load_workbook(path, read_only=True, data_only=True)
                try:
                    if 'Conversion Review' in wb.sheetnames:
                        result = 'REVIEW: PDF extraction; reconcile totals and map columns before use'
                    else:
                        result = 'READABLE: workbook opens; report layout/mapping still needs validation'
                finally:
                    wb.close()
            except Exception as exc:
                result = 'ERROR: ' + str(exc).splitlines()[0]
        else:
            result = 'NOT READY: unsupported file format'
        rows.append([rel.parts[0], str(rel.parent), path.name, 0, result])
    return rows


def convert_pdf(source, output):
    """Extract text PDF tables locally; preserve text/precision and require review."""
    try:
        import pdfplumber
    except ImportError as exc:
        raise RuntimeError('PDF conversion needs its dependencies. Run bash Install_Update_Mac.command in the app folder, then restart.') from exc
    source, output = Path(source), Path(output)
    if source.suffix.lower() != '.pdf' or output.suffix.lower() != '.xlsx':
        raise ValueError('Choose a PDF input and a new .xlsx output.')
    if output.exists():
        raise FileExistsError('Output already exists. Choose a new filename; existing workbooks are preserved.')
    wb = Workbook(); review = wb.active; review.title = 'Conversion Review'
    review.append(['Status', 'REVIEW REQUIRED — extraction is not a validated bank/tax import'])
    review.append(['Source PDF', source.name])
    review.append(['Next steps', 'Compare every page, balances and totals with the PDF; correct dates/amounts and map columns.'])
    review.append(['Tax returns', 'Tax input still requires the original V5.1 .xlsx headings and layout.'])
    review.append(['Finance', 'Extracted rows support manual review; automated reconciliation is not implemented.'])
    review.append(['Page', 'Extracted tables', 'Text fallback / review notes'])
    tables = text_pages = 0
    def write_text(ws, values):
        for target_row, row in enumerate(values, 1):
            if len(row) > 16384 or target_row > 1048576:
                raise ValueError('Extraction exceeds Excel limits; convert a smaller PDF.')
            for col, value in enumerate(row, 1):
                text = '' if value is None else str(value)
                if len(text) > 32767:
                    raise ValueError('Extracted text exceeds the Excel cell limit; split the PDF.')
                cell = ws.cell(target_row, col, text)
                cell.data_type = 's'  # preserve literal text, including strings beginning with '='
    with pdfplumber.open(source) as pdf:
        for number, page in enumerate(pdf.pages, 1):
            extracted = page.extract_tables()
            usable = [table for table in extracted if any(any(v for v in row) for row in table)]
            text = page.extract_text() or ''
            if not text.strip() and not usable:
                raise ValueError(f'Page {number} has no extractable text/table. A scanned or blank page needs OCR/manual review; no output was saved.')
            for index, table in enumerate(usable, 1):
                ws = wb.create_sheet(f'P{number} Table {index}')
                write_text(ws, table); ws.freeze_panes = 'A2'
                tables += 1
            # Always retain page text for checking rows outside detected tables.
            if text.strip():
                ws = wb.create_sheet(f'P{number} Text')
                write_text(ws, [[line] for line in text.splitlines()])
                text_pages += 1
            review.append([number, len(usable), 'Compare with original PDF; text copy included' if text.strip() else 'Review table-only page'])
    if not tables and not text_pages:
        raise ValueError('No extractable pages found; no output was saved.')
    # Stage before publishing; exclusive link prevents replacing an existing workbook.
    staged = None
    try:
        with tempfile.NamedTemporaryFile(dir=output.parent, suffix='.xlsx', delete=False) as f:
            staged = Path(f.name)
        wb.save(staged)
        os.link(staged, output)
    finally:
        if staged:
            staged.unlink(missing_ok=True)
    return {'output': str(output), 'tables': tables, 'text_pages': text_pages}
