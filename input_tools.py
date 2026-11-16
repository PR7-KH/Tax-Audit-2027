"""Source inventory, format validation and local PDF extraction helpers."""
from collections import Counter
from pathlib import Path
import hashlib
import json
import re
import shutil
import os
import tempfile
from openpyxl import Workbook, load_workbook
import engine


def visible_files(directory):
    if directory is None:
        return []
    return sorted(p for p in directory.rglob('*') if p.is_file() and not any(
        part.startswith(('.', '~$')) for part in p.relative_to(directory).parts))


def target_folders(folder):
    """Only direct input folders recognized by the tax/finance system."""
    root = Path(folder)
    aliases = {name: name for name in engine.cons.LAYOUTS}
    aliases.update({'TOI': 'TOI', 'INFO': 'Info'})
    for label, names in engine.FINANCE_FOLDERS.values():
        aliases.update({name: label for name in names})
    return [(aliases[d.name.upper()], d, d.name.upper() in engine.cons.LAYOUTS)
            for d in sorted(root.iterdir())
            if d.is_dir() and d.name.upper() in aliases and d.resolve().is_relative_to(root.resolve())]


def target_files(folder):
    return sorted({p for _, directory, _ in target_folders(folder) for p in visible_files(directory)
                   if p.resolve().is_relative_to(Path(folder).resolve())})


def file_hash(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024*1024), b''):
            digest.update(block)
    return digest.hexdigest()


def conversion_records(folder):
    path = Path(folder)/'.tax-audit-conversions.json'
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding='utf-8'))
    if data.get('version') != 1 or not isinstance(data.get('files'), dict):
        raise ValueError('Unsupported conversion history; preserve .tax-audit-conversions.json and contact support.')
    return data['files']


def converted_output(root, source, records):
    root = Path(root).resolve(); source = Path(source).resolve()
    record = records.get(str(source.relative_to(root)))
    if not isinstance(record, dict):
        return None
    output = (root/record.get('output', '')).resolve()
    if not output.is_relative_to(root) or output.suffix.lower() != '.xlsx' or not output.is_file():
        return None
    if file_hash(source) != record.get('source_sha256'):
        return None
    return output


def folder_status(root, files, records, tax=False):
    pending = [p for p in files if p.suffix.lower()=='.pdf' and not converted_output(root,p,records)]
    if pending:
        return f'Not ready: {len(pending)} PDF(s) need conversion (select row to convert)'
    ordinary = [p for p in files if p.suffix.lower()!='.pdf']
    base = format_status(ordinary, tax=tax)
    if base.startswith('Not ready'):
        return base
    converted = len(files)-len(ordinary)
    if converted:
        return (f'Converted: {converted} Excel extraction(s) available; ' +
                ('review and V5.1 mapping required' if tax else 'review required'))
    return base


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
    records = conversion_records(root)
    for tax in engine.cons.LAYOUTS:
        directory = engine.find_dir(root, tax)
        files = visible_files(directory)
        groups = {}
        for p in files:
            rel = p.relative_to(directory)
            groups.setdefault(rel.parts[0] if len(rel.parts) > 1 else 'Root (move to year folder)', []).append(p)
        for year, paths in sorted(groups.items()):
            status = folder_status(root, paths, records, tax=True)
            if year.startswith('Root'):
                status = 'Not ready: monthly files must be under year folders'
            rows.append([tax, year, len(paths), status, format_counts(paths)])
        if not groups:
            rows.append([tax, '—', 0, 'Missing' if directory is None else 'Empty', '—'])
    for name in ('TOI', 'Info'):
        directory = engine.find_dir(root, name)
        files = visible_files(directory)
        rows.append([name, 'All', len(files), folder_status(root, files, records, tax=True), format_counts(files)])
    for key, (label, _) in engine.FINANCE_FOLDERS.items():
        files = visible_files(engine.finance_folder(root, key))
        rows.append([label, 'All', len(files), folder_status(root, files, records), format_counts(files)])
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
    records = conversion_records(root)
    for path in target_files(root):
        rel = path.relative_to(root)
        suffix = path.suffix.lower()
        if suffix == '.pdf':
            output = converted_output(root,path,records)
            result = (f'CONVERTED: {output.name}; review and mapping required' if output else
                      'NOT READY: PDF needs conversion; select row and click Convert selected; scans need OCR')
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


def selected_pdfs(folder, values=None, overview=False):
    root = Path(folder)
    pdfs = [p for p in target_files(root) if p.suffix.lower()=='.pdf']
    if values is None:
        return pdfs
    if overview:
        label, scope = str(values[0]), str(values[1])
        selected = []
        for name, directory, _ in target_folders(root):
            if label not in (name, directory.name):
                continue
            for p in pdfs:
                if not p.is_relative_to(directory):
                    continue
                rel = p.relative_to(directory)
                if scope == 'All' or (scope.startswith('Root') and len(rel.parts)==1) or (len(rel.parts)>1 and rel.parts[0]==scope):
                    selected.append(p)
        return sorted(set(selected))
    return [p for p in pdfs if p.name==str(values[2]) and str(p.relative_to(root).parent)==str(values[1])]


def view_pdfs(folder, values, key):
    if values is None:
        return selected_pdfs(folder)
    if key in ('sources','check_fin'):
        return selected_pdfs(folder,[values[0],values[1] if key=='sources' else 'All'],overview=True)
    if key.startswith('fin_'):
        directory = engine.finance_folder(folder,key.split('_',1)[1])
        source = directory/str(values[0]) if directory else None
        return [p for p in selected_pdfs(folder) if p==source]
    return selected_pdfs(folder,values)


def batch_convert_pdfs(folder, sources=None, progress=None):
    """Convert target PDFs without dialogs. Keep originals and never replace workbooks."""
    root = Path(folder).resolve()
    allowed = {p.resolve() for p in selected_pdfs(root)}
    chosen = sorted(allowed if sources is None else {Path(p).resolve() for p in sources})
    if not set(chosen) <= allowed:
        raise ValueError('Batch conversion is restricted to recognized target input folders.')
    records = conversion_records(root)
    results = []
    for index, source in enumerate(chosen, 1):
        rel = source.relative_to(root)
        try:
            existing = converted_output(root, source, records)
            if existing:
                status, output = 'ALREADY CONVERTED: Excel available; review required', existing
            else:
                # Generic extraction must not be picked up as a V5.1 monthly return.
                directory = source.parent/'.pdf-extractions'
                directory.mkdir(exist_ok=True)
                output = directory/(source.stem+'_review.xlsx')
                version = 2
                while output.exists():
                    output = directory/f'{source.stem}_review_{version}.xlsx'
                    version += 1
                before = file_hash(source)
                convert_pdf(source,output)
                if file_hash(source) != before:
                    raise ValueError('Source changed during conversion; output needs manual review and was not registered.')
                records[str(rel)] = {'source_sha256': before, 'output': str(output.relative_to(root))}
                staged = None
                try:
                    with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=root, prefix='.conversion-', delete=False) as f:
                        staged = Path(f.name)
                        json.dump({'version':1, 'files':records},f,indent=2)
                    os.replace(staged,root/'.tax-audit-conversions.json')
                finally:
                    if staged:
                        staged.unlink(missing_ok=True)
                status = 'CONVERTED: Excel available; review and mapping required'
            results.append([rel.parts[0], str(rel.parent), source.name, 0, status, str(output.relative_to(root))])
        except Exception as exc:
            results.append([rel.parts[0], str(rel.parent), source.name, 0, 'ERROR: '+str(exc).splitlines()[0], ''])
        if progress:
            progress(index,len(chosen))
    return results


def classify_download(path):
    """Classify only explicit filename tokens; leave uncertain files alone."""
    text = re.sub(r'[^a-z0-9]+',' ',Path(path).stem.casefold()).strip()
    patterns = {
        'PR': r'\b(purchases?|pr)\b', 'SR': r'\b(sales?|sr)\b',
        'TOS': r'\b(salary tax|salarytax|tos)\b', 'WHT': r'\b(withholding tax|withholdingtax|wht)\b',
        'GL': r'\b(gl|general ledger)\b', 'TB': r'\b(tb|trial balance)\b',
        'BANK': r'\b(bank|bank statement)\b', 'AR': r'\b(ar|receivables?|a r)\b',
        'EXPENSE': r'\b(expenses?|purchase ledger)\b', 'TOI':r'\btoi\b', 'Info':r'\b(company info|info)\b',
    }
    found=[name for name,pattern in patterns.items() if re.search(pattern,text)]
    # Explicit purchase-ledger files are finance, not tax purchases.
    if 'EXPENSE' in found and 'purchase ledger' in text and 'PR' in found:
        found.remove('PR')
    if len(found)!=1:
        return None,None,'Ambiguous filename' if found else 'Unrecognized filename'
    target=found[0]
    year=None
    if target in engine.cons.LAYOUTS:
        years=set(re.findall(r'(?<!\d)(20\d{2})(?!\d)',Path(path).stem))
        if len(years)!=1:
            return None,None,'Tax filename must identify one year (20xx)'
        year=next(iter(years))
    return target,year,''


def organize_downloads(folder):
    """First-selection setup only. Existing target folders disable this workflow."""
    root=Path(folder).resolve()
    if target_folders(root):
        return []
    downloads=[p for p in root.iterdir() if p.is_dir() and p.name.casefold()=='download']
    if len(downloads)!=1 or not downloads[0].resolve().is_relative_to(root):
        return []
    download=downloads[0]
    reports=[]
    journal=root/'.tax-audit-imports.jsonl'
    for source in visible_files(download):
        rel=source.relative_to(root)
        if source.is_symlink() or not source.resolve().is_relative_to(download.resolve()):
            reports.append([str(rel),'','SKIPPED: linked file outside Download scope']);continue
        if source.suffix.lower() not in engine.FIN_EXT:
            reports.append([str(rel),'','SKIPPED: unsupported format']);continue
        target,year,reason=classify_download(source)
        if not target:
            reports.append([str(rel),'','SKIPPED: '+reason]);continue
        directory=root/target
        if year:
            directory/=year
        destination=directory/source.name
        created=False
        try:
            directory.mkdir(parents=True,exist_ok=True)
            if any(p.name.casefold()==source.name.casefold() for p in directory.iterdir()):
                reports.append([str(rel),str(destination.relative_to(root)),'SKIPPED: destination already exists']);continue
            digest=file_hash(source)
            record={'source':str(rel),'destination':str(destination.relative_to(root)),'sha256':digest}
            with journal.open('a',encoding='utf-8') as log:
                log.write(json.dumps(dict(record,status='planned'))+'\n');log.flush();os.fsync(log.fileno())
            with destination.open('xb') as output:
                created=True
                with source.open('rb') as stream:
                    shutil.copyfileobj(stream,output)
                output.flush();os.fsync(output.fileno())
            if file_hash(destination)!=digest or file_hash(source)!=digest:
                raise ValueError('File changed during transfer; original retained')
            shutil.copystat(source,destination)
            source.unlink()
            with journal.open('a',encoding='utf-8') as log:
                log.write(json.dumps(dict(record,status='moved'))+'\n')
            reports.append([str(rel),str(destination.relative_to(root)),'MOVED'])
        except Exception as exc:
            if created and source.exists():
                destination.unlink(missing_ok=True)
            reports.append([str(rel),str(destination.relative_to(root)),'ERROR: '+str(exc).splitlines()[0]])
    return reports
