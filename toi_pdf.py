"""Port of V5.1 Module10 TOI PDF layout and multi-workbook export."""
from copy import copy
from contextlib import contextmanager, nullcontext
from pathlib import Path
import hashlib
import re
import shutil
import subprocess
import sys
import tempfile
import warnings
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.properties import PageSetupProperties

# step: row height, A-D widths before VBA's /0.2135 conversion, orientation, pages tall, final column
STEP_LAYOUTS = {
    1:(23,(10,5,7,7),'landscape',1,'D'), 2:(22,(1,7,4,5),'landscape',1,'H'),
    3:(22,(9,4,3,10),'landscape',1,'E'), 4:(22,(10,2,6,6),'portrait',2,'D'),
    5:(22,(15,2,6,6),'portrait',1,'D'), 6:(22,(15,2,6,6),'portrait',1,'D'),
    7:(22,(15,2,6,6),'portrait',1,'D'), 8:(22,(17,2,6,6),'portrait',2,'D'),
    9:(20,(10,10,10,10),'landscape',1,'G'), 10:(35,(2,8,8,9),'landscape',1,'L'),
    11:(35,(7,7,10,10),'landscape',1,'E'), 12:(35,(8,4,4,7),'landscape',1,'I'),
    13:(40,(2,7,7,7),'landscape',1,'F'), 14:(28,(2,10,10,10),'landscape',1,'E'),
    15:(30,(5,10,10,12),'portrait',6,'F'), 16:(22,(10,5,5,5),'portrait',1,'D'),
    17:(22,(10,10,10,10),'landscape',1,'D'),
}


def parse_steps(value):
    match = re.fullmatch(r'\s*(\d+)\s*-\s*(\d+)\s*', str(value))
    if not match:
        raise ValueError('Enter a step range such as 1-17 or 4-4.')
    start,end = map(int,match.groups())
    if not 1 <= start <= end <= 17:
        raise ValueError('Step range must be within 1–17, with start before end.')
    return start,end


def find_soffice():
    found = shutil.which('soffice') or shutil.which('libreoffice')
    if found:
        return found
    mac = Path('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    if mac.is_file():
        return str(mac)
    raise RuntimeError('TOI PDF export requires LibreOffice. Install LibreOffice in Applications, then retry. Source Excel files are unchanged.')


def select_renderer(renderer='auto'):
    if renderer not in ('auto','excel','libreoffice'):
        raise ValueError('Choose Microsoft Excel or LibreOffice as the PDF renderer.')
    return ('excel' if sys.platform in ('win32','darwin') else 'libreoffice') if renderer=='auto' else renderer


@contextmanager
def excel_application():
    """Own a separate Excel instance; initialize COM on the Windows worker thread."""
    if sys.platform not in ('win32','darwin'):
        raise RuntimeError('Microsoft Excel automation needs Windows or Mac. Choose LibreOffice on this platform.')
    app=com=None
    try:
        try:
            if sys.platform=='win32':
                import pythoncom
                com=pythoncom
                com.CoInitialize()
            import xlwings as xw
            app=xw.App(visible=False,add_book=False)
            app.display_alerts=False
            app.screen_updating=False
        except ImportError as exc:
            raise RuntimeError('Excel PDF export needs the Python Excel bridge. Run Install_Update_Windows.bat or bash Install_Update_Mac.command once, then restart. No Excel add-in is required.') from exc
        except Exception as exc:
            raise RuntimeError('Cannot start Microsoft Excel for PDF export. Open desktop Excel once and finish setup. On Mac, allow the launching app to control Excel in Privacy & Security > Automation. '+str(exc)) from exc
        yield app
    finally:
        try:
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    warnings.warn('The export Excel instance could not close automatically; close its unused window manually.',RuntimeWarning)
        finally:
            if com is not None:
                com.CoUninitialize()


def render_excel(app, prepared, pdf, selected):
    book=None
    try:
        book=app.books.open(str(prepared),update_links=False,read_only=True)
        for name in selected:
            try:
                book.sheets[name].range('E:I').columns.autofit()
            except Exception:
                pass  # Module10 also tolerates AutoFit on protected/unsupported columns.
        app.calculate()
        book.to_pdf(path=str(pdf),include=selected,quality='standard',show=False)
    finally:
        if book is not None:
            book.close()  # discard temporary edits; never open the original in Excel


def toi_workbook_fonts(company_folder):
    root=Path(company_folder)
    toi=next((p for p in root.iterdir() if p.is_dir() and p.name.casefold()=='toi'),None)
    if toi is None:
        raise ValueError('No TOI folder found.')
    rows=[]
    for source in sorted(toi.glob('*')):
        if source.suffix.lower()!='.xlsx' or source.name.startswith(('.','~$')):
            continue
        try:
            wb=load_workbook(source,read_only=True,keep_links=False)
            try:
                names=sorted({f.name for f in wb._fonts if f.name})
            finally:
                wb.close()
            rows.append([source.name,', '.join(names),
                         'Workbook font references only; install matching fonts before export.'])
        except Exception as exc:
            rows.append([source.name,'','ERROR: '+str(exc).splitlines()[0]])
    return rows


def prepare_workbook(source, output, start, end):
    wb = load_workbook(source, keep_links=False)
    selected = []
    for ws in wb.worksheets:
        match = re.fullmatch(r'step(\d+)',ws.title,re.I)
        number = int(match.group(1)) if match else 0
        if not start <= number <= end:
            ws.sheet_state = 'hidden'
            continue
        selected.append(ws)
        ws.sheet_state = 'visible'
        height,widths,orientation,pages,column = STEP_LAYOUTS[number]
        last = max((row[0].row for row in ws.iter_rows(min_col=1,max_col=1)
                    if row[0].value not in (None,'')), default=1)
        for r in range(1,last+1):
            ws.row_dimensions[r].height=height
        for c,width in zip('ABCD',widths):
            ws.column_dimensions[c].width=width/0.2135
        for c in range(5,10):
            letter = chr(64+c)
            # Approximate VBA AutoFit; actual glyph metrics differ by renderer/font.
            length = max((len(str(ws.cell(r,c).value or '')) for r in range(1,last+1)),default=0)
            ws.column_dimensions[letter].width=min(255,max(8,length+2))
        for row in ws.iter_rows(min_row=1,max_row=last,max_col=4):
            for cell in row:
                alignment=copy(cell.alignment)
                alignment.horizontal='left' if cell.column==1 else 'center'
                cell.alignment=alignment
        ws.print_area=f'A1:{column}{last}'
        ws.page_setup.orientation=orientation
        ws.page_setup.paperSize=ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth=1
        ws.page_setup.fitToHeight=pages
        ws.page_setup.scale=None
        if ws.sheet_properties.pageSetUpPr is None:
            ws.sheet_properties.pageSetUpPr=PageSetupProperties()
        ws.sheet_properties.pageSetUpPr.fitToPage=True
        ws.page_margins=PageMargins(left=.5,right=.5,top=.5,bottom=.5,header=.3,footer=.3)
        ws.print_options.horizontalCentered=True
        ws.print_options.verticalCentered=False
    if not selected:
        raise ValueError(f'No step{start}–step{end} worksheets found in this workbook.')
    for property_name in ('creator','lastModifiedBy','title','subject','description','keywords','category','identifier','contentStatus','version'):
        setattr(wb.properties,property_name,None)
    wb.active=wb.index(selected[0])
    wb.save(output); wb.close()
    return [ws.title for ws in selected]


@contextmanager
def export_workspace(destination,source,backend):
    """Stable Mac work paths retain Excel grants; other platforms use temporary copies."""
    if backend!='excel' or sys.platform!='darwin':
        with tempfile.TemporaryDirectory(prefix='toi-pdf-') as temp:
            temp=Path(temp);out=temp/'pdf';out.mkdir()
            yield temp/'prepared.xlsx',out/'prepared.pdf',temp
        return
    import fcntl
    directory=destination/'.excel-work'
    marker=directory/'.tax-audit-owner'
    existed=directory.exists()
    directory.mkdir(exist_ok=True)
    owner='Tax Audit TOI PDF workspace v1'
    if not existed:
        with marker.open('x',encoding='utf-8') as stream:
            stream.write(owner)
    if not marker.is_file() or marker.read_text(encoding='utf-8')!=owner:
        raise RuntimeError('Existing .excel-work folder is not a recognized app workspace; preserve it and rename it before retrying.')
    with (directory/'.export.lock').open('a') as lock:
        try:
            fcntl.flock(lock,fcntl.LOCK_EX|fcntl.LOCK_NB)
        except BlockingIOError as exc:
            raise RuntimeError('Another TOI export is using this PDF work folder. Retry after it finishes.') from exc
        try:
            key=hashlib.sha256(source.name.encode('utf-8')).hexdigest()
            prepared=directory/(key+'.xlsx')
            # The PDF folder is stable and visible when Excel asks for folder access.
            pdf=destination/('.working-'+key+'.pdf')
            pdf.unlink(missing_ok=True)
            yield prepared,pdf,None
        finally:
            fcntl.flock(lock,fcntl.LOCK_UN)


def export_toi_pdfs(company_folder, step_range='1-17', progress=None, renderer='auto'):
    start,end=parse_steps(step_range)
    root=Path(company_folder)
    toi=next((p for p in root.iterdir() if p.is_dir() and p.name.casefold()=='toi'),None)
    if toi is None:
        raise ValueError('The company folder has no TOI folder.')
    files=sorted(p for p in toi.glob('*') if p.suffix.lower()=='.xlsx' and not p.name.startswith(('.','~$')))
    if not files:
        raise ValueError('No .xlsx workbooks found directly inside TOI.')
    backend=select_renderer(renderer)
    soffice=find_soffice() if backend=='libreoffice' else None
    destination=root/'TOI_PDFs';destination.mkdir(exist_ok=True)
    results=[]
    with (excel_application() if backend=='excel' else nullcontext(None)) as app:
        for index,source in enumerate(files,1):
            try:
                with export_workspace(destination,source,backend) as (prepared,pdf,temp):
                    selected=prepare_workbook(source,prepared,start,end)
                    if backend=='excel':
                        render_excel(app,prepared,pdf,selected)
                    else:
                        profile=(temp/'profile').as_uri()
                        command=[soffice,'-env:UserInstallation='+profile,'--headless','--convert-to','pdf:calc_pdf_Export',
                                 '--outdir',str(pdf.parent),str(prepared)]
                        completed=subprocess.run(command,capture_output=True,text=True,timeout=180,check=False)
                        if completed.returncode:
                            raise RuntimeError('LibreOffice export failed: '+(completed.stderr or completed.stdout).strip()[:500])
                    if not pdf.is_file() or pdf.stat().st_size<5 or pdf.read_bytes()[:5]!=b'%PDF-':
                        raise RuntimeError(f'{backend} did not produce a valid PDF.')
                    base=f'{source.stem}_TOI_Report_{start}-{end}'
                    target=destination/(base+'.pdf');version=2
                    while True:
                        try:
                            with target.open('xb') as output:
                                try:
                                    with pdf.open('rb') as stream:
                                        shutil.copyfileobj(stream,output)
                                except BaseException:
                                    target.unlink(missing_ok=True)
                                    raise
                            break
                        except FileExistsError:
                            target=destination/f'{base}_{version}.pdf';version+=1
                    results.append([source.name,', '.join(selected),'EXPORTED',str(target)])
            except Exception as exc:
                results.append([source.name,f'{start}-{end}','ERROR: '+str(exc).splitlines()[0],''])
            if progress:
                progress(index,len(files))
    return results
