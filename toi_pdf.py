"""Port of V5.1 Module10 TOI PDF layout and multi-workbook export."""
from copy import copy
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
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


def export_toi_pdfs(company_folder, step_range='1-17', progress=None):
    start,end=parse_steps(step_range)
    root=Path(company_folder)
    toi=next((p for p in root.iterdir() if p.is_dir() and p.name.casefold()=='toi'),None)
    if toi is None:
        raise ValueError('The company folder has no TOI folder.')
    files=sorted(p for p in toi.glob('*') if p.suffix.lower()=='.xlsx' and not p.name.startswith(('.','~$')))
    if not files:
        raise ValueError('No .xlsx workbooks found directly inside TOI.')
    soffice=find_soffice()
    destination=root/'TOI_PDFs';destination.mkdir(exist_ok=True)
    results=[]
    for index,source in enumerate(files,1):
        try:
            with tempfile.TemporaryDirectory(prefix='toi-pdf-') as temp:
                temp=Path(temp);prepared=temp/'prepared.xlsx';out=temp/'pdf';out.mkdir()
                selected=prepare_workbook(source,prepared,start,end)
                profile=(temp/'profile').as_uri()
                command=[soffice,'-env:UserInstallation='+profile,'--headless','--convert-to','pdf:calc_pdf_Export',
                         '--outdir',str(out),str(prepared)]
                completed=subprocess.run(command,capture_output=True,text=True,timeout=180,check=False)
                pdf=out/'prepared.pdf'
                if completed.returncode or not pdf.is_file() or pdf.stat().st_size<5 or pdf.read_bytes()[:5]!=b'%PDF-':
                    raise RuntimeError('LibreOffice did not produce a PDF. '+(completed.stderr or completed.stdout).strip()[:500])
                base=f'{source.stem}_TOI_Report_{start}-{end}'
                target=destination/(base+'.pdf');version=2
                # Exclusive creation preserves previous PDFs, including another export running concurrently.
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
