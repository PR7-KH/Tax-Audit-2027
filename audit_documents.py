"""VBA Module11 document requests and editable Word meeting-note generation."""
from datetime import datetime,timezone,timedelta
from pathlib import Path
import os
import re
import shutil
import subprocess
import tempfile
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import load_workbook
from openpyxl.styles import Font,Alignment
import engine
from toi_pdf import select_renderer,find_soffice,excel_application,render_excel,export_workspace

HERE=Path(__file__).resolve().parent
MEETING_TEMPLATE=HERE/'assets'/'meeting_note_template.docx'
MEETING_FIELDS=('title','reference','notice','meeting_details','auditors','representatives','company_info',
                'purpose','discussion','actions','requested_documents','closing')


def company_info(folder):
    directory=engine.find_dir(Path(folder),'INFO')
    source=engine.first_xlsx(directory) if directory else None
    if source is None:
        raise ValueError('Put the company Info workbook in Info (separate from Template.xlsx).')
    wb=load_workbook(source,data_only=True,read_only=True)
    try:
        ws=engine.ci_sheet(wb,'Sheet1')
        if ws is None:raise ValueError('Company Info workbook needs Sheet1.')
        return [(str(ws.cell(r,1).value) if ws.cell(r,1).value is not None else '',
                 str(ws.cell(r,2).value) if ws.cell(r,2).value is not None else '') for r in (1,2,5,6,3)]
    finally:wb.close()


def request_template(folder):
    directory=engine.find_dir(Path(folder),'INFO')
    source=next((p for p in directory.iterdir() if p.name.casefold()=='template.xlsx'),None) if directory else None
    if source is None:raise ValueError('Document requests need Info/Template.xlsx with Sheet1 and Sheet2.')
    return source


def request_items(folder):
    wb=load_workbook(request_template(folder),data_only=True,read_only=True)
    try:
        ws=engine.ci_sheet(wb,'Sheet2')
        if engine.ci_sheet(wb,'Sheet1') is None or ws is None:raise ValueError('Template.xlsx needs Sheet1 (report) and Sheet2 (items).')
        items=[]
        for row in range(7,1001):
            value=ws.cell(row,1).value
            if value is None or str(value).strip()=='':break
            items.append(str(value).strip())
        if not items:raise ValueError('No request items found in Template.xlsx Sheet2!A7 onward.')
        return items
    finally:wb.close()


def publish_unique(directory,stem,suffix,save):
    """Stage and publish exclusively, preserving all previous user files."""
    directory=Path(directory);directory.mkdir(parents=True,exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=directory,suffix=suffix,delete=False) as stream:staged=Path(stream.name)
    try:
        save(staged)
        target=directory/(stem+suffix);version=2
        while True:
            try:
                with target.open('xb') as outgoing:
                    try:
                        with staged.open('rb') as incoming:shutil.copyfileobj(incoming,outgoing)
                        outgoing.flush();os.fsync(outgoing.fileno())
                    except BaseException:
                        target.unlink(missing_ok=True);raise
                return target
            except FileExistsError:target=directory/f'{stem}_{version}{suffix}';version+=1
    finally:staged.unlink(missing_ok=True)


def build_request(folder,indices):
    items=request_items(folder)
    selected=sorted(set(indices))
    if not selected:raise ValueError('Select at least one requested document.')
    if len(selected)>17:raise ValueError('The Module11 template has 17 slots (rows 14–30); select at most 17 items.')
    if any(not isinstance(i,int) or i<0 or i>=len(items) for i in selected):raise ValueError('Invalid request selection; reload the template.')
    details=company_info(folder)
    wb=load_workbook(request_template(folder))
    ws=engine.ci_sheet(wb,'Sheet1')
    for sheet in list(wb):
        if sheet!=ws:wb.remove(sheet)
    ws.title='Doc Requirement'
    for row,(label,value) in enumerate(details,6):
        for col,text in enumerate((label,value),1):
            cell=ws.cell(row,col,text);cell.data_type='s'
            cell.font=Font(name=engine.FONT,size=10)
            cell.alignment=Alignment(horizontal='left')
    for row in range(14,31):
        cell=ws.cell(row,1);cell.value=None
        cell.font=Font(name=engine.FONT,size=10);cell.alignment=Alignment(horizontal='left')
    for number,index in enumerate(selected,1):ws.cell(13+number,1,f'              {number}. {items[index]}')
    for row in range(16,31):ws.row_dimensions[row].hidden=ws.cell(row,1).value in (None,'')
    ws.sheet_state='visible';wb.active=0
    try:return publish_unique(Path(folder)/'Audit_Tools','Doc_Requirement','.xlsx',wb.save)
    finally:wb.close()


def request_pdf(folder,indices,renderer='auto'):
    source=build_request(folder,indices)
    destination=source.parent;backend=select_renderer(renderer)
    with export_workspace(destination,Path('Doc_Requirement.xlsx'),backend) as (prepared,pdf,temp):
        # Preserve the template's page setup. Copy into the granted work file.
        with source.open('rb') as incoming,prepared.open('wb') as outgoing:shutil.copyfileobj(incoming,outgoing)
        if backend=='excel':
            with excel_application() as app:render_excel(app,prepared,pdf,['Doc Requirement'],autofit=False)
        else:
            command=[find_soffice(),'-env:UserInstallation='+(temp/'profile').as_uri(),'--headless','--convert-to',
                     'pdf:calc_pdf_Export','--outdir',str(pdf.parent),str(prepared)]
            result=subprocess.run(command,capture_output=True,text=True,timeout=180)
            if result.returncode:raise RuntimeError('PDF renderer failed: '+result.stderr[:300])
        if not pdf.is_file() or pdf.read_bytes()[:5]!=b'%PDF-':raise RuntimeError('The PDF renderer did not produce a valid request PDF.')
        return publish_unique(destination,'Doc_Requirement','.pdf',lambda p:shutil.copyfile(pdf,p))


def default_meeting(details):
    values=dict.fromkeys(MEETING_FIELDS,'')
    values.update(title='កំណត់ហេតុស្តីពីការធ្វើសវនកម្ម',
                  meeting_details=datetime.now(timezone(timedelta(hours=7))).strftime('%Y-%m-%d'),
                  company_info='\n'.join(f'{label}: {value}' for label,value in details))
    return values


def save_meeting(folder,values):
    if not str(values.get('title','')).strip():raise ValueError('Enter a meeting-note title.')
    doc=Document(MEETING_TEMPLATE)
    for node in list(doc._element.iter(qn('w:t'))):
        text=node.text or ''
        rendered=re.sub(r'\{\{(\w+)\}\}',lambda m:str(values.get(m.group(1),'')),text)
        if rendered==text:continue
        paragraph=node
        while paragraph is not None and paragraph.tag!=qn('w:p'):paragraph=paragraph.getparent()
        if paragraph is not None and not paragraph.xpath('.//w:drawing | .//w:pict'):
            properties=paragraph.find(qn('w:pPr'))
            if properties is not None:
                numbering=properties.find(qn('w:numPr'))
                if numbering is not None:properties.remove(numbering)
                align=properties.find(qn('w:jc'))
                if align is not None:align.set(qn('w:val'),'center' if '{{title}}' in text else 'left')
        lines=rendered.split('\n');node.text=lines[0]
        node.set(qn('xml:space'),'preserve')
        current=node
        for line in lines[1:]:
            br=OxmlElement('w:br');current.addnext(br)
            new=OxmlElement('w:t');new.text=line;new.set(qn('xml:space'),'preserve');br.addnext(new);current=new
    doc.core_properties.title=str(values['title'])
    doc.core_properties.author='Tax Audit';doc.core_properties.last_modified_by='Tax Audit'
    doc.core_properties.created=datetime.now(timezone.utc)
    doc.core_properties.modified=datetime.now(timezone.utc)
    return publish_unique(Path(folder)/'Audit_Tools','Meeting_Note','.docx',doc.save)
