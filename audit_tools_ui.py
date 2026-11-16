"""Audit Tools tabs, kept independent of the tax calculation UI."""
from pathlib import Path
import tkinter as tk
from tkinter import ttk,scrolledtext
import audit_documents as documents


class AuditToolsPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg='#f0f4f8')
        self.app=app;self.loaded_folder=None;self.last_meeting=None;self.last_request=None
        self.company=[];self.item_vars=[];self.fields={}
        tabs=self.tabs=ttk.Notebook(self);tabs.pack(fill='both',expand=True)
        request=ttk.Frame(tabs);meeting=ttk.Frame(tabs)
        tabs.add(request,text='Document Request');tabs.add(meeting,text='Meeting Notes')
        bar=ttk.Frame(request);bar.pack(fill='x',pady=5)
        ttk.Button(bar,text='Create / Open Excel',command=lambda:self.create_request(False)).pack(side='left',padx=4)
        ttk.Button(bar,text='Preview / Print PDF',command=lambda:self.create_request(True)).pack(side='left',padx=4)
        ttk.Button(bar,text='Clear selection',command=self.clear_items).pack(side='left',padx=4)
        self.request_status=ttk.Label(request,text='Uses Info/Template.xlsx · select up to 17 items',wraplength=800)
        self.request_status.pack(fill='x',padx=8,pady=4)
        self.company_label=ttk.Label(request,text='',wraplength=850);self.company_label.pack(fill='x',padx=8,pady=5)
        self.item_canvas=tk.Canvas(request,highlightthickness=0,bg='white')
        scroll=ttk.Scrollbar(request,orient='vertical',command=self.item_canvas.yview)
        scroll.pack(side='right',fill='y');self.item_canvas.pack(fill='both',expand=True)
        self.item_canvas.configure(yscrollcommand=scroll.set)
        self.items_frame=ttk.Frame(self.item_canvas)
        window=self.item_canvas.create_window((0,0),window=self.items_frame,anchor='nw')
        self.item_canvas.bind('<Configure>',lambda e:self.item_canvas.itemconfigure(window,width=e.width))
        self.items_frame.bind('<Configure>',lambda e:self.item_canvas.configure(scrollregion=self.item_canvas.bbox('all')))
        toolbar=ttk.Frame(meeting);toolbar.pack(fill='x',pady=5)
        ttk.Button(toolbar,text='Save Word',command=self.save_meeting).pack(side='left',padx=4)
        ttk.Button(toolbar,text='Open Word / Print',command=self.open_meeting).pack(side='left',padx=4)
        self.meeting_status=ttk.Label(toolbar,text='Fill the note, then Save Word.');self.meeting_status.pack(side='left',padx=10)
        canvas=self.meeting_canvas=tk.Canvas(meeting,highlightthickness=0,bg='white')
        y=ttk.Scrollbar(meeting,orient='vertical',command=canvas.yview);y.pack(side='right',fill='y')
        canvas.pack(fill='both',expand=True);canvas.configure(yscrollcommand=y.set)
        form=ttk.Frame(canvas);window2=canvas.create_window((0,0),window=form,anchor='nw')
        canvas.bind('<Configure>',lambda e:canvas.itemconfigure(window2,width=e.width))
        form.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')))
        labels=(('title','Title',2),('company_info','Company Info (auto-filled; editable)',6),
                ('meeting_details','Meeting date / time / place',2),('reference','Legal / other reference',3),
                ('notice','Audit notice / reference',2),('auditors','Auditors / attendees',3),
                ('representatives','Company representatives',3),('purpose','Purpose / agenda',3),
                ('discussion','Discussion / findings',8),('requested_documents','Requested documents',5),
                ('actions','Decisions / actions / follow-up',5),('closing','Closing / end time',3))
        for key,label,height in labels:
            ttk.Label(form,text=label).pack(fill='x',padx=12,pady=(8,2))
            text=scrolledtext.ScrolledText(form,height=height,wrap='word',font=('Helvetica',11),undo=True)
            text.pack(fill='x',expand=True,padx=12,pady=(0,3));self.fields[key]=text
        self.reset()

    def reset(self):
        self.loaded_folder=None;self.last_meeting=None;self.last_request=None;self.company=[]
        self.tabs.select(0)
        for widget in self.items_frame.winfo_children():widget.destroy()
        self.item_vars=[];self.company_label.config(text='')
        for field in self.fields.values():
            field.delete('1.0','end');field.edit_reset()
        self.item_canvas.yview_moveto(0);self.meeting_canvas.yview_moveto(0)
        self.request_status.config(text='Select a company folder with Info/Template.xlsx.')
        self.meeting_status.config(text='Fill the note, then Save Word.')

    def load(self):
        folder=self.app.folder.get()
        if self.loaded_folder==folder:return
        self.reset()
        self.app.run_bg(self.load_job,folder)

    def load_job(self,folder):
        data={'folder':folder,'company':[],'items':[],'info_error':'','template_error':''}
        try:data['company']=documents.company_info(folder)
        except Exception as exc:data['info_error']=str(exc)
        try:data['items']=documents.request_items(folder)
        except Exception as exc:data['template_error']=str(exc)
        return ('audit_tools_loaded',data)

    def populate(self,data):
        self.loaded_folder=data['folder'];self.company=data['company']
        self.company_label.config(text='\n'.join(f'{a}: {b}' for a,b in self.company) or data['info_error'])
        self.request_status.config(text=data['template_error'] or 'Select up to 17 requested documents, then create the report.')
        for widget in self.items_frame.winfo_children():widget.destroy()
        self.item_vars=[]
        for index,item in enumerate(data['items']):
            variable=tk.BooleanVar(value=False);self.item_vars.append(variable)
            tk.Checkbutton(self.items_frame,text=f'{index+1}. {item}',variable=variable,
                           anchor='w',justify='left',wraplength=max(250,self.item_canvas.winfo_width()-60),
                           bg='white',font=('Helvetica',11)).pack(fill='x',padx=10,pady=4)
        for key,value in documents.default_meeting(self.company).items():
            self.fields[key].delete('1.0','end');self.fields[key].insert('1.0',value);self.fields[key].edit_reset()
        self.meeting_status.config(text=data['info_error'] or 'Company Info loaded. Save Word before using Refresh.')

    def clear_items(self):
        for variable in self.item_vars:variable.set(False)

    def create_request(self,pdf):
        if self.app.busy:return
        selected=[i for i,v in enumerate(self.item_vars) if v.get()]
        if not selected or len(selected)>17:
            self.request_status.config(text='Select between 1 and 17 items.');return
        folder=self.app.folder.get()
        renderer='excel' if self.app.toi_renderer.get()=='Microsoft Excel' else 'libreoffice'
        self.app.run_bg(self.request_job,folder,selected,pdf,renderer)

    @staticmethod
    def request_job(folder,selected,pdf,renderer):
        path=documents.request_pdf(folder,selected,renderer) if pdf else documents.build_request(folder,selected)
        return ('audit_document_saved','request',str(path))

    def save_meeting(self):
        if self.app.busy:return
        values={key:text.get('1.0','end-1c') for key,text in self.fields.items()}
        self.app.run_bg(self.meeting_job,self.app.folder.get(),values)

    @staticmethod
    def meeting_job(folder,values):
        return ('audit_document_saved','meeting',str(documents.save_meeting(folder,values)))

    def open_meeting(self):
        if self.last_meeting:self.app.open_if(self.last_meeting)
        else:self.meeting_status.config(text='Save Word first, then open it for review/printing.')

    def saved(self,kind,path):
        if kind=='meeting':
            self.last_meeting=path;self.meeting_status.config(text='Saved: '+Path(path).name)
        else:
            self.last_request=path;self.request_status.config(text='Saved: '+Path(path).name)
            self.app.open_if(path)
