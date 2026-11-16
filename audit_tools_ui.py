"""Tax Audit Tools tabs, kept independent of the tax calculation UI."""
from pathlib import Path
import tkinter as tk
from tkinter import ttk,scrolledtext
import audit_documents as documents
from tax_law_ui import TaxLawPanel

TEXT,WHITE,LINE='#172b43','#ffffff','#d7e1eb'


class AuditToolsPanel(tk.Frame):
    def __init__(self,parent,app):
        super().__init__(parent,bg='#f0f4f8')
        self.app=app;self.loaded_folder=None;self.last_meeting=None;self.last_request=None
        self.company=[];self.item_vars=[];self.item_labels=[];self.fields={}
        tabs=self.tabs=ttk.Notebook(self);tabs.pack(fill='both',expand=True)
        request=ttk.Frame(tabs);meeting=ttk.Frame(tabs)
        tabs.add(request,text='Document Request');tabs.add(meeting,text='Meeting Notes')
        self.law_panel=TaxLawPanel(tabs);tabs.add(self.law_panel,text='Tax Law Search')
        bar=ttk.Frame(request);bar.pack(fill='x',pady=5)
        ttk.Button(bar,text='Save / Open Excel',command=lambda:self.create_request(False)).pack(side='left',padx=4)
        ttk.Button(bar,text='Save / Preview PDF',command=lambda:self.create_request(True)).pack(side='left',padx=4)
        ttk.Button(bar,text='Clear selection',command=self.clear_items).pack(side='left',padx=4)
        self.request_status=ttk.Label(request,text='Select up to 17 items · saving replaces the current document request',wraplength=800)
        self.request_status.pack(fill='x',padx=8,pady=4)
        self.company_label=ttk.Label(request,text='',wraplength=850);self.company_label.pack(fill='x',padx=8,pady=5)
        self.item_canvas=tk.Canvas(request,highlightthickness=0,bg='white')
        scroll=ttk.Scrollbar(request,orient='vertical',command=self.item_canvas.yview)
        scroll.pack(side='right',fill='y');self.item_canvas.pack(fill='both',expand=True)
        self.item_canvas.configure(yscrollcommand=scroll.set)
        self.items_frame=ttk.Frame(self.item_canvas)
        self.item_window=self.item_canvas.create_window((0,0),window=self.items_frame,anchor='nw')
        self.item_canvas.bind('<Configure>',self.resize_items)
        self.items_frame.bind('<Configure>',lambda e:self.item_canvas.configure(scrollregion=self.item_canvas.bbox('all')))
        self.bind_scroll(self.item_canvas,self.item_canvas)
        self.bind_scroll(self.items_frame,self.item_canvas)
        toolbar=ttk.Frame(meeting);toolbar.pack(fill='x',pady=5)
        ttk.Button(toolbar,text='Save Word',command=self.save_meeting).pack(side='left',padx=4)
        ttk.Button(toolbar,text='Open Word / Print',command=self.open_meeting).pack(side='left',padx=4)
        self.meeting_status=ttk.Label(toolbar,text='Fill the note, then Save Word.');self.meeting_status.pack(side='left',padx=10)
        panes=self.meeting_panes=ttk.Panedwindow(meeting,orient='horizontal')
        panes.pack(fill='both',expand=True)
        details=ttk.Frame(panes,width=380);writing=ttk.Frame(panes,width=800)
        panes.add(details,weight=1);panes.add(writing,weight=3)
        self.meeting_sized=False;panes.bind('<Configure>',self.size_meeting_panes)
        canvas=self.meeting_canvas=tk.Canvas(details,highlightthickness=0,bg=WHITE,width=360)
        y=ttk.Scrollbar(details,orient='vertical',command=canvas.yview);y.pack(side='right',fill='y')
        canvas.pack(fill='both',expand=True);canvas.configure(yscrollcommand=y.set)
        form=ttk.Frame(canvas);window2=canvas.create_window((0,0),window=form,anchor='nw')
        canvas.bind('<Configure>',lambda e:canvas.itemconfigure(window2,width=e.width))
        form.bind('<Configure>',lambda e:canvas.configure(scrollregion=canvas.bbox('all')))
        self.bind_scroll(canvas,canvas);self.bind_scroll(form,canvas)
        labels=(('title','Title',2),('company_info','Company Info (auto-filled; editable)',6),
                ('meeting_date','Meeting date (YYYY-MM-DD)',1),('meeting_details','Time / place',2),
                ('reference','Legal / other reference',3),
                ('notice','Audit notice / reference',2),('auditors','Auditors / attendees',3),
                ('representatives','Company representatives',3),('purpose','Purpose / agenda',3))
        for key,label,height in labels:
            caption=ttk.Label(form,text=label);caption.pack(fill='x',padx=10,pady=(8,2))
            self.bind_scroll(caption,canvas)
            text=self.add_field(form,key,height)
            text.pack(fill='x',expand=True,padx=10,pady=(0,3))
        editor_tabs=self.editor_tabs=ttk.Notebook(writing);editor_tabs.pack(fill='both',expand=True,padx=(8,0))
        note=ttk.Frame(editor_tabs);followup=ttk.Frame(editor_tabs)
        editor_tabs.add(note,text='Writing note');editor_tabs.add(followup,text='Requests & follow-up')
        ttk.Label(note,text='Meeting note / discussion / findings').pack(fill='x',padx=10,pady=(10,4))
        self.add_field(note,'discussion',24).pack(fill='both',expand=True,padx=10,pady=(0,10))
        for key,label,height in (('requested_documents','Requested documents',5),
                                 ('actions','Decisions / actions / follow-up',6),('closing','Closing / end time',2)):
            ttk.Label(followup,text=label).pack(fill='x',padx=10,pady=(8,2))
            self.add_field(followup,key,height).pack(fill='both',expand=True,padx=10,pady=(0,6))
        self.reset()

    def add_field(self,parent,key,height):
        field=scrolledtext.ScrolledText(parent,height=height,width=25,wrap='word',font=('Helvetica',12),undo=True,
                                       bg=WHITE,fg=TEXT,insertbackground=TEXT,selectbackground='#087f8c',selectforeground=WHITE,
                                       padx=8,pady=6,spacing1=3,spacing3=3)
        self.fields[key]=field
        return field

    @staticmethod
    def bind_scroll(widget,canvas):
        def scroll(event):
            if getattr(event,'num',None) in (4,5):units=-1 if event.num==4 else 1
            else:
                delta=event.delta
                units=-int(delta/120) if abs(delta)>=120 else (-1 if delta>0 else 1)
            canvas.yview_scroll(units,'units')
            return 'break'
        for event in ('<MouseWheel>','<Button-4>','<Button-5>'):widget.bind(event,scroll)

    def resize_items(self,event):
        self.item_canvas.itemconfigure(self.item_window,width=event.width)
        for label in self.item_labels:label.config(wraplength=max(180,event.width-90))

    def size_meeting_panes(self,event):
        if not self.meeting_sized and event.width>100:
            self.meeting_panes.sashpos(0,max(280,int(event.width*.30)))
            self.meeting_sized=True

    def reset(self,reset_law=True):
        self.loaded_folder=None;self.last_meeting=None;self.last_request=None;self.company=[]
        self.tabs.select(0);self.editor_tabs.select(0)
        for widget in self.items_frame.winfo_children():widget.destroy()
        self.item_vars=[];self.item_labels=[];self.company_label.config(text='')
        for field in self.fields.values():
            field.delete('1.0','end');field.edit_reset()
        self.item_canvas.yview_moveto(0);self.meeting_canvas.yview_moveto(0)
        self.request_status.config(text='Select a company folder with Info/Template.xlsx.')
        self.meeting_status.config(text='Fill the note, then Save Word.')
        if reset_law:self.law_panel.reset()

    def load(self):
        folder=self.app.folder.get()
        if not folder.strip() or not Path(folder).is_dir():
            if self.loaded_folder is not None:self.reset(reset_law=False)
            self.request_status.config(text='Select a company folder with Info/Template.xlsx to create document requests.')
            self.meeting_status.config(text='Select a company folder to fill and save a meeting note.')
            self.tabs.select(self.law_panel)
            self.app.status('Tax Law Search is ready. Select a company folder when you need company documents.')
            return
        if self.loaded_folder==folder:return
        self.reset(reset_law=False)
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
        self.request_status.config(text=data['template_error'] or 'Select up to 17 items · saving replaces Doc_Requirement.xlsx / .pdf')
        for widget in self.items_frame.winfo_children():widget.destroy()
        self.item_vars=[];self.item_labels=[]
        for index,item in enumerate(data['items']):
            variable=tk.BooleanVar(value=False);self.item_vars.append(variable)
            background=WHITE if index%2==0 else '#edf3f8'
            row=tk.Frame(self.items_frame,bg=background,highlightthickness=1,highlightbackground=LINE)
            row.pack(fill='x',padx=8,pady=3);row.columnconfigure(1,weight=1)
            checkbox=tk.Checkbutton(row,variable=variable,bg=background,fg=TEXT,activebackground=background,
                                    activeforeground=TEXT,selectcolor=WHITE,highlightbackground=background)
            checkbox.grid(row=0,column=0,sticky='nw',padx=(5,0),pady=6)
            label=tk.Label(row,text=f'{index+1}. {item}',bg=background,fg=TEXT,anchor='w',justify='left',
                           font=('Helvetica',12),wraplength=max(180,self.item_canvas.winfo_width()-90),cursor='hand2')
            label.grid(row=0,column=1,sticky='ew',padx=(5,12),pady=8);self.item_labels.append(label)
            label.bind('<Button-1>',lambda event,v=variable:v.set(not v.get()))
            for widget in (row,checkbox,label):self.bind_scroll(widget,self.item_canvas)
        for key,value in documents.default_meeting(self.company).items():
            self.fields[key].delete('1.0','end');self.fields[key].insert('1.0',value);self.fields[key].edit_reset()
        self.meeting_status.config(text=data['info_error'] or 'Company Info loaded. Save Word before using Refresh.')

    def clear_items(self):
        for variable in self.item_vars:variable.set(False)

    def create_request(self,pdf):
        if self.app.busy or not self.app.need_folder():return
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
        if self.app.busy or not self.app.need_folder():return
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
