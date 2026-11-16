"""Tax and Financial Control Center: one window for tax audit and financial audit.

Click a report to VIEW it. Nothing is written to disk until you press an Export button.
"""
from collections import OrderedDict
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
import subprocess as _sound_subprocess
from urllib.parse import urlsplit
from urllib.request import url2pathname
from tkinter import filedialog, messagebox, ttk

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import calc  # noqa: E402
import engine  # noqa: E402
from input_tools import (source_inventory, validate_sources, convert_pdf, selected_pdfs,
                         batch_convert_pdfs, conversion_records, converted_output, target_files, view_pdfs, organize_downloads)
from toi_pdf import export_toi_pdfs, parse_steps, toi_workbook_fonts
from audit_tools_ui import AuditToolsPanel

NAVY, SIDE, SIDE_HOVER, ACCENT = "#101f35", "#14263f", "#203a58", "#087f8c"
GREEN, BG, CARD, TEXT, MUTED = "#167459", "#f0f4f8", "#ffffff", "#172b43", "#61748b"
LINE = "#d7e1eb"
FONT = "Helvetica"
NUMBER = re.compile(r"^-?[\d,]+(\.\d+)?%?$")

# (kind, key, label)   kind: group | item | sub
MENU = [
    ("group", "", "TAX AUDIT"),
    ("item", "dashboard", "Audit Dashboard"),
    ("item", "tois", "TOIs Overview"),
    ("item", "mtr", "MTR Overview"),
    ("item", "consolidate", "Consolidated data"),
    ("item", "tps", "Tax Paid Summary"),
    ("item", "audit_tools", "Tax Audit Tools"),
    ("group", "", "FINANCIAL AUDIT"),
    ("item", "check_fin", "Check source folders"),
    ("item", "fin_SALES", "Sales"),
    ("item", "fin_GL", "General ledger"),
    ("item", "fin_BANK", "Bank"),
    ("item", "fin_TB", "Trial balance"),
    ("item", "fin_AR", "A/R"),
    ("item", "fin_EXPENSE", "Expense"),
    ("item", "tie_out", "Tax vs ledger tie-out"),
    ("group", "", "SYSTEM"),
    ("item", "log", "Processing log"),
]
ACTION_TITLES = {"import_results": "Download import results","conversion_results": "PDF conversion results","sources": "Source overview", "source_files": "File details",
                 "check_tax": "Validate source files", "build_all": "Build all data"}
SUBTITLE = {
    "sources": "Company file counts by tax year and financial folder. No workbook contents are opened.",
    "check_tax": "Reads every monthly tax file with the same reader as V5.1 and reports problems before any build.",
    "build_all": "Builds every tax report in memory. Nothing is saved until you export.",
    "consolidate": "Consolidated monthly data with grand and year totals at the top (PR, SR, TOS, WHT).",
    "summary": "Year totals for each tax type.",
    "toi": "TOI steps 4 to 7 with Compare, % Change and Remark.",
    "fulltoi": "All TOI steps 1 to 17 in one sheet.",
    "mtr": "File-by-file detail for all four tax types.",
    "mtr_pr": "PR file list with year subtotals.", "mtr_sr": "SR file list with year subtotals.",
    "mtr_tos": "TOS file list with year subtotals.", "mtr_wht": "WHT file list with year subtotals.",
    "tps": "Tax paid by item for the two years, from the Info folder.",
    "dashboard": "Audit rules from the ADT sheet applied to TOI. Check the Processing log for corrected V5.1 numbers.",
    "check_fin": "Finds the GL, TB, Bank, Sales, A/R and Expense sub-folders inside the same company folder.",
    "tie_out": "Compares tax returns with the ledger and bank.",
    "log": "Warnings and corrections from the last build.",
}
EXPORT_NAME = {"mtr": "MTR_Overview"}


def open_path(path):
    if sys.platform == "darwin":
        subprocess.run(["open", str(path)])
    elif sys.platform.startswith("win"):
        import os
        os.startfile(str(path))  # noqa
    else:
        subprocess.run(["xdg-open", str(path)])


def report_row_tag(values, index, numbered=False):
    """Presentation only: emphasize report sections, totals and source status."""
    content = values[1:] if numbered else values
    text = [str(v).strip().casefold() for v in content if v not in (None, "")]
    if any(v.startswith(("error:", "cannot open:", "#ref!", "#value!", "#div/0!")) for v in text):
        return "error"
    if any(v in ("missing", "empty", "folder not found") or v.startswith(("not ready:", "review:")) for v in text):
        return "warning"
    if text and (text[0] in ("total", "grand total") or text[0].startswith(("subtotal ", "year "))):
        return "total"
    if text and (text[0] == "year" or text[0].startswith(("folder:", "year:"))):
        return "section"
    return "odd" if index % 2 else "even"




def mtr_source_path(wb,sheet,row):
    if not sheet or not sheet.startswith('MTR_') or sheet not in wb:
        return None
    link=wb[sheet].cell(row,2).hyperlink
    if link is None or not link.target:
        return None
    target=link.target
    parsed=urlsplit(target)
    if parsed.scheme.casefold()=='file':
        path=url2pathname(parsed.path)
        if parsed.netloc and parsed.netloc.casefold()!='localhost':
            path='//'+parsed.netloc+path
        return Path(path)
    # Accept legacy plain local paths, including Windows drive letters.
    if not parsed.scheme or (len(parsed.scheme)==1 and len(target)>2 and target[1]==':'):
        return Path(target)
    return None


class Btn(tk.Label):
    """Flat button built from a Label so colours survive macOS dark mode."""

    def __init__(self, parent, text, command, color=ACCENT, hover=NAVY, **kw):
        kw.setdefault("padx", 14)
        kw.setdefault("pady", 7)
        super().__init__(parent, text=text, bg=color, fg="white", cursor="hand2",
                         font=(FONT, 11, "bold"), **kw)
        self.color, self.hover, self.command, self.enabled = color, hover, command, True
        self.bind("<Button-1>", lambda e: self.command() if self.enabled else None)
        self.bind("<Enter>", lambda e: self.config(bg=self.hover) if self.enabled else None)
        self.bind("<Leave>", lambda e: self.config(bg=self.color if self.enabled else "#9aa7b8"))

    def set_enabled(self, on):
        self.enabled = on
        self.config(bg=self.color if on else "#9aa7b8")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Tax & Financial Control Center")
        self.geometry("1320x820")
        self.minsize(1000, 640)
        self.configure(bg=BG)
        self.q = queue.Queue()
        self.folder = tk.StringVar()
        self.months = tk.IntVar(value=3)
        self.link_all = tk.BooleanVar(value=False)
        self.partial_build = False
        self.built = None          # (key, workbook, log, calc)
        self.current = None        # dict describing the report on screen
        self.last_export = None
        self.menu_labels = {}
        self.busy = False
        self.reset_pending = False
        self.audit_panel = None
        self.page_cache = OrderedDict()
        self.sheet_pages = {}
        self.sheet_dimensions = {}
        self.page_size = 200
        self.zoom = 90
        self.fit_columns = True
        self.inventory = None
        self.search_hits = []
        self.search_index = -1
        self.search_target = None
        self.import_results = []
        self.last_conversion_results = []
        self.conversion_note = ""
        self.last_table = None
        self.setup_styles()
        self.build_header()
        self.build_body()
        self.status("Select the company folder to begin.")
        self.after(120, self.poll)
        self.after_idle(self.maximize_window)

    def maximize_window(self):
        # macOS zoom keeps the title bar, Dock and menu bar available.
        try:
            self.state("zoomed")
        except tk.TclError:
            try:
                self.attributes("-zoomed", True)
            except tk.TclError:
                self.geometry(f"{self.winfo_screenwidth()}x{max(640, self.winfo_screenheight()-80)}+0+25")

    # ------------------------------------------------------------ layout
    def setup_styles(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                     rowheight=27, borderwidth=1, relief="solid", font=(FONT, 10, "bold"))
        st.configure("Treeview.Heading", background="#e4edf5", foreground=NAVY,
                     font=(FONT, 11, "bold"), relief="raised", padding=(10, 10))
        st.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "white")])
        st.configure("Vertical.TScrollbar", background="#c4cedb", troughcolor=BG, borderwidth=0)
        st.configure("Horizontal.TScrollbar", background="#c4cedb", troughcolor=BG, borderwidth=0)

    def build_header(self):
        head = self.header = tk.Frame(self, bg=NAVY)
        head.pack(fill="x")
        tk.Label(head, text="AUDIT  /  FINANCE", bg=NAVY, fg="white",
                 font=(FONT, 13, "bold")).pack(side="left", padx=12, pady=5)
        tk.Label(head, text="PRIVATE WORKSPACE  •  LOCAL FILES", bg=NAVY, fg="#9db4d3",
                 font=(FONT, 9)).pack(side="left", pady=5)

        bar = self.folder_bar = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        bar.pack(fill="x")
        row = tk.Frame(bar, bg=CARD)
        row.pack(fill="x", padx=8, pady=4)
        tk.Label(row, text="Company folder", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(side="left")
        tk.Entry(row, textvariable=self.folder, bg="#f7f9fc", fg=TEXT, insertbackground=TEXT,
                 relief="flat", highlightthickness=1, highlightbackground=LINE,
                 highlightcolor=ACCENT, font=(FONT, 12)).pack(side="left", fill="x", expand=True, padx=10, ipady=5)
        Btn(row, "Select folder", self.pick_folder).pack(side="left", padx=(0, 6))
        Btn(row, "Source overview", lambda: self.open_report("sources"), padx=9).pack(side="left", padx=(0, 4))
        Btn(row, "Build all data", lambda: self.open_report("build_all"), padx=9).pack(side="left", padx=(0, 4))
        Btn(row, "Open folder", lambda: self.open_if(self.folder.get()), color="#5d7391").pack(side="left", padx=(0, 6))
        Btn(row, "Refresh", self.refresh, color="#5d7391").pack(side="left")

    def build_body(self):
        body = self.body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)
        side = self.sidebar = tk.Frame(body, bg=SIDE, width=225)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        canvas = tk.Canvas(side, bg=SIDE, highlightthickness=0, width=225)
        inner = tk.Frame(canvas, bg=SIDE)
        window = canvas.create_window((0, 0), window=inner, anchor="nw", width=225)
        scroll = ttk.Scrollbar(side, orient="vertical", command=canvas.yview)
        scroll.pack(side="right", fill="y")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        canvas.pack(fill="both", expand=True)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        for kind, key, label in MENU:
            if kind == "group":
                tk.Label(inner, text=label, bg=SIDE, fg="#7f9cc4", font=(FONT, 10, "bold"),
                         anchor="w").pack(fill="x", padx=18, pady=(16, 4))
                continue
            lab = tk.Label(inner, text=("      " + label) if kind == "sub" else label, bg=SIDE_HOVER, fg="#d5e1f1", relief="solid", borderwidth=1,
                           anchor="w", padx=10 if kind == "item" else 16, pady=7, cursor="hand2",
                           font=(FONT, 11 if kind == "item" else 10))
            lab.pack(fill="x", padx=(18 if kind == "sub" else 7, 7), pady=3)
            lab.bind("<MouseWheel>", lambda event: canvas.yview_scroll(-1 if event.delta > 0 else 1, "units"))
            lab.bind("<Button-1>", lambda e, k=key: self.open_report(k))
            lab.bind("<Enter>", lambda e, l=lab: l.config(bg=SIDE_HOVER) if l is not self.active_label() else None)
            lab.bind("<Leave>", lambda e, l=lab: l.config(bg=SIDE_HOVER) if l is not self.active_label() else None)
            self.menu_labels[key] = lab
        self.active_key = None

        main = self.main_panel = tk.Frame(body, bg=BG)
        main.pack(side="left", fill="both", expand=True, padx=6, pady=4)
        top = self.report_header = tk.Frame(main, bg=BG)
        top.pack(fill="x")
        titles = tk.Frame(top, bg=BG)
        titles.pack(side="left", fill="x", expand=True)
        self.title_lbl = tk.Label(titles, text="Your audit workspace", bg=BG, fg=NAVY, font=(FONT, 14, "bold"), anchor="w")
        self.title_lbl.pack(fill="x")
        self.sub_lbl = tk.Label(titles, text="Choose a report on the left. Reports open for viewing; export when you want a file.",
                                bg=BG, fg=MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=700)
        # Descriptions are shown in the status area; keep the preview tall.
        tools = tk.Frame(top, bg=BG)
        tools.pack(side="right")
        self.btn_export = Btn(tools, "Export this report", self.export_this, color=GREEN, hover="#14664a")
        self.btn_export.pack(side="left", padx=4)
        Btn(tools, "Export everything", self.export_all).pack(side="left", padx=4)
        Btn(tools, "Open last export", lambda: self.open_if(self.last_export), color="#5d7391").pack(side="left", padx=4)

        self.sub_lbl.bind("<Configure>", lambda event: self.sub_lbl.configure(
            wraplength=max(240, event.width - 12)))
        self.company_label = tk.Label(main, text="Select a company folder", bg=BG, fg=MUTED,
                                      font=(FONT, 10), anchor="w")
        # The selected company path is already visible in the folder bar.
        self.tax_options = tk.Frame(main, bg=BG)
        self.months_label = tk.Label(self.tax_options, text="Months of newest year (older year: all 12)", bg=BG, fg=TEXT)
        self.months_label.pack(side="left")
        tk.Spinbox(self.tax_options, from_=1, to=12, width=4, textvariable=self.months).pack(side="left", padx=8)
        Btn(self.tax_options, "Apply", self.apply_months).pack(side="left")

        self.toi_pdf_options = tk.Frame(main, bg=BG)
        tk.Label(self.toi_pdf_options, text="VBA Module 10 · steps", bg=BG, fg=TEXT).pack(side="left")
        self.toi_steps = tk.StringVar(value="1-17")
        tk.Entry(self.toi_pdf_options, textvariable=self.toi_steps, width=8).pack(side="left", padx=6)
        self.toi_renderer = tk.StringVar(value="Microsoft Excel" if sys.platform in ("win32","darwin") else "LibreOffice")
        ttk.Combobox(self.toi_pdf_options, textvariable=self.toi_renderer,
                     values=("Microsoft Excel","LibreOffice"),state="readonly",width=16).pack(side="left",padx=4)
        Btn(self.toi_pdf_options, "Workbook fonts", self.font_help, padx=8).pack(side="left",padx=3)
        Btn(self.toi_pdf_options, "Export TOI PDFs", self.export_toi_pdf).pack(side="left", padx=3)
        self.toi_output=tk.StringVar(value="TOI_PDFs")
        ttk.Label(self.toi_pdf_options,text="Save in").pack(side="left",padx=2)
        ttk.Combobox(self.toi_pdf_options,textvariable=self.toi_output,values=("TOI_PDFs","TOI"),state="readonly",width=10).pack(side="left",padx=2)
        Btn(self.toi_pdf_options, "Open PDF folder", lambda:self.open_if(self.pdf_output_folder())).pack(side="left", padx=3)
        self.source_controls = tk.Frame(main, bg=BG)
        Btn(self.source_controls, "Organize Download", self.organize_download, padx=8).pack(side="left", padx=3)
        Btn(self.source_controls, "Format summary", lambda: self.open_report("sources")).pack(side="left", padx=3)
        Btn(self.source_controls, "Validate files", lambda: self.open_report("check_tax")).pack(side="left", padx=3)
        Btn(self.source_controls, "Convert selected row", lambda: self.convert_target_pdfs(False), padx=8).pack(side="left", padx=3)
        Btn(self.source_controls, "Convert all target PDFs", lambda: self.convert_target_pdfs(True), padx=8).pack(side="left", padx=3)
        Btn(self.source_controls, "Open Excel", self.open_converted_excel, padx=8).pack(side="left", padx=3)
        Btn(self.source_controls, "Import log", lambda: self.open_report("import_results"), padx=6).pack(side="left", padx=2)
        Btn(self.source_controls, "Results", lambda: self.open_report("conversion_results"), padx=8).pack(side="left", padx=3)
        actions = self.source_controls.winfo_children()
        for widget in actions:
            widget.pack_forget()
        for index, widget in enumerate(actions):
            widget.grid(row=index//4, column=index%4, sticky="ew", padx=2, pady=2)
        for column in range(4):
            self.source_controls.columnconfigure(column, weight=1)
        self.search_controls = tk.Frame(main, bg=BG)
        tk.Label(self.search_controls, text="Search all tax sheets", bg=BG, fg=TEXT).pack(side="left")
        self.search_query = tk.StringVar()
        search_entry = tk.Entry(self.search_controls, textvariable=self.search_query, width=30)
        search_entry.pack(side="left", padx=6)
        search_entry.bind("<Return>", lambda event: self.find_text())
        Btn(self.search_controls, "Find", self.find_text).pack(side="left", padx=2)
        Btn(self.search_controls, "Next match", self.next_match).pack(side="left", padx=2)
        self.search_label = tk.Label(self.search_controls, text="", bg=BG, fg=MUTED)
        self.search_label.pack(side="left", padx=6)

        self.picker = tk.Frame(main, bg=BG)
        self.picker.pack(fill="x", pady=(3, 0))
        card = self.preview_card = tk.Frame(main, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True, pady=(3, 3))
        preview_tools = tk.Frame(card, bg="#e4edf5")
        preview_tools.pack(fill="x")
        Btn(preview_tools, "−", lambda: self.change_zoom(-10)).pack(side="left", padx=2)
        self.zoom_label = tk.Label(preview_tools, text="90%", bg="#e4edf5", fg=NAVY)
        self.zoom_label.pack(side="left")
        Btn(preview_tools, "+", lambda: self.change_zoom(10)).pack(side="left", padx=2)
        Btn(preview_tools, "Fit columns", self.toggle_fit).pack(side="left", padx=4)
        Btn(preview_tools, "Actual widths", self.actual_widths).pack(side="left", padx=2)
        self.preview_expanded = False
        self.expand_button = Btn(preview_tools, "Expand preview", self.toggle_preview)
        self.expand_button.pack(side="left", padx=4)
        self.dashboard_options = tk.Checkbutton(preview_tools, text="Fill blank ADT links",
            variable=self.link_all, bg="#e4edf5", command=self.invalidate)
        Btn(preview_tools, "Previous", lambda: self.change_page(-1)).pack(side="right", padx=2)
        Btn(preview_tools, "Next", lambda: self.change_page(1)).pack(side="right", padx=2)
        self.page_label = tk.Label(preview_tools, text="", bg="#e4edf5", fg=NAVY)
        self.page_label.pack(side="right", padx=8)
        self.tree = ttk.Treeview(card, show="headings", selectmode="browse")
        for tag, options in {
            "even": {"background": "#ffffff"},
            "odd": {"background": "#f2f6fa"},
            "source_link": {"foreground": "#126fa3"},
            "section": {"background": "#dce9f3", "foreground": NAVY},
            "total": {"background": "#d9eee7", "foreground": "#12533f"},
            "warning": {"background": "#fff1dc", "foreground": "#854c0a"},
            "error": {"background": "#fbe5e7", "foreground": "#9a293c"},
        }.items():
            self.tree.tag_configure(tag, **options)
        ys = ttk.Scrollbar(card, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(card, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.pack(side="right", fill="y")
        xs.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Configure>", self.resize_preview)
        self.tree.bind("<Double-1>", self.pdf_row_action)
        self.tree.bind("<ButtonRelease-1>", self.open_mtr_source)
        self.tree.bind("<Motion>", self.mtr_link_cursor)
        self.status_lbl = tk.Label(main, text="", bg=BG, fg=MUTED, anchor="w", font=(FONT, 10))
        self.status_lbl.pack(fill="x")

    # ------------------------------------------------------------ helpers
    def active_label(self):
        return self.menu_labels.get(self.active_key)

    def mark_active(self, key):
        for k, lab in self.menu_labels.items():
            lab.config(bg=SIDE_HOVER, fg="#d5e1f1")
        self.active_key = key
        if key in self.menu_labels:
            self.menu_labels[key].config(bg=ACCENT, fg="white")

    def status(self, msg):
        self.status_lbl.config(text=msg)

    def open_if(self, p):
        if p and Path(p).exists():
            open_path(p)
        else:
            messagebox.showinfo("Not found", "Nothing to open yet.")

    def pick_folder(self):
        d = filedialog.askdirectory(title="Select the company folder")
        if d and not self.busy:
            self.folder.set(d)
            self.invalidate()
            self.run_bg(self.job_prepare_sources)

    def organize_download(self):
        if self.busy or not self.need_folder():
            return
        self.invalidate()
        self.run_bg(self.job_prepare_sources, True)

    def job_prepare_sources(self, force=False):
        report = organize_downloads(self.job_settings[0], allow_existing=force)
        return ("prepared_sources",report,force)

    def need_folder(self):
        if not self.folder.get().strip() or not Path(self.folder.get()).is_dir():
            messagebox.showwarning("Select folder", "Click Select folder and choose the company folder first.")
            return False
        return True

    def invalidate(self):
        if self.busy:
            return
        self.built = None
        self.page_cache.clear()
        self.sheet_pages.clear()
        self.sheet_dimensions.clear()
        self.inventory = None
        self.search_hits = []
        self.search_target = None
        self.conversion_note = ""
        self.import_results = []
        self.last_conversion_results = []

    def refresh(self):
        if self.busy:
            self.reset_pending = True
            self.status("Reset requested. The current operation will finish safely, then the session will clear.")
            return
        self.reset_session()

    def reset_session(self):
        if self.preview_expanded:
            self.toggle_preview()
        self.invalidate()
        self.folder.set("");self.months.set(3);self.link_all.set(False)
        self.partial_build=False;self.current=None;self.last_export=None;self.reset_pending=False
        self.search_index=-1;self.search_query.set("");self.search_label.config(text="")
        self.page_size=200;self.zoom=90;self.fit_columns=True
        self.toi_steps.set("1-17")
        self.toi_renderer.set("Microsoft Excel" if sys.platform in ("win32","darwin") else "LibreOffice")
        if hasattr(self,"toi_output"):self.toi_output.set("TOI_PDFs")
        for widget in (self.source_controls,self.search_controls,self.tax_options,self.toi_pdf_options):widget.pack_forget()
        self.dashboard_options.pack_forget()
        if self.audit_panel is not None:
            self.audit_panel.reset();self.audit_panel.pack_forget()
        if not self.preview_card.winfo_manager():self.preview_card.pack(fill="both",expand=True,before=self.status_lbl,pady=3)
        self.set_picker([],"");self.tree.delete(*self.tree.get_children());self.tree["columns"]=()
        self.tree.xview_moveto(0);self.tree.yview_moveto(0)
        self.last_table=None;self.page_label.config(text="");self.company_label.config(text="Select a company folder")
        self.title_lbl.config(text="Your audit workspace");self.sub_lbl.config(text="")
        self.mark_active(None);self.btn_export.set_enabled(False)
        self.setup_styles();self.zoom_label.config(text="90%");self.config(cursor="")
        self.status("Select the company folder to begin.")

    def notify_build_done(self):
        try:
            if sys.platform=="win32":
                import winsound
                winsound.MessageBeep(winsound.MB_OK)
            elif sys.platform=="darwin" and Path('/System/Library/Sounds/Glass.aiff').is_file():
                _sound_subprocess.Popen(['afplay','/System/Library/Sounds/Glass.aiff'],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
            else:self.bell()
        except (OSError,ImportError):self.bell()

    # ------------------------------------------------------------ table drawing
    def draw(self, columns, rows, align_right=True, tags=None):
        self.tree.delete(*self.tree.get_children())
        ids = [f"c{i}" for i in range(len(columns))]
        self.tree["columns"] = ids
        self.last_table = (columns, rows, align_right, tags)
        scale = self.zoom / 100
        available = max(100, self.tree.winfo_width() - 24)
        weights = []
        for i, name in enumerate(columns):
            longest = max([len(str(name))] + [len(str(r[i])) for r in rows[:100] if i < len(r)])
            weights.append(5 if name == "#" else max(9, min(longest + 2, 45)))
        weight_total = sum(weights) or 1
        for i, (cid, name) in enumerate(zip(ids, columns)):
            cells = [str(r[i]) for r in rows[:100] if i < len(r) and r[i] not in (None, "")]
            numeric = align_right and cells and sum(bool(NUMBER.match(c)) for c in cells) / len(cells) > 0.6
            natural = max(55, min(max([len(name)] + [len(c) for c in cells]) * 7 + 20, 400))
            width = max(12, int(available * weights[i] / weight_total)) if self.fit_columns else int(natural * scale)
            self.tree.heading(cid, text=name)
            self.tree.column(cid, width=width, minwidth=10,
                             anchor="e" if numeric else "w", stretch=False)
        for index, r in enumerate(rows):
            values = ["" if v is None else v for v in r]
            tag = tags[index] if tags else report_row_tag(values, index)
            self.tree.insert("", "end", values=values, tags=(tag,))
        self.company_label.config(text=Path(self.folder.get()).name if self.folder.get() else "Select a company folder")

    def apply_months(self):
        if self.busy:
            return
        try:
            months = self.months.get()
            if not 1 <= months <= 12:
                raise ValueError()
        except (ValueError, tk.TclError):
            messagebox.showwarning("Months", "Enter a number from 1 to 12.")
            return
        self.open_report("tps")

    def change_zoom(self, delta):
        self.zoom = max(50, min(150, self.zoom + delta))
        self.zoom_label.config(text=f"{self.zoom}%")
        style = ttk.Style(self)
        style.configure("Treeview", font=(FONT, max(7, round(11*self.zoom/100)), "bold"),
                        rowheight=max(18, round(30*self.zoom/100)))
        style.configure("Treeview.Heading", font=(FONT, max(7, round(11*self.zoom/100)), "bold"))
        self.redraw_table()

    def redraw_table(self):
        if self.last_table is not None:
            self.draw(*self.last_table[:3], tags=self.last_table[3])

    def toggle_fit(self):
        self.fit_columns = True
        self.redraw_table()

    def actual_widths(self):
        self.fit_columns = False
        self.redraw_table()

    def toggle_preview(self):
        self.preview_expanded = not self.preview_expanded
        if self.preview_expanded:
            for widget in (self.header, self.folder_bar, self.sidebar, self.report_header):
                widget.pack_forget()
            self.expand_button.config(text="Restore layout")
        else:
            self.folder_bar.pack(fill="x", before=self.body)
            self.header.pack(fill="x", before=self.folder_bar)
            self.sidebar.pack(side="left", fill="y", before=self.main_panel)
            self.report_header.pack(fill="x", before=self.picker)
            self.expand_button.config(text="Expand preview")

    def resize_preview(self, event):
        if self.fit_columns:
            # Resize columns only; don't rebuild thousands of table items.
            ids = self.tree["columns"]
            widths = [self.tree.column(cid, "width") for cid in ids]
            total = sum(widths) or 1
            available = max(100, event.width - 24)
            for cid, width in zip(ids, widths):
                self.tree.column(cid, width=max(10, int(available * width / total)))

    def change_page(self, direction):
        if self.busy or not self.current or not self.current.get("sheet"):
            return
        name = self.current["sheet"]
        total = self.sheet_dimensions[name][1]
        self.sheet_pages[name] = max(0, min((total-1)//self.page_size,
                                           self.sheet_pages.get(name, 0) + direction))
        self.show_sheet(name)

    def job_page(self, name, page):
        wb, cal = self.built[1], self.built[3]
        ws = wb[name]
        # Compute dimensions once; OpenPyXL max_column scans all stored cells.
        if name not in self.sheet_dimensions:
            self.sheet_dimensions[name] = (ws.max_column, ws.max_row)
        ncols, total = self.sheet_dimensions[name]
        start = page * self.page_size + 1
        rows, tags = [], []
        for r in range(start, min(total + 1, start + self.page_size)):
            cells = [ws.cell(r, c) for c in range(1, ncols + 1)]
            values = [calc.display(cal.value(name, r, c), cell.number_format)
                      for c, cell in enumerate(cells, 1)]
            tag = report_row_tag(values, r-1)
            if tag in ("even", "odd") and any(cell.value is not None and cell.font.bold for cell in cells):
                tag = "section"
            if name.startswith("MTR_") and len(cells)>1 and cells[1].hyperlink:
                tag="source_link"
            rows.append([r] + values)
            tags.append(tag)
        return ("page", name, page, ncols, total, rows, tags)

    def present_page(self, msg):
        _, name, page, ncols, total, rows, tags = msg
        self.draw(["#"] + [calc_letter(i+1) for i in range(ncols)], rows, tags=tags)
        if name == "Tax Paid Summary" and self.built:
            label = self.built[1][name]["D9"].value
            self.months_label.config(text=f"Newest year: {label} · older year: 12 months")
        self.page_label.config(text=f"Page {page+1} / {max(1, (total+self.page_size-1)//self.page_size)}")
        if self.search_target and self.search_target[0] == name:
            target = self.search_target[1]
            for item in self.tree.get_children():
                if int(self.tree.item(item, "values")[0]) == target:
                    self.tree.selection_set(item)
                    self.tree.focus(item)
                    self.tree.see(item)
                    break
        hint=" Click a file name in column B to open its source workbook." if name.startswith("MTR_") else ""
        self.status(f"{name}: rows {page*self.page_size+1:,}–{min(total,(page+1)*self.page_size):,} of {total:,}. All rows available using Next / Previous."+hint)

    def set_picker(self, names, active):
        for w in self.picker.winfo_children():
            w.destroy()
        if len(names) < 2:
            return
        for n in names:
            b = Btn(self.picker, "TOI compare" if n == "TOI" else n, lambda n=n: self.show_sheet(n), color=ACCENT if n == active else "#9fb2cc",
                    hover=NAVY, padx=12, pady=4)
            b.pack(side="left", padx=(0, 6))

    def show_sheet(self, name):
        if self.busy:
            return
        self.toi_pdf_options.pack_forget()
        self.btn_export.set_enabled((self.current or {}).get("key") in engine.SECTIONS)
        if name == "Export TOI PDF":
            self.current.pop("sheet",None)
            self.set_picker(list(self.current["tables"]), name)
            self.toi_pdf_options.pack(fill="x", after=self.picker, pady=4)
            self.btn_export.set_enabled(False)
            self.page_label.config(text="")
            self.draw(["TOI PDF export"], [["Export selected steps from every TOI .xlsx file to TOI_PDFs."],
                      ["Use 1-17 for all steps, or a range such as 4-4."],
                      ["Microsoft Excel is recommended for Khmer, matching the original VBA renderer."],
                      ["Install matching workbook fonts (for example Khmer OS Siemreap); LibreOffice is an alternative."],
                      ["On Mac, grant Excel access to TOI_PDFs when requested; stable work paths reduce repeat requests."],
                      ["Source workbooks and existing PDFs are preserved."]])
            return
        self.current["sheet"] = name
        self.set_picker(list(self.current["tables"]), name)
        page = self.sheet_pages.get(name, 0)
        cache_key = (name, page)
        if cache_key in self.page_cache:
            self.page_cache.move_to_end(cache_key)
            self.present_page(self.page_cache[cache_key])
        else:
            self.run_bg(self.job_page, name, page)

    def pdf_output_folder(self):
        root=Path(self.folder.get())
        return (engine.find_dir(root,"TOI") or root/"TOI") if self.toi_output.get()=="TOI" else root/"TOI_PDFs"

    def export_toi_pdf(self):
        if self.busy or not self.need_folder():
            return
        value = self.toi_steps.get()
        try:
            parse_steps(value)
        except ValueError as exc:
            self.status(str(exc));return
        backend = "excel" if self.toi_renderer.get()=="Microsoft Excel" else "libreoffice"
        self.run_bg(self.job_toi_pdf,value,backend,self.toi_output.get())

    def font_help(self):
        if self.busy or not self.need_folder():
            return
        self.run_bg(lambda: ("toi_fonts",toi_workbook_fonts(self.job_settings[0])))

    def job_toi_pdf(self,value,backend='auto',output='TOI_PDFs'):
        rows = export_toi_pdfs(self.job_settings[0],value,
                    progress=lambda done,total: self.q.put(("pdf_progress",done,total)),renderer=backend,output_folder=output)
        return ("toi_pdfs_done",rows,output)

    def job_tois(self):
        key = self.job_settings
        if self.built and self.built[0] == key and {"TOI","FULL TOI"} <= set(self.built[1].sheetnames):
            wb = self.built[1]
        else:
            root = Path(key[0]); directory = engine.find_dir(root,"TOI")
            source = engine.first_xlsx(directory) if directory else None
            if source is None:
                raise ValueError("Put the TOI Excel workbook(s) in the TOI folder first.")
            wb,log = engine.Workbook(),engine.Log();wb.remove(wb.active)
            input_book = engine.load_workbook(source,data_only=True)
            try:
                engine.build_toi(wb,input_book,log)
                engine.build_full_toi(wb,input_book,log)
            finally:
                input_book.close()
            self.built = (key,wb,log,calc.Calc(wb))
            self.partial_build = True
            self.page_cache.clear();self.sheet_pages.clear();self.sheet_dimensions.clear()
        return ("sheets","tois",dict.fromkeys(["TOI","FULL TOI","Export TOI PDF"]))

    def mtr_link_at(self,event):
        if self.busy or not self.built or not self.current:
            return None
        sheet=self.current.get("sheet")
        if not sheet or not sheet.startswith("MTR_") or self.tree.identify_column(event.x)!="#3":
            return None
        item=self.tree.identify_row(event.y)
        values=self.tree.item(item,"values") if item else ()
        if not values:
            return None
        try:
            return mtr_source_path(self.built[1],sheet,int(values[0]))
        except (ValueError,TypeError):
            return None

    def mtr_link_cursor(self,event):
        self.tree.configure(cursor="hand2" if self.mtr_link_at(event) else "")

    def open_mtr_source(self,event):
        path=self.mtr_link_at(event)
        if path is None:
            return
        if not path.is_file():
            self.status("Source workbook no longer exists: "+str(path))
            return
        try:
            open_path(path)
        except OSError as exc:
            self.status("Could not open source workbook: "+str(exc))

    def pdf_row_action(self, event):
        if self.current and self.current.get("key") in ("check_tax", "conversion_results"):
            row = self.tree.identify_row(event.y)
            if row:
                self.tree.selection_set(row)
                values = self.tree.item(row,"values")
                if str(values[4]).startswith(("CONVERTED:", "ALREADY CONVERTED:")):
                    self.open_converted_excel()
                else:
                    self.convert_target_pdfs(False)

    def open_converted_excel(self):
        if self.busy or not self.need_folder():
            return
        selection = self.tree.selection()
        if not selection:
            self.status("Select a converted PDF row in Results or Validate files first.")
            return
        values = self.tree.item(selection[0], "values")
        sources = view_pdfs(self.folder.get(), values, (self.current or {}).get("key", "sources"))
        records = conversion_records(self.folder.get())
        outputs = [converted_output(self.folder.get(),p,records) for p in sources]
        outputs = [p for p in outputs if p]
        if outputs:
            open_path(outputs[0] if len(outputs)==1 else outputs[0].parent)
        else:
            self.status("No registered Excel extraction for this row. Use Convert selected row first.")

    def convert_target_pdfs(self, all_targets=False):
        if self.busy or not self.need_folder():
            return
        values = None
        if not all_targets:
            selection = self.tree.selection()
            if not selection:
                self.status("Select a source folder or PDF row first, or use Convert all target PDFs.")
                return
            values = self.tree.item(selection[0], "values")
        key = (self.current or {}).get("key", "sources")
        self.run_bg(self.job_batch_convert, values, key)

    def job_batch_convert(self, values, key):
        folder = self.job_settings[0]
        sources = view_pdfs(folder,values,key)
        if not sources:
            return ("batch_done", [], "No PDFs found in the selected target scope.")
        results = batch_convert_pdfs(folder,sources,progress=lambda done,total: self.q.put(("batch_progress",done,total)))
        failures = sum(r[4].startswith("ERROR") for r in results)
        note = f"{len(results)-failures} of {len(results)} PDFs have Excel extractions; {failures} failed. Review/mapping required. See Results."
        return ("batch_done", results, note)

    def pick_pdf(self):
        if self.busy:
            return
        source = filedialog.askopenfilename(title="Select a text PDF to extract", filetypes=[("PDF", "*.pdf")])
        if not source:
            return
        output = filedialog.asksaveasfilename(title="Save extracted workbook for review", defaultextension=".xlsx",
                    initialdir=str(Path(source).parent), initialfile=Path(source).stem+"_review.xlsx",
                    filetypes=[("Excel", "*.xlsx")])
        if output:
            self.run_bg(lambda: ("converted", convert_pdf(source, output)))

    def find_text(self):
        if self.busy or not self.current or self.current.get("key") != "consolidate":
            return
        query = self.search_query.get().strip()
        self.tree.selection_remove(self.tree.selection())
        self.search_hits = []
        self.search_target = None
        self.search_index = -1
        if not query:
            self.search_label.config(text="Enter a value to find")
            return
        self.run_bg(self.job_search, query, list(self.current["tables"]))

    def job_search(self, query, names):
        wb, cal = self.built[1], self.built[3]
        needle = query.casefold()
        hits = []
        for name in names:
            ws = wb[name]
            # All rows, not just the loaded preview page.
            ncols, nrows = ws.max_column, ws.max_row
            for row in ws.iter_rows(min_row=1, max_row=nrows, max_col=ncols):
                if any(needle in str(cell.value if cell.data_type != 'f' else cal.value(name, cell.row, cell.column)).casefold()
                       or needle in calc.display(cal.value(name, cell.row, cell.column), cell.number_format).casefold()
                       for cell in row if cell.value is not None):
                    hits.append((name, row[0].row))
        return ("search", hits)

    def next_match(self):
        if self.busy or not self.search_hits:
            return
        self.search_index = (self.search_index + 1) % len(self.search_hits)
        name, row = self.search_hits[self.search_index]
        self.search_target = (name, row)
        self.sheet_pages[name] = (row-1)//self.page_size
        self.search_label.config(text=f"{self.search_index+1} / {len(self.search_hits)} matching rows · {name} row {row}")
        self.show_sheet(name)

    # ------------------------------------------------------------ opening a report
    def open_report(self, key):
        if self.busy:
            self.status("Please wait, still working…")
            return
        if self.audit_panel is not None:self.audit_panel.pack_forget()
        if not self.preview_card.winfo_manager():self.preview_card.pack(fill="both",expand=True,before=self.status_lbl,pady=3)
        label = ACTION_TITLES.get(key) or next(m[2] for m in MENU if m[1] == key)
        self.mark_active(key)
        self.title_lbl.config(text=label.strip())
        self.sub_lbl.config(text=SUBTITLE.get(key, ""))
        self.current = {"key": key}
        self.search_controls.pack_forget()
        self.source_controls.pack_forget()
        self.search_target = None
        self.search_hits = []
        self.search_label.config(text="")
        if key == "consolidate":
            self.search_controls.pack(fill="x", before=self.picker, pady=3)
        if key in ("sources", "source_files", "check_tax", "conversion_results", "import_results", "check_fin") or key.startswith("fin_"):
            self.source_controls.pack(fill="x", before=self.picker, pady=3)
        self.page_label.config(text="")
        self.tax_options.pack_forget()
        self.toi_pdf_options.pack_forget()
        self.dashboard_options.pack_forget()
        if key == "dashboard":
            self.dashboard_options.pack(side="left", padx=6)
        if key == "tps":
            self.tax_options.pack(fill="x", before=self.picker, pady=4)
        self.btn_export.set_enabled(key in engine.SECTIONS)
        if key == "audit_tools":
            if not self.need_folder():return
            self.set_picker([],"");self.preview_card.pack_forget()
            if self.audit_panel is None:self.audit_panel=AuditToolsPanel(self.main_panel,self)
            self.audit_panel.pack(fill="both",expand=True,before=self.status_lbl,pady=4)
            self.audit_panel.load()
            return
        if key == "import_results":
            self.set_picker([], "")
            self.draw(["Download source", "Destination", "Result"],self.import_results)
            self.status("Organize Download starts an import; this log shows moved/skipped files. Existing files are never replaced.")
            return
        if key == "conversion_results":
            self.set_picker([], "")
            self.draw(["Folder", "Scope", "PDF", "Rows", "Result", "Excel output"], self.last_conversion_results)
            self.status("Double-click converted rows to open Excel, or failed PDF rows to retry conversion.")
            return
        if key in ("sources", "source_files"):
            if self.need_folder():
                self.run_bg(self.job_inventory if key == "sources" else self.job_source_files)
            return
        if key in ("check_fin", "tie_out") or key.startswith("fin_"):
            if not self.need_folder():
                return
            self.run_bg(self.job_finance, key)
            return
        if key == "check_tax":
            if not self.need_folder():
                return
            self.run_bg(self.job_check_tax)
            return
        if not self.need_folder():
            return
        if key == "tois":
            self.run_bg(self.job_tois)
        else:
            self.run_bg(self.job_report, key)

    def run_bg(self, fn, *args):
        if self.busy:
            return
        # Read Tk variables only on the main thread, never in worker jobs.
        try:
            self.job_settings = (self.folder.get(), self.months.get(), self.link_all.get())
        except tk.TclError:
            messagebox.showwarning("Options", "Enter a valid month count.")
            return
        self.busy = True
        self.status("Working…")
        self.config(cursor="watch")

        def go():
            try:
                self.q.put(fn(*args))
            except Exception as exc:
                self.q.put(("error", str(exc)))
        threading.Thread(target=go, daemon=True).start()

    # ------------------------------------------------------------ background jobs (no Tk calls here)
    def ensure_built(self, summary_only=False):
        key = self.job_settings
        needs_full = not summary_only and getattr(self, "partial_build", False) is True
        needs_summary = summary_only and self.built is not None and "Tax Paid Summary" not in self.built[1]
        if self.built is None or self.built[0] != key or needs_full or needs_summary:
            if summary_only and self.built is None:
                info_dir = engine.find_dir(Path(key[0]), "INFO")
                info_file = engine.first_xlsx(info_dir) if info_dir else None
                if info_file is None:
                    raise ValueError("Tax Paid Summary needs an Excel workbook in the Info folder.")
                info = engine.load_workbook(info_file, data_only=True)
                try:
                    wb, log = engine.Workbook(), engine.Log()
                    wb.remove(wb.active)
                    engine.company_info(wb, info, log)
                    engine.tax_paid_summary(wb, info, key[1], log)
                finally:
                    info.close()
                self.partial_build = True
            elif self.built and not needs_full and self.built[0][0] == key[0] and self.built[0][2] == key[2]:
                wb, log = self.built[1:3]
                info_dir = engine.find_dir(Path(key[0]), "INFO")
                info_file = engine.first_xlsx(info_dir) if info_dir else None
                if info_file:
                    info = engine.load_workbook(info_file, data_only=True)
                    try:
                        engine.tax_paid_summary(wb, info, key[1], log)
                    finally:
                        info.close()
            else:
                wb, log = engine.build_workbook(key[0], key[1], key[2])
                self.partial_build = False
            self.page_cache.clear()
            self.sheet_pages.clear()
            self.sheet_dimensions.clear()
            self.built = (key, wb, log, calc.Calc(wb))
        return self.built

    def job_inventory(self):
        folder = self.job_settings[0]
        if self.inventory is None or self.inventory[0] != folder:
            self.inventory = (folder, source_inventory(folder))
        return ("table", ["Source", "Year / scope", "Files", "Status", "Formats"], self.inventory[1],
                self.conversion_note or "Target-folder inventory. Use Validate files to check contents.", False)

    def job_source_files(self):
        root = Path(self.job_settings[0])
        rows = []
        for path in target_files(root):
            if (path.is_file() and path.suffix.lower() in engine.FIN_EXT and not any(
                    part.startswith((".", "~$")) for part in path.relative_to(root).parts)):
                rel = path.relative_to(root)
                kind = {".xlsx": "Excel", ".xlsm": "Excel macro", ".xls": "Excel legacy",
                        ".csv": "CSV", ".pdf": "PDF"}[path.suffix.lower()]
                rows.append([rel.parts[0], rel.parts[1] if len(rel.parts) > 2 else "—", path.name, kind, str(rel)])
        return ("table", ["Folder", "Year / subfolder", "File name", "Format", "Relative path"], rows,
                f"{len(rows):,} files found. Inventory only; PDF/CSV presence does not imply tax import support.", False)

    def job_check_tax(self):
        rows = validate_sources(self.job_settings[0])
        bad = sum(1 for r in rows if r[4] != "READY")
        displayed = [list(r)+[("Open Excel" if str(r[4]).startswith("CONVERTED:") else "Convert PDF")
                              if str(r[2]).lower().endswith(".pdf") else ""] for r in rows]
        return ("table", ["Source", "Scope", "Source file", "Rows", "Result", "Action (double-click)"], displayed,
                f"{len(rows) - bad} READY tax entries; all other rows require the action shown.", False)

    def job_finance(self, key):
        folder = self.job_settings[0]
        summary, files = engine.check_finance(folder)
        statuses = {r[0]: r[3] for r in source_inventory(folder)}
        summary = [(*r[:4], statuses.get(r[0],r[4])) for r in summary]
        records = conversion_records(folder)
        updated = []
        for row in files:
            label, name, sheets, count, note = row
            if str(name).lower().endswith(".pdf"):
                finance_key = next(k for k,v in engine.FINANCE_FOLDERS.items() if v[0]==label)
                directory = engine.finance_folder(folder,finance_key)
                output = converted_output(folder,directory/name,records)
                if output:
                    note = "CONVERTED: Excel available; review required"
            updated.append((label,name,sheets,count,note))
        files = updated
        if key in ("check_fin", "tie_out"):
            if key == "tie_out":
                return ("table", ["Tax vs ledger"], [["Needs sample GL, bank, TB and A/R files before it can be built."]],
                        "Not built yet.", False)
            return ("table", ["Source", "Folder", "Files", "Formats", "Readiness"], [list(r) for r in summary],
                    "Place documents inside the sub-folders. PDFs are flagged, never assumed correct.", False)
        fkey = key.split("_", 1)[1]
        label = engine.FINANCE_FOLDERS[fkey][0]
        rows = [list(r[1:]) for r in files if r[0] == label]
        if not rows:
            rows = [["(no files found)", "", "", "Add files to the " + label + " sub-folder"]]
        return ("table", ["File", "Sheets", "Rows", "Result"], rows,
                f"{label}: reconciliation is added after sample files are reviewed.", False)

    def job_report(self, key):
        built = self.ensure_built(summary_only=key == "tps")
        _k, wb, log, cal = built
        if key == "log":
            rows = [list(r) for r in log]
            return ("table", ["Time", "Where", "Message"], rows, f"{len(rows)} messages.", False)
        if key == "build_all":
            rows = [[ws.title, ws.max_row, ws.max_column] for ws in wb.worksheets]
            return ("table", ["Report sheet", "Rows", "Columns"], rows,
                    f"Built {len(rows)} sheets in memory. Open each report on the left, or press Export everything.",
                    True)
        names = [n for n in wb.sheetnames if n in engine.SECTIONS[key][1]]
        if key == "dashboard":
            names = [n for n in names if n != "TOI"]  # keep hidden workbook dependency for dashboard formulas/export
        if key == "mtr":
            names = [n for n in ("MTR_PR", "MTR_SR", "MTR_TOS", "MTR_WHT", "Summary MTR") if n in names]
        tables = dict.fromkeys(names)
        return ("sheets", key, tables)

    # ------------------------------------------------------------ queue results -> screen
    def poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                kind = msg[0]
                if kind in ("batch_progress", "pdf_progress"):
                    if self.reset_pending:continue
                    action = "Exporting TOI PDFs" if kind == "pdf_progress" else "Converting target PDFs"
                    self.status(f"{action}: {msg[1]} / {msg[2]} completed…")
                    continue
                self.busy = False
                self.config(cursor="")
                if self.reset_pending:
                    self.reset_pending=False;self.reset_session();continue
                if kind == "audit_tools_loaded":
                    if self.audit_panel is not None:self.audit_panel.populate(msg[1])
                    self.status("Tax Audit Tools ready.")
                elif kind == "audit_document_saved":
                    if self.audit_panel is not None:self.audit_panel.saved(msg[1],msg[2])
                    self.status("Saved: "+msg[2])
                elif kind == "prepared_sources":
                    self.import_results = msg[1]
                    moved = sum(r[2]=="MOVED" for r in msg[1])
                    file_rows = [r for r in msg[1] if not r[2].startswith("INFO:")]
                    if file_rows:
                        self.conversion_note = f"Download setup: {moved} moved, {len(file_rows)-moved} left for review. See Import log."
                    elif msg[1]:
                        self.conversion_note = msg[1][0][2]
                    self.inventory = None
                    self.open_report("import_results" if len(msg)>2 and msg[2] else "sources")
                elif kind == "toi_fonts":
                    self.draw(["TOI workbook", "Font families referenced", "Note"],msg[1])
                    self.status("On Mac, install matching fonts in Font Book and restart Excel. This lists workbook fonts; installed fonts are not verified.")
                elif kind == "toi_pdfs_done":
                    self.draw(["TOI workbook", "Steps", "Result", "PDF file"],msg[1])
                    successful = sum(r[2]=="EXPORTED" for r in msg[1])
                    self.status(f"TOI PDF export: {successful}/{len(msg[1])} workbooks exported. See {msg[2] if len(msg)>2 else 'TOI_PDFs'}.")
                elif kind == "batch_done":
                    self.last_conversion_results = msg[1]
                    self.conversion_note = msg[2]
                    self.inventory = None
                    self.open_report("sources")
                    self.status(msg[2])
                elif kind == "converted":
                    result = msg[1]
                    self.inventory = None
                    self.status("Extracted workbook saved for manual review: " + result['output'])
                    messagebox.showinfo("PDF extraction complete — review required",
                        f"{result['tables']} tables and {result['text_pages']} text pages saved.\n\n"
                        "Compare balances and every page with the PDF, correct data and map columns. "
                        "This file is not automatically approved for tax import or reconciliation.\n\n" + result['output'])
                elif kind == "search":
                    self.search_hits = msg[1]
                    self.search_index = -1
                    if self.search_hits:
                        self.next_match()
                    else:
                        self.search_label.config(text="No matches in any tax sheet")
                elif kind == "page":
                    self.page_cache[(msg[1], msg[2])] = msg
                    while len(self.page_cache) > 32:
                        self.page_cache.popitem(last=False)
                    self.present_page(msg)
                elif kind == "table":
                    _k, cols, rows, note, _exp = msg
                    self.set_picker([], "")
                    self.draw(cols, rows)
                    self.status(note)
                    if _exp:self.notify_build_done()
                elif kind == "sheets":
                    _k, key, tables = msg
                    self.current = {"key": key, "tables": tables}
                    if tables:
                        self.show_sheet(next(iter(tables)))
                    else:
                        self.draw(["Result"], [["This report was not produced. See Processing log."]])
                elif kind == "error":
                    self.status("Stopped: " + msg[1])
                    messagebox.showerror("Could not finish", msg[1])
                elif kind == "exported":
                    self.last_export = msg[1]
                    self.status("Saved: " + msg[1])
                    if messagebox.askyesno("Export finished", f"Saved:\n{msg[1]}\n\nOpen it now?"):
                        open_path(msg[1])
        except queue.Empty:
            pass
        self.after(120, self.poll)

    # ------------------------------------------------------------ export
    def export(self, sections, default_name):
        if self.busy:
            self.status("Please wait for the current operation to finish.")
            return
        if not self.need_folder():
            return
        path = filedialog.asksaveasfilename(
            title="Export to Excel", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=str(Path.home() / "Documents"), initialfile=default_name)
        if not path:
            return
        folder, months, link = self.folder.get(), self.months.get(), self.link_all.get()

        def job():
            engine.build(folder, path, months, link, sections)
            return ("exported", path)
        self.run_bg(job)

    def export_this(self):
        key = (self.current or {}).get("key")
        if key not in engine.SECTIONS:
            messagebox.showinfo("Export", "Open a report first. This view has nothing to export.\n"
                                          "Use Export everything to save all reports.")
            return
        stem = Path(self.folder.get()).name.replace(" ", "_")[:30]
        self.export([key], f"{stem}_{EXPORT_NAME.get(key, key)}.xlsx")

    def export_all(self):
        stem = Path(self.folder.get()).name.replace(" ", "_")[:30] if self.folder.get() else "Audit"
        self.export(None, f"{stem}_Audit_Output.xlsx")


def calc_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


if __name__ == "__main__":
    App().mainloop()
