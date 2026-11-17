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
import finance_view  # noqa: E402
import source_check  # noqa: E402
from input_tools import batch_convert_pdfs, organize_downloads
from toi_pdf import export_toi_pdfs, parse_steps, toi_workbook_fonts
from audit_tools_ui import AuditToolsPanel
from heading_filter import FilterPopup

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
    ("item", "fin_SALES", "Sales"),
    ("item", "fin_GL", "General ledger"),
    ("item", "fin_BANK", "Bank"),
    ("item", "fin_TB", "Trial balance"),
    ("item", "fin_AR", "A/R"),
    ("item", "fin_EXPENSE", "Expense"),
    ("item", "tax_risk", "Tax risk finding"),
    ("group", "", "AUDIT ANALYSIS"),
    ("item", "an_REVENUE", "Revenue"),
    ("item", "an_EXPENSE", "Expense"),
    ("item", "an_TOS", "TOS (salary tax)"),
    ("group", "", "SYSTEM"),
    ("item", "log", "Processing log"),
]
ACTION_TITLES = {"sources": "Source inspect", "build_all": "Build all data"}
SUBTITLE = {
    "sources": "Checks every folder and validates every file in one pass. ✓ ready, ⚠ note, ✗ not ready, — not provided.",
    "build_all": "Builds every tax report, loads GL / Sales / Bank and the tax risk finding, so every menu opens without waiting. Nothing is saved until you export.",
    "consolidate": "Consolidated monthly data with grand and year totals at the top (PR, SR, TOS, WHT).",
    "summary": "Year totals for each tax type.",
    "toi": "TOI steps 4 to 7 with Compare, % Change and Remark.",
    "fulltoi": "All TOI steps 1 to 17 in one sheet.",
    "mtr": "File-by-file detail for all four tax types.",
    "mtr_pr": "PR file list with year subtotals.", "mtr_sr": "SR file list with year subtotals.",
    "mtr_tos": "TOS file list with year subtotals.", "mtr_wht": "WHT file list with year subtotals.",
    "tps": "Tax paid by item for the two years, from the Info folder.",
    "dashboard": "Audit rules from the ADT sheet applied to TOI. Check the Processing log for corrected V5.1 numbers.",
    "fin_SALES": "Sales listing rows. Find, filter by month, click a heading to sort; Export saves the filtered rows.",
    "fin_GL": "General ledger lines. Find in any column (for example Account Name = sales), filter by month, sort, export.",
    "fin_BANK": "All bank statements in one aligned list, with a Type column that classifies each credit.",
    "tax_risk": "Total sales, cash-in and A/R, and expense compared across monthly tax, TOI, listing, GL, TB and bank, then the risk findings.",
    "log": "Warnings and corrections from the last build.",
    "an_REVENUE": "Sales per SR monthly returns, sales listing, GL, TB and TOI: year, month, invoice and buyer comparisons.",
    "an_EXPENSE": "Purchases per PR returns vs GL accounts payable, GL expense 5xxxx / 6xxxx vs TOI and TB, and TOS salary vs GL salary.",
    "an_TOS": "Salary tax risk from the monthly TOS staff lists: foreign staff ≤ 4,000,000 KHR, missing months, leavers, Khmer missing months.",
}
ANALYSIS_BUILDERS = {
    "an_REVENUE": ("revenue_expense", "build_revenue"),
    "an_EXPENSE": ("revenue_expense", "build_expense"),
    "an_TOS": ("tos_analysis", "build"),
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
        self.data_tables = {}      # (company folder, menu key) -> finance_view.Table
        self.data_state = {}
        self.inspection = None     # source_check.inspect() result for the selected folder
        self.source_per_file = False
        self.source_paths = []
        self.risk_cache = {}       # folder -> ready-to-draw table message
        self.risk_reports = {}     # folder -> risk_view report (for export)
        self.cons_tables = {}      # consolidated sheet name -> finance_view.Table
        self.cons_state = {}
        self.analyses = {}         # (company folder, an_ key) -> audit_analysis.Analysis
        self.filter_popup = None
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
        Btn(row, "Source inspect", self.reinspect, padx=9).pack(side="left", padx=(0, 4))
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
        source_bar = tk.Frame(self.source_controls, bg=BG)
        source_bar.pack(fill="x")
        Btn(source_bar, "Organize Download", self.organize_download, padx=8).pack(side="left", padx=(0, 4))
        Btn(source_bar, "Convert PDFs", self.convert_pending_pdfs, padx=8).pack(side="left", padx=4)
        self.per_file_btn = Btn(source_bar, "Show each file", self.toggle_source_files, color="#5d7391", padx=8)
        self.per_file_btn.pack(side="left", padx=4)
        Btn(source_bar, "Open selected", self.open_selected_source, color="#5d7391", padx=8).pack(side="left", padx=4)
        Btn(source_bar, "Build all data", lambda: self.open_report("build_all"), color=GREEN, hover="#14664a", padx=12).pack(side="right")
        self.source_banner = tk.Label(self.source_controls, text="", bg="#e3f2ec", fg="#12533f", font=(FONT, 11, "bold"),
                                      anchor="w", justify="left", padx=10, pady=5, wraplength=1000)
        self.source_banner.pack(fill="x", pady=(4, 0))
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
        small = (FONT, 10)
        self.cons_controls = tk.Frame(main, bg=BG)
        self.cons_column, self.cons_value, self.cons_month = tk.StringVar(value="All columns"), tk.StringVar(value="All"), tk.StringVar(value="All months")
        self.cons_choices = {}
        tk.Label(self.cons_controls, text="Filter", bg=BG, fg=TEXT, font=small).pack(side="left")
        self.cons_column_box = ttk.Combobox(self.cons_controls, textvariable=self.cons_column, state="readonly", width=18, font=small)
        self.cons_column_box.pack(side="left", padx=3)
        self.cons_column_box.bind("<<ComboboxSelected>>", lambda event: self.cons_column_changed())
        self.cons_value_box = ttk.Combobox(self.cons_controls, textvariable=self.cons_value, state="disabled", width=24, font=small)
        self.cons_value_box.pack(side="left", padx=3)
        self.cons_value_box.bind("<<ComboboxSelected>>", lambda event: self.apply_cons_filter())
        self.cons_month_box = ttk.Combobox(self.cons_controls, textvariable=self.cons_month, state="readonly", width=10, font=small)
        self.cons_month_box.pack(side="left", padx=3)
        self.cons_month_box.bind("<<ComboboxSelected>>", lambda event: self.apply_cons_filter())
        Btn(self.cons_controls, "Clear", self.clear_cons_filter, color="#5d7391", padx=6, pady=2).pack(side="left", padx=3)
        self.cons_sum = tk.Label(self.cons_controls, text="", bg=BG, fg="#12533f", font=(FONT, 10, "bold"), anchor="w", justify="left", wraplength=700)
        self.cons_sum.pack(side="left", fill="x", expand=True, padx=6)

        self.data_controls = tk.Frame(main, bg=BG)
        self.data_find, self.data_month = tk.StringVar(), tk.StringVar(value="All months")
        self.data_column, self.data_value = tk.StringVar(value="All columns"), tk.StringVar(value="All")
        self.data_choices = {}
        row1 = tk.Frame(self.data_controls, bg=BG)
        row1.pack(fill="x")
        tk.Label(row1, text="Column", bg=BG, fg=TEXT, width=7, anchor="w").pack(side="left")
        self.data_column_box = ttk.Combobox(row1, textvariable=self.data_column, state="readonly", width=26)
        self.data_column_box.pack(side="left", padx=4)
        self.data_column_box.bind("<<ComboboxSelected>>", lambda event: self.data_column_changed())
        tk.Label(row1, text="Value", bg=BG, fg=TEXT).pack(side="left", padx=(8, 0))
        self.data_value_box = ttk.Combobox(row1, textvariable=self.data_value, state="disabled", width=42)
        self.data_value_box.pack(side="left", padx=4)
        self.data_value_box.bind("<<ComboboxSelected>>", lambda event: self.apply_data_filter())
        tk.Label(row1, text="Month", bg=BG, fg=TEXT).pack(side="left", padx=(8, 0))
        self.data_month_box = ttk.Combobox(row1, textvariable=self.data_month, state="readonly", width=11)
        self.data_month_box.pack(side="left", padx=4)
        self.data_month_box.bind("<<ComboboxSelected>>", lambda event: self.apply_data_filter())
        row2 = tk.Frame(self.data_controls, bg=BG)
        row2.pack(fill="x", pady=(3, 0))
        tk.Label(row2, text="Find", bg=BG, fg=TEXT, width=7, anchor="w").pack(side="left")
        data_entry = tk.Entry(row2, textvariable=self.data_find, width=36)
        data_entry.pack(side="left", padx=4)
        data_entry.bind("<Return>", lambda event: self.apply_data_filter())
        Btn(row2, "Filter", self.apply_data_filter, padx=8).pack(side="left", padx=2)
        Btn(row2, "Clear", self.clear_data_filter, color="#5d7391", padx=8).pack(side="left", padx=2)
        tk.Label(row2, text="Tip: click any column heading to filter or sort it, like Excel.",
                 bg=BG, fg=MUTED).pack(side="left", padx=8)
        self.data_totals = tk.Label(self.data_controls, text="", bg="#e3f2ec", fg="#12533f", font=(FONT, 12, "bold"),
                                    anchor="w", justify="left", padx=10, pady=6, wraplength=1000)
        self.data_totals.pack(fill="x", pady=(5, 0))

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
        self.inspection = None
        self.data_tables.clear()
        self.risk_cache.clear()
        self.risk_reports.clear()
        self.cons_tables.clear()
        self.cons_state = {}
        self.__dict__.setdefault("analyses", {}).clear()
        self.close_filter_popup()

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
        for widget in (self.source_controls,self.search_controls,self.cons_controls,self.data_controls,self.tax_options,self.toi_pdf_options):widget.pack_forget()
        self.data_tables.clear();self.data_state={}
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
        if self.current and self.current.get("data"):
            self.decorate_data_headings()
        elif self.current and self.current.get("key") == "consolidate" and self.cons_state.get("sheet") in self.cons_tables:
            self.decorate_cons_headings()

    # ------------------------------------------------------------ finance data view (GL, Sales, Bank ...)
    def current_table(self):
        return self.data_tables.get((str(Path(self.folder.get())), (self.current or {}).get("key")))

    def show_data(self, key, view=None):
        table = self.data_tables[(str(Path(self.folder.get())), key)]
        self.current = {"key": key, "data": True, "view": view}
        self.data_state = {"page": 0, "sort": None, "indices": [], "filters": {}}
        self.data_find.set("")
        self.data_month.set("All months")
        self.data_column.set(table.default_field() or "All columns")
        self.data_column_box.config(values=["All columns"] + table.columns)
        self.data_month_box.config(values=["All months"] + table.months)
        self.refresh_data_values()
        self.source_controls.pack_forget()
        self.data_controls.pack(fill="x", before=self.picker, pady=3)
        self.btn_export.set_enabled(True)
        analysis = self.__dict__.get("analyses", {}).get((str(Path(self.folder.get())), key))
        if analysis is not None:
            self.set_picker(list(analysis.views), view, command=lambda n, k=key: self.show_analysis(k, n))
        else:
            self.set_picker([], "")
        self.apply_data_filter()

    def refresh_data_values(self):
        """Fill the Value list with every value of the chosen column (account names, buyers, bank accounts...)."""
        table = self.current_table()
        column = self.data_column.get()
        choices = table.choices(column) if table is not None and column != "All columns" else []
        self.data_choices = {label: value for label, value in choices}
        self.data_value.set("All")
        self.data_value_box.config(values=["All"] + [label for label, _ in choices],
                                   state="readonly" if choices else "disabled")

    def data_column_changed(self):
        self.refresh_data_values()
        self.apply_data_filter()

    def apply_data_filter(self):
        table = self.current_table()
        if table is None or not (self.current or {}).get("data"):
            return
        column, month, label = self.data_column.get(), self.data_month.get(), self.data_value.get()
        field = None if column == "All columns" else column
        value = self.data_choices.get(label) if label != "All" else None
        self.data_state["indices"] = table.view(
            self.data_find.get(), None, None if month == "All months" else month,
            self.data_state.get("sort"), field, value, self.data_state.get("filters"))
        self.data_state["page"] = 0
        self.render_data()

    def clear_data_filter(self):
        if not (self.current or {}).get("data"):
            return
        self.data_find.set("")
        self.data_month.set("All months")
        self.data_value.set("All")
        self.data_state["filters"] = {}
        self.apply_data_filter()

    def sort_data(self, column):
        if self.busy or not (self.current or {}).get("data"):
            return
        sort = self.data_state.get("sort")
        desc = bool(sort and sort[0] == column and not sort[1])
        self.data_state["sort"] = (column, desc)
        self.apply_data_filter()

    def render_data(self):
        table = self.current_table()
        indices = self.data_state["indices"]
        size = self.page_size
        pages = max(1, (len(indices) + size - 1) // size)
        page = max(0, min(self.data_state["page"], pages - 1))
        self.data_state["page"] = page
        shown = indices[page * size:(page + 1) * size]
        rows = [table.display(i) for i in shown]
        tags = None
        if table.row_tags:
            tags = [table.row_tags[i] if table.row_tags[i] in ("error", "warning", "total", "section") else ("odd" if n % 2 else "even")
                    for n, i in enumerate(shown)]
        fit = self.fit_columns
        if getattr(table, "wide", False):
            self.fit_columns = False     # many month columns: keep readable widths and scroll sideways
        try:
            self.draw(table.columns, rows, tags=tags)
        finally:
            self.fit_columns = fit
        self.page_label.config(text=f"Page {page + 1} / {pages}")
        label = next((m[2] for m in MENU if m[1] == self.current["key"]), "Data")
        if self.current.get("view"):
            label += " · " + self.current["view"]
        first = page * size + 1 if indices else 0
        text = f"{label}: rows {first:,}–{min(len(indices), (page + 1) * size):,} of {len(indices):,}"
        if len(indices) != len(table.rows):
            text += f" (filtered from {len(table.rows):,})"
        active = table.filter_text(self.data_state.get("filters"))
        self.data_totals.config(text="SUM of the filtered rows (all pages):   " + table.summary(indices)
                                + (f"\nHeading filters → {active}" if active else ""))
        if table.notes:
            text += " · Skipped: " + "; ".join(table.notes[:2])
        if getattr(table, "description", ""):
            text += " · " + table.description
        text += " · Click a column heading to filter or sort it (like Excel)."
        self.status(text)

    def decorate_data_headings(self):
        table = self.current_table()
        if table is None:
            return
        sort = self.data_state.get("sort")
        filters = self.data_state.get("filters") or {}
        for index, cid in enumerate(self.tree["columns"]):
            self.tree.heading(cid, text=self.heading_text(table.columns[index], index, sort, filters),
                              command=lambda i=index: self.open_data_filter(i))

    # ------------------------------------------------------------ Excel-style heading filters
    @staticmethod
    def heading_text(name, index, sort, filters):
        arrow = (" ▼" if sort[1] else " ▲") if sort and sort[0] == index else ""
        return ("⧩ " if filters.get(index) else "") + str(name) + arrow + "  ▾"

    def close_filter_popup(self):
        if self.__dict__.get("filter_popup") is not None:
            try:
                self.filter_popup.destroy()
            except tk.TclError:
                pass
        self.filter_popup = None

    def heading_position(self, index):
        """Screen x, y just below the heading of column index."""
        x = self.tree.winfo_rootx() - int(self.tree.xview()[0] * sum(self.tree.column(c, "width") for c in self.tree["columns"]))
        for cid in list(self.tree["columns"])[:index]:
            x += self.tree.column(cid, "width")
        return x + 4, self.tree.winfo_rooty() + 30

    def popup_filter(self, table, index, filters, base_indices, on_change, on_sort):
        """Open the filter list for one column: values come from rows passing the other filters."""
        if self.busy or table is None or index >= len(table.columns):
            return
        self.close_filter_popup()
        others = table.apply_filters(base_indices, filters, skip=index)
        values, complete = table.distinct(index, others)
        spec = filters.get(index) or {}

        def apply(values_set, text):
            if values_set is None and not text:
                filters.pop(index, None)
            else:
                filters[index] = {"values": values_set, "text": text}
            on_change()

        def clear():
            filters.pop(index, None)
            on_change()
        x, y = self.heading_position(index)
        self.filter_popup = FilterPopup(self, x, y, table.columns[index], values, spec.get("values"), spec.get("text", ""),
                                        complete, on_apply=apply, on_sort=lambda desc: on_sort(index, desc), on_clear=clear)

    def open_data_filter(self, index):
        table = self.current_table()
        if not (self.current or {}).get("data") or table is None:
            return
        filters = self.data_state.setdefault("filters", {})
        column, month, label = self.data_column.get(), self.data_month.get(), self.data_value.get()
        base = table.view(self.data_find.get(), None, None if month == "All months" else month, None,
                          None if column == "All columns" else column,
                          self.data_choices.get(label) if label != "All" else None)

        def sort(i, desc):
            self.data_state["sort"] = (i, desc)
            self.apply_data_filter()
        self.popup_filter(table, index, filters, base, self.apply_data_filter, sort)

    def decorate_cons_headings(self):
        table = self.cons_tables.get(self.cons_state.get("sheet"))
        if table is None:
            return
        sort = self.cons_state.get("sort")
        filters = self.cons_state.get("filters") or {}
        for index, cid in enumerate(self.tree["columns"]):
            if index >= len(table.columns):
                break
            current = self.tree.heading(cid, "text")
            name = table.columns[index] if self.cons_state.get("indices") is not None else (current or table.columns[index])
            name = name.replace("⧩ ", "").replace("  ▾", "").replace(" ▲", "").replace(" ▼", "")
            self.tree.heading(cid, text=self.heading_text(name, index, sort, filters), command=lambda i=index: self.open_cons_filter(i))

    def open_cons_filter(self, index):
        table = self.cons_tables.get(self.cons_state.get("sheet"))
        if table is None:
            return
        filters = self.cons_state.setdefault("filters", {})
        column, label, month = self.cons_column.get(), self.cons_value.get(), self.cons_month.get()
        base = table.view(month=None if month == "All months" else month,
                          field=None if column == "All columns" else column,
                          value=self.cons_choices.get(label) if label != "All" else None)

        def sort(i, desc):
            self.cons_state["sort"] = (i, desc)
            self.apply_cons_filter()
        self.popup_filter(table, index, filters, base, self.apply_cons_filter, sort)

    # ------------------------------------------------------------ audit analysis (Revenue, Expense, TOS)
    def job_analysis(self, key):
        import importlib
        module, fn = ANALYSIS_BUILDERS[key]
        folder = str(Path(self.job_settings[0]))
        analysis = getattr(importlib.import_module(module), fn)(folder)
        return ("analysis", folder, key, analysis)

    def show_analysis(self, key, view=None):
        if self.busy:
            return
        folder = str(Path(self.folder.get()))
        analysis = self.analyses.get((folder, key))
        if analysis is None:
            return
        view = view if view in analysis.views else next(iter(analysis.views))
        self.data_tables[(folder, key)] = analysis.views[view]
        self.show_data(key, view)

    def export_analysis(self):
        import audit_analysis
        key = (self.current or {}).get("key")
        analysis = self.analyses.get((str(Path(self.folder.get())), key))
        if analysis is None or self.busy:
            return
        company = Path(self.folder.get()).name
        stem = company.replace(" ", "_")[:30]
        path = filedialog.asksaveasfilename(
            title=f"Export {analysis.title} to Excel", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=str(Path.home() / "Documents"),
            initialfile=f"{stem}_{analysis.key.title()}_Analysis{('_' + analysis.period.split(' ')[-1]) if analysis.period else ''}.xlsx")
        if path:
            self.run_bg(lambda: (audit_analysis.save_workbook(analysis, path, company), ("exported", path))[1])

    def export_data(self):
        table = self.current_table()
        if table is None or self.busy:
            return
        stem = Path(self.folder.get()).name.replace(" ", "_")[:30]
        path = filedialog.asksaveasfilename(
            title="Export filtered rows to Excel", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=str(Path.home() / "Documents"), initialfile=f"{stem}_{self.current['key']}_filtered.xlsx")
        if not path:
            return
        indices = tuple(self.data_state["indices"])

        def job():
            table.export(path, indices)
            return ("exported", path)
        self.run_bg(job)

    def job_tax_risk(self):
        import risk_view
        folder = self.job_settings[0]
        report = risk_view.build_report(folder)
        rows, tags = risk_view.display_rows(report)
        notes = ("  Notes: " + "; ".join(report["notes"][:3])) if report["notes"] else ""
        msg = ("table", risk_view.COLUMNS, rows, f"Tax risk finding {report['year']}. Export this report saves it to Excel.{notes}", False, tags)
        self.risk_reports[str(Path(folder))] = report
        self.risk_cache[str(Path(folder))] = msg
        return msg

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
        if not self.busy and self.current and self.current.get("data") and self.current_table() is not None:
            self.data_state["page"] += direction
            self.render_data()
            return
        if not self.busy and self.current and self.current.get("key") == "consolidate" and self.cons_state.get("indices") is not None:
            self.cons_state["page"] += direction
            self.render_cons()
            return
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

    def set_picker(self, names, active, command=None):
        for w in self.picker.winfo_children():
            w.destroy()
        if len(names) < 2:
            return
        command = command or self.show_sheet
        row = tk.Frame(self.picker, bg=BG)
        row.pack(fill="x")
        width = max(600, self.main_panel.winfo_width() - 20)
        used = 0
        for n in names:
            text = "TOI compare" if n == "TOI" else n
            need = len(text) * 8 + 36
            if used and used + need > width:   # wrap long view lists onto a second line
                row = tk.Frame(self.picker, bg=BG)
                row.pack(fill="x", pady=(3, 0))
                used = 0
            b = Btn(row, text, lambda n=n: command(n), color=ACCENT if n == active else "#9fb2cc",
                    hover=NAVY, padx=10 if command is not self.show_sheet else 12, pady=4)
            b.pack(side="left", padx=(0, 6))
            used += need

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
        if self.current.get("key") == "consolidate" and self.cons_state.get("sheet") != name:
            self.setup_cons_filter(name)
        elif self.current.get("key") == "consolidate":
            self.cons_state["indices"] = None
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
        if self.current and self.current.get("key") == "sources":
            row = self.tree.identify_row(event.y)
            if row:
                self.tree.selection_set(row)
                self.open_selected_source()

    # ------------------------------------------------------------ source inspect
    def inspection_ready(self):
        return self.inspection is not None and self.inspection["folder"] == str(Path(self.folder.get()))

    def show_inspection(self):
        res = self.inspection
        rows, tags, paths = source_check.display(res, self.source_per_file)
        self.source_paths = paths
        self.set_picker([], "")
        self.draw(source_check.FILE_COLUMNS if self.source_per_file else source_check.COLUMNS, rows, align_right=False, tags=tags)
        bad = res["counts"].get(source_check.BAD, 0)
        self.source_banner.config(text=source_check.banner(res), bg="#fbe5e7" if bad else "#e3f2ec", fg="#9a293c" if bad else "#12533f")
        self.per_file_btn.config(text="Show folders" if self.source_per_file else "Show each file")
        note = self.conversion_note
        self.conversion_note = ""
        self.status(note or "Double-click a row to open that file or folder; a ✗ PDF row converts that PDF.")

    def reinspect(self):
        """Top-bar button: always look at the folder again (files may have been added)."""
        if self.busy:
            return
        self.inspection = None
        self.open_report("sources")

    def toggle_source_files(self):
        if self.busy or not self.inspection_ready():
            return
        self.source_per_file = not self.source_per_file
        if (self.current or {}).get("key") == "sources":
            self.show_inspection()

    def selected_source_path(self):
        selection = self.tree.selection()
        if not selection or (self.current or {}).get("key") != "sources":
            return None
        index = self.tree.index(selection[0])
        return self.source_paths[index] if index < len(self.source_paths) else None

    def open_selected_source(self):
        if self.busy:
            return
        path = self.selected_source_path()
        if path is None:
            self.status("Select a row first.")
            return
        if self.inspection_ready() and path in self.inspection["pending_pdfs"]:
            self.run_bg(self.job_convert, [path])
            return
        self.open_if(str(path))

    def convert_pending_pdfs(self):
        if self.busy or not self.need_folder():
            return
        if not self.inspection_ready():
            self.status("Open Source inspect first.")
            return
        pending = self.inspection["pending_pdfs"]
        if not pending:
            self.status("No PDF needs converting.")
            return
        self.run_bg(self.job_convert, list(pending))

    def job_inspect(self, import_results=None):
        return ("inspected", source_check.inspect(self.job_settings[0], import_results), None)

    def job_convert(self, paths):
        """Convert PDFs, then save each extraction as a standard Excel next to the PDF (same name), then re-inspect."""
        folder = Path(self.job_settings[0])
        results = batch_convert_pdfs(folder, paths, progress=lambda done, total: self.q.put(("batch_progress", done, total)))
        made, failed = [], []
        for top, parent, name, _rows, status, output in results:
            source = folder / parent / name
            if status.startswith("ERROR") or not output:
                failed.append(f"{name}: {status.replace('ERROR: ', '')}")
                continue
            key = next((k for k in engine.FINANCE_FOLDERS if engine.finance_folder(folder, k) and source.is_relative_to(engine.finance_folder(folder, k))), "")
            try:
                ok, why = source_check.standardize_extraction(folder / output, source.with_suffix(".xlsx"), key)
            except Exception as exc:
                ok, why = False, str(exc).splitlines()[0]
            (made if ok else failed).append(f"{name}: {why}")
        note = f"Converted {len(made)} of {len(results)} PDF(s) to standard Excel."
        if failed:
            note += " Not usable: " + "; ".join(failed[:3])
        self.inspection = None
        return ("inspected", source_check.inspect(folder, self.import_results), note)

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
        if self.cons_state.get("sheet") == name and self.cons_state.get("indices") is not None:
            self.setup_cons_filter(name)
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
        self.cons_controls.pack_forget()
        self.source_controls.pack_forget()
        self.data_controls.pack_forget()
        self.search_target = None
        self.search_hits = []
        self.search_label.config(text="")
        if key == "consolidate":
            self.search_controls.pack(fill="x", before=self.picker, pady=(3, 0))
            self.cons_controls.pack(fill="x", before=self.picker, pady=(2, 3))
        if key == "sources":
            self.source_controls.pack(fill="x", before=self.picker, pady=3)
        self.page_label.config(text="")
        self.tax_options.pack_forget()
        self.toi_pdf_options.pack_forget()
        self.dashboard_options.pack_forget()
        if key == "dashboard":
            self.dashboard_options.pack(side="left", padx=6)
        if key == "tps":
            self.tax_options.pack(fill="x", before=self.picker, pady=4)
        self.btn_export.set_enabled(key in engine.SECTIONS or key == "tax_risk")
        self.close_filter_popup()
        if key == "audit_tools":
            self.set_picker([],"");self.preview_card.pack_forget()
            if self.audit_panel is None:self.audit_panel=AuditToolsPanel(self.main_panel,self)
            self.audit_panel.pack(fill="both",expand=True,before=self.status_lbl,pady=4)
            self.audit_panel.load()
            return
        if not self.need_folder():
            return
        folder = str(Path(self.folder.get()))
        if key == "sources":
            if self.inspection_ready():
                self.show_inspection()
            else:
                self.run_bg(self.job_inspect, self.import_results)
            return
        if key == "tax_risk":
            if folder in self.risk_cache:
                self.q.put(self.risk_cache[folder])
            else:
                self.run_bg(self.job_tax_risk)
            return
        if key.startswith("an_"):
            if (folder, key) in self.analyses:
                self.show_analysis(key)
            else:
                self.run_bg(self.job_analysis, key)
            return
        if key.startswith("fin_"):
            if (folder, key) in self.data_tables:
                self.show_data(key)
            elif self.inspection_ready():
                self.show_not_ready(key)
            else:
                self.run_bg(self.job_inspect, self.import_results)
            return
        if key == "build_all":
            self.run_bg(self.job_build_all, self.import_results)
            return
        if self.built_ready():
            # Already built: open instantly, no background job.
            self.job_settings = self.settings_now()
            self.q.put(self.job_tois() if key == "tois" else self.job_report(key))
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

    def settings_now(self):
        try:
            return (self.folder.get(), self.months.get(), self.link_all.get())
        except tk.TclError:
            return None

    def built_ready(self):
        return self.built is not None and not self.partial_build and self.built[0] == self.settings_now()

    def show_not_ready(self, key):
        """A finance menu whose folder has no usable file: show its Source inspect lines."""
        label = engine.FINANCE_FOLDERS[key.split("_", 1)[1]][0]
        res = self.inspection
        rows = [r for r in res["rows"] if r["values"][1] == label] + [r for r in res["files"] if r["values"][1] == label]
        self.set_picker([], "")
        self.draw(source_check.FILE_COLUMNS, [r["values"] for r in rows], align_right=False,
                  tags=[source_check.TAG[r["mark"]] for r in rows])
        self.status(f"{label} has no usable Excel or CSV file yet. Fix it in Source inspect, then open this menu again.")

    # ------------------------------------------------------------ consolidated data filter
    def cons_table(self, name):
        """finance_view.Table of the data rows of one consolidated sheet (PR, SR, TOS, WHT)."""
        wb, cal = self.built[1], self.built[3]
        ws = wb[name]
        r = 1
        while r <= ws.max_row and str(ws.cell(r, 1).value or "").startswith(("Grand Total", "Year ")):
            r += 1
        top, sub = r, r + 1
        width = ws.max_column
        spans = {}
        for merged in ws.merged_cells.ranges:
            if merged.min_row <= top <= merged.max_row:
                for col in range(merged.min_col, merged.max_col + 1):
                    spans[col] = ws.cell(merged.min_row, merged.min_col).value
        labels, used = [], set()
        for col in range(1, width + 1):
            a = spans.get(col, ws.cell(top, col).value)
            b = ws.cell(sub, col).value if sub > top else None
            a = re.sub(r"\s+", " ", str(a)).strip() if a not in (None, "") else ""
            b = re.sub(r"\s+", " ", str(b)).strip() if b not in (None, "") and str(b) != a else ""
            label = (f"{a} / {b}" if a and b else a or b) or calc_letter(col)
            while label in used:
                label += " "
            used.add(label)
            labels.append(label)
        rows = []
        for number, values in enumerate(ws.iter_rows(min_row=sub + 1, max_col=width, values_only=True), start=sub + 1):
            if all(v in (None, "") for v in values):
                continue
            values = [cal.value(name, number, i + 1) if isinstance(v, str) and v.startswith("=") else v for i, v in enumerate(values)]
            rows.append([number] + [finance_view._clean(v) for v in values])
        table = finance_view.Table(["Row"] + labels, rows)
        layout = engine.cons.LAYOUTS.get(name.upper())
        if layout is not None:
            table.money = table.amount = list(range(layout.first_col, layout.last_col + 1))
        table.date_col = 2 if width >= 2 else None
        table.refresh_months()
        return table

    def setup_cons_filter(self, name):
        table = self.cons_tables.get(name)
        if table is None:
            self.cons_controls.pack_forget()
            return
        if not self.cons_controls.winfo_manager():
            self.cons_controls.pack(fill="x", before=self.picker, pady=(2, 3))
        self.cons_state = {"sheet": name, "indices": None, "page": 0, "filters": {}, "sort": None}
        self.cons_column.set("All columns")
        self.cons_value.set("All")
        self.cons_month.set("All months")
        self.cons_choices = {}
        self.cons_column_box.config(values=["All columns"] + table.columns[1:])
        self.cons_value_box.config(values=["All"], state="disabled")
        self.cons_month_box.config(values=["All months"] + table.months)
        self.cons_sum.config(text="Σ " + table.summary(range(len(table.rows)), compact=True))

    def cons_column_changed(self):
        table = self.cons_tables.get(self.cons_state.get("sheet"))
        if table is None:
            return
        column = self.cons_column.get()
        choices = table.choices(column) if column != "All columns" else []
        self.cons_choices = dict(choices)
        self.cons_value.set("All")
        self.cons_value_box.config(values=["All"] + [label for label, _ in choices], state="readonly" if choices else "disabled")
        self.apply_cons_filter()

    def apply_cons_filter(self):
        if self.busy:
            return
        name = self.cons_state.get("sheet")
        table = self.cons_tables.get(name)
        if table is None:
            return
        column, label, month = self.cons_column.get(), self.cons_value.get(), self.cons_month.get()
        value = self.cons_choices.get(label) if label != "All" else None
        filters = self.cons_state.get("filters") or {}
        sort = self.cons_state.get("sort")
        if value is None and month == "All months" and not filters and not sort:
            self.cons_state.update(indices=None, page=0)
            self.cons_sum.config(text="Σ " + table.summary(range(len(table.rows)), compact=True))
            self.show_sheet(name)
            return
        indices = table.view(month=None if month == "All months" else month, sort=sort,
                             field=None if column == "All columns" else column, value=value, filters=filters)
        self.cons_state.update(indices=indices, page=0)
        self.render_cons()

    def clear_cons_filter(self):
        name = self.cons_state.get("sheet")
        if name:
            self.setup_cons_filter(name)
            self.show_sheet(name)

    def render_cons(self):
        table = self.cons_tables[self.cons_state["sheet"]]
        indices = self.cons_state["indices"]
        size = self.page_size
        pages = max(1, (len(indices) + size - 1) // size)
        page = self.cons_state["page"] = max(0, min(self.cons_state["page"], pages - 1))
        rows = [table.display(i) for i in indices[page * size:(page + 1) * size]]
        self.draw(table.columns, rows, tags=["odd" if n % 2 else "even" for n in range(len(rows))])
        self.page_label.config(text=f"Page {page + 1} / {pages}")
        active = table.filter_text(self.cons_state.get("filters"))
        self.cons_sum.config(text="Σ filtered: " + table.summary(indices, compact=True) + (f"  ·  ⧩ {active}" if active else ""))
        self.status(f"{self.cons_state['sheet']}: {len(indices):,} of {len(table.rows):,} rows match the filter. Clear shows the full sheet with its totals.")

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
        if key == "consolidate":
            for name in names:
                if name not in self.cons_tables:
                    self.cons_tables[name] = self.cons_table(name)
        tables = dict.fromkeys(names)
        return ("sheets", key, tables)

    def job_build_all(self, import_results=None):
        """Everything the menus need, so they all open without waiting afterwards."""
        folder = self.job_settings[0]
        res = source_check.inspect(folder, import_results)
        notes, pages, rows = [], {}, []
        mark = source_check
        try:
            _k, wb, _log, _cal = self.ensure_built()
            for name in wb.sheetnames:
                pages[(name, 0)] = self.job_page(name, 0)
            for name in ("PR", "SR", "TOS", "WHT"):
                if name in wb.sheetnames:
                    self.cons_tables[name] = self.cons_table(name)
            rows.append([mark.READY, "Tax reports", f"{len(wb.sheetnames)} sheets ready: " + ", ".join(wb.sheetnames)])
        except Exception as exc:
            rows.append([mark.BAD, "Tax reports", "Not built: " + (str(exc).splitlines()[0] if str(exc) else type(exc).__name__)])
        for key, (label, _names) in engine.FINANCE_FOLDERS.items():
            table = res["tables"].get(key)
            line = next((r for r in res["rows"] if r["values"][1] == label), None)
            if table is not None:
                rows.append([mark.READY, label, f"{len(table.rows):,} rows ready"])
            else:
                rows.append([line["mark"] if line else mark.NONE, label, line["values"][5] if line else "Not provided"])
        risk = None
        try:
            risk = self.job_tax_risk()
            report = self.risk_reports.get(str(Path(folder)))
            high = sum(f["sev"] == "High" for f in report["findings"]) if report else 0
            rows.append([mark.READY, "Tax risk finding", f"{report['year']}: {len(report['findings'])} findings, {high} High" if report else "Ready"])
        except Exception as exc:
            rows.append([mark.BAD, "Tax risk finding", "Not built: " + (str(exc).splitlines()[0] if str(exc) else type(exc).__name__)])
        analyses = []
        for key in ANALYSIS_BUILDERS:
            label = next(m[2] for m in MENU if m[1] == key)
            try:
                _kind, afolder, _key, analysis = self.job_analysis(key)
                analyses.append((afolder, key, analysis))
                ready = not all(t == "error" for t in (analysis.views["Summary"].row_tags or [])) if "Summary" in analysis.views else False
                rows.append([mark.READY if ready else mark.NOTE, f"Analysis: {label}", analysis.headline()
                             + (" · " + analysis.notes[0] if analysis.notes else "")])
            except Exception as exc:
                rows.append([mark.BAD, f"Analysis: {label}", "Not built: " + (str(exc).splitlines()[0] if str(exc) else type(exc).__name__)])
        return ("built_all", res, pages, risk, rows, analyses)

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
                        self.conversion_note = f"Organize Download: {moved} moved, {len(file_rows)-moved} left in Download (see the Download rows)."
                    elif msg[1]:
                        self.conversion_note = msg[1][0][2].replace("INFO: ", "")
                    self.inspection = None
                    self.open_report("sources")
                elif kind == "toi_fonts":
                    self.draw(["TOI workbook", "Font families referenced", "Note"],msg[1])
                    self.status("On Mac, install matching fonts in Font Book and restart Excel. This lists workbook fonts; installed fonts are not verified.")
                elif kind == "toi_pdfs_done":
                    self.draw(["TOI workbook", "Steps", "Result", "PDF file"],msg[1])
                    successful = sum(r[2]=="EXPORTED" for r in msg[1])
                    self.status(f"TOI PDF export: {successful}/{len(msg[1])} workbooks exported. See {msg[2] if len(msg)>2 else 'TOI_PDFs'}.")
                elif kind == "inspected":
                    res = msg[1]
                    self.inspection = res
                    for fkey, table in res["tables"].items():
                        self.data_tables[(res["folder"], "fin_" + fkey)] = table
                    if msg[2]:
                        self.conversion_note = msg[2]
                    key = (self.current or {}).get("key")
                    if key == "sources":
                        self.show_inspection()
                    elif key and key.startswith("fin_") and res["folder"] == str(Path(self.folder.get())):
                        self.open_report(key)
                elif kind == "built_all":
                    _k, res, pages, risk, rows, analyses = msg
                    for afolder, akey, analysis in analyses:
                        self.analyses[(afolder, akey)] = analysis
                    self.inspection = res
                    for fkey, table in res["tables"].items():
                        self.data_tables[(res["folder"], "fin_" + fkey)] = table
                    for cache_key, page in pages.items():
                        self.page_cache[cache_key] = page
                    self.set_picker([], "")
                    self.draw(["", "Area", "Result"], rows, align_right=False,
                              tags=[source_check.TAG.get(r[0], "even") for r in rows])
                    self.status("Build all data finished. Every menu on the left now opens without waiting.")
                    self.notify_build_done()
                elif kind == "analysis":
                    _k, folder, key, analysis = msg
                    self.analyses[(folder, key)] = analysis
                    if (self.current or {}).get("key") == key and folder == str(Path(self.folder.get())):
                        self.show_analysis(key)
                        self.status(analysis.headline() + ". Pick a view above; click a column heading to filter. "
                                    "Export this report saves every view to Excel." +
                                    (" Notes: " + "; ".join(analysis.notes[:2]) if analysis.notes else ""))
                elif kind == "search":
                    self.search_hits = msg[1]
                    self.search_index = -1
                    if self.search_hits:
                        self.next_match()
                    else:
                        self.search_label.config(text="No matches in any tax sheet")
                elif kind == "page":
                    self.page_cache[(msg[1], msg[2])] = msg
                    while len(self.page_cache) > 160:
                        self.page_cache.popitem(last=False)
                    self.present_page(msg)
                elif kind == "table":
                    _k, cols, rows, note, _exp, *extra = msg
                    self.set_picker([], "")
                    self.draw(cols, rows, tags=extra[0] if extra else None)
                    self.status(note)
                    if _exp:self.notify_build_done()
                elif kind == "sheets":
                    _k, key, tables = msg
                    self.current = {"key": key, "tables": tables}
                    if tables:
                        self.show_sheet(next(iter(tables)))
                    else:
                        self.cons_controls.pack_forget()
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
        if str((self.current or {}).get("key") or "").startswith("an_"):
            self.export_analysis()
            return
        if (self.current or {}).get("data"):
            self.export_data()
            return
        key = (self.current or {}).get("key")
        if key == "tax_risk":
            self.export_risk()
            return
        if key not in engine.SECTIONS:
            messagebox.showinfo("Export", "Open a report first. This view has nothing to export.\n"
                                          "Use Export everything to save all reports.")
            return
        stem = Path(self.folder.get()).name.replace(" ", "_")[:30]
        self.export([key], f"{stem}_{EXPORT_NAME.get(key, key)}.xlsx")

    def export_risk(self):
        import risk_view
        report = self.risk_reports.get(str(Path(self.folder.get())))
        if report is None or self.busy:
            return
        stem = Path(self.folder.get()).name.replace(" ", "_")[:30]
        path = filedialog.asksaveasfilename(
            title="Export tax risk finding to Excel", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            initialdir=str(Path.home() / "Documents"), initialfile=f"{stem}_Tax_Risk_{report['year']}.xlsx")
        if path:
            self.run_bg(lambda: (risk_view.save_workbook(report, path), ("exported", path))[1])

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
