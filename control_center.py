"""Tax and Financial Control Center: one window for tax audit and financial audit.

Click a report to VIEW it. Nothing is written to disk until you press an Export button.
"""
from collections import OrderedDict, Counter
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import calc  # noqa: E402
import engine  # noqa: E402

NAVY, SIDE, SIDE_HOVER, ACCENT = "#101f35", "#14263f", "#203a58", "#087f8c"
GREEN, BG, CARD, TEXT, MUTED = "#167459", "#f0f4f8", "#ffffff", "#172b43", "#61748b"
LINE = "#d7e1eb"
FONT = "Helvetica"
NUMBER = re.compile(r"^-?[\d,]+(\.\d+)?%?$")

# (kind, key, label)   kind: group | item | sub
MENU = [
    ("group", "", "TAX AUDIT"),
    ("item", "sources", "Source overview"),
    ("item", "check_tax", "Validate tax files"),
    ("item", "build_all", "Build all data"),
    ("item", "consolidate", "Consolidated data"),
    ("item", "summary", "Summary MTR"),
    ("item", "toi", "TOI compare"),
    ("item", "fulltoi", "FULL TOI"),
    ("item", "mtr", "MTR detail  (all four)"),
    ("item", "tps", "Tax Paid Summary"),
    ("item", "dashboard", "Audit Dashboard"),
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
SUBTITLE = {
    "sources": "Company file counts by tax year and financial folder. No workbook contents are opened.",
    "check_tax": "Reads every monthly tax file with the same reader as V5.1 and reports problems before any build.",
    "build_all": "Builds every tax report in memory. Nothing is saved until you export.",
    "consolidate": "Consolidated monthly data with file subtotals and year totals (PR, SR, TOS, WHT).",
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
EXPORT_NAME = {"mtr": "MTR_detail"}


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
    if any(v in ("missing", "empty", "folder not found") for v in text):
        return "warning"
    if text and (text[0] in ("total", "grand total") or text[0].startswith("subtotal ")):
        return "total"
    if text and (text[0] == "year" or text[0].startswith(("folder:", "year:"))):
        return "section"
    return "odd" if index % 2 else "even"


def source_inventory(folder):
    """Fast counts only: do not open thousands of workbooks on folder selection."""
    root = Path(folder)
    rows = []
    for tax in engine.cons.LAYOUTS:
        directory = engine.find_dir(root, tax)
        counts = Counter(year for year, path in engine.cons.files_in(directory)) if directory else Counter()
        for year, count in sorted(counts.items()):
            rows.append([tax, year, count, "Found"])
        if not counts:
            rows.append([tax, "—", 0, "Missing" if directory is None else "Empty"])
    for name in ("TOI", "Info"):
        directory = engine.find_dir(root, name)
        count = len([p for p in directory.glob("*.xlsx") if not p.name.startswith((".", "~$"))]) if directory else 0
        rows.append([name, "All", count, "Found" if count else "Missing"])
    for key, (label, _) in engine.FINANCE_FOLDERS.items():
        directory = engine.finance_folder(root, key)
        count = sum(1 for p in directory.rglob("*") if p.is_file() and
                    p.suffix.lower() in engine.FIN_EXT and not any(
                        part.startswith((".", "~$")) for part in p.relative_to(directory).parts)) if directory else 0
        rows.append([label, "All", count, "Found" if count else "No files"])
    return rows


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
        self.built = None          # (key, workbook, log, calc)
        self.current = None        # dict describing the report on screen
        self.last_export = None
        self.menu_labels = {}
        self.busy = False
        self.page_cache = OrderedDict()
        self.sheet_pages = {}
        self.sheet_dimensions = {}
        self.page_size = 200
        self.zoom = 90
        self.fit_columns = True
        self.inventory = None
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
        head = tk.Frame(self, bg=NAVY)
        head.pack(fill="x")
        tk.Label(head, text="AUDIT  /  FINANCE", bg=NAVY, fg="white",
                 font=(FONT, 18, "bold")).pack(side="left", padx=22, pady=(14, 14))
        tk.Label(head, text="PRIVATE WORKSPACE  •  LOCAL FILES", bg=NAVY, fg="#9db4d3",
                 font=(FONT, 11)).pack(side="left", pady=(18, 0))

        bar = tk.Frame(self, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        bar.pack(fill="x")
        row = tk.Frame(bar, bg=CARD)
        row.pack(fill="x", padx=18, pady=(10, 4))
        tk.Label(row, text="Company folder", bg=CARD, fg=MUTED, font=(FONT, 11, "bold")).pack(side="left")
        tk.Entry(row, textvariable=self.folder, bg="#f7f9fc", fg=TEXT, insertbackground=TEXT,
                 relief="flat", highlightthickness=1, highlightbackground=LINE,
                 highlightcolor=ACCENT, font=(FONT, 12)).pack(side="left", fill="x", expand=True, padx=10, ipady=5)
        Btn(row, "Select folder", self.pick_folder).pack(side="left", padx=(0, 6))
        Btn(row, "Open folder", lambda: self.open_if(self.folder.get()), color="#5d7391").pack(side="left", padx=(0, 6))
        Btn(row, "Refresh", self.refresh, color="#5d7391").pack(side="left")
        opt = tk.Frame(bar, bg=CARD)
        opt.pack(fill="x", padx=18, pady=(0, 10))
        tk.Label(opt, text="SOURCE WORKSPACE  /  Tax returns & financial documents",
                 bg=CARD, fg=MUTED, font=(FONT, 10)).pack(side="left")
        tk.Checkbutton(opt, text="Fill ADT rows V5.1 leaves blank", variable=self.link_all, bg=CARD, fg=TEXT,
                       activebackground=CARD, activeforeground=TEXT, selectcolor=CARD, font=(FONT, 10),
                       command=self.invalidate).pack(side="right")

    def build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)
        side = tk.Frame(body, bg=SIDE, width=205)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        canvas = tk.Canvas(side, bg=SIDE, highlightthickness=0, width=205)
        inner = tk.Frame(canvas, bg=SIDE)
        canvas.create_window((0, 0), window=inner, anchor="nw", width=205)
        canvas.pack(fill="both", expand=True)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"))
        for kind, key, label in MENU:
            if kind == "group":
                tk.Label(inner, text=label, bg=SIDE, fg="#7f9cc4", font=(FONT, 10, "bold"),
                         anchor="w").pack(fill="x", padx=18, pady=(16, 4))
                continue
            lab = tk.Label(inner, text=("      " + label) if kind == "sub" else label, bg=SIDE, fg="#d5e1f1",
                           anchor="w", padx=18 if kind == "item" else 30, pady=6, cursor="hand2",
                           font=(FONT, 12 if kind == "item" else 11))
            lab.pack(fill="x")
            lab.bind("<Button-1>", lambda e, k=key: self.open_report(k))
            lab.bind("<Enter>", lambda e, l=lab: l.config(bg=SIDE_HOVER) if l is not self.active_label() else None)
            lab.bind("<Leave>", lambda e, l=lab: l.config(bg=SIDE) if l is not self.active_label() else None)
            self.menu_labels[key] = lab
        self.active_key = None

        main = tk.Frame(body, bg=BG)
        main.pack(side="left", fill="both", expand=True, padx=12, pady=8)
        top = tk.Frame(main, bg=BG)
        top.pack(fill="x")
        titles = tk.Frame(top, bg=BG)
        titles.pack(side="left", fill="x", expand=True)
        self.title_lbl = tk.Label(titles, text="Your audit workspace", bg=BG, fg=NAVY, font=(FONT, 17, "bold"), anchor="w")
        self.title_lbl.pack(fill="x")
        self.sub_lbl = tk.Label(titles, text="Choose a report on the left. Reports open for viewing; export when you want a file.",
                                bg=BG, fg=MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=700)
        self.sub_lbl.pack(fill="x")
        tools = tk.Frame(main, bg=BG)
        tools.pack(fill="x", pady=(6, 4))
        self.btn_export = Btn(tools, "Export this report", self.export_this, color=GREEN, hover="#14664a")
        self.btn_export.pack(side="left", padx=4)
        Btn(tools, "Export everything", self.export_all).pack(side="left", padx=4)
        Btn(tools, "Open last export", lambda: self.open_if(self.last_export), color="#5d7391").pack(side="left", padx=4)

        self.sub_lbl.bind("<Configure>", lambda event: self.sub_lbl.configure(
            wraplength=max(240, event.width - 12)))
        self.company_label = tk.Label(main, text="Select a company folder", bg=BG, fg=MUTED,
                                      font=(FONT, 10), anchor="w")
        self.company_label.pack(fill="x", pady=4)
        self.tax_options = tk.Frame(main, bg=BG)
        tk.Label(self.tax_options, text="Months of earlier year", bg=BG, fg=TEXT).pack(side="left")
        tk.Spinbox(self.tax_options, from_=1, to=12, width=4, textvariable=self.months).pack(side="left", padx=8)
        Btn(self.tax_options, "Apply", self.apply_months).pack(side="left")

        self.picker = tk.Frame(main, bg=BG)
        self.picker.pack(fill="x", pady=(10, 0))
        card = tk.Frame(main, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True, pady=(8, 8))
        preview_tools = tk.Frame(card, bg="#e4edf5")
        preview_tools.pack(fill="x")
        Btn(preview_tools, "−", lambda: self.change_zoom(-10)).pack(side="left", padx=2)
        self.zoom_label = tk.Label(preview_tools, text="90%", bg="#e4edf5", fg=NAVY)
        self.zoom_label.pack(side="left")
        Btn(preview_tools, "+", lambda: self.change_zoom(10)).pack(side="left", padx=2)
        Btn(preview_tools, "Fit columns", self.toggle_fit).pack(side="left", padx=4)
        Btn(preview_tools, "Previous", lambda: self.change_page(-1)).pack(side="right", padx=2)
        Btn(preview_tools, "Next", lambda: self.change_page(1)).pack(side="right", padx=2)
        self.page_label = tk.Label(preview_tools, text="", bg="#e4edf5", fg=NAVY)
        self.page_label.pack(side="right", padx=8)
        self.tree = ttk.Treeview(card, show="headings", selectmode="browse")
        for tag, options in {
            "even": {"background": "#ffffff"},
            "odd": {"background": "#f2f6fa"},
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
        self.status_lbl = tk.Label(main, text="", bg=BG, fg=MUTED, anchor="w", font=(FONT, 10))
        self.status_lbl.pack(fill="x")

    # ------------------------------------------------------------ helpers
    def active_label(self):
        return self.menu_labels.get(self.active_key)

    def mark_active(self, key):
        for k, lab in self.menu_labels.items():
            lab.config(bg=SIDE, fg="#d5e1f1")
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
            self.open_report("sources")

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

    def refresh(self):
        if self.busy:
            return
        self.invalidate()
        if self.current and self.current.get("key"):
            self.open_report(self.current["key"])

    # ------------------------------------------------------------ table drawing
    def draw(self, columns, rows, align_right=True, tags=None):
        self.tree.delete(*self.tree.get_children())
        ids = [f"c{i}" for i in range(len(columns))]
        self.tree["columns"] = ids
        scale = self.zoom / 100
        available = max(300, self.tree.winfo_width() - 24)
        for i, (cid, name) in enumerate(zip(ids, columns)):
            cells = [str(r[i]) for r in rows[:100] if i < len(r) and r[i] not in (None, "")]
            numeric = align_right and cells and sum(bool(NUMBER.match(c)) for c in cells) / len(cells) > 0.6
            natural = max(55, min(max([len(name)] + [len(c) for c in cells]) * 7 + 20, 400))
            width = max(45, available // max(1, len(ids))) if self.fit_columns else natural
            self.tree.heading(cid, text=name)
            self.tree.column(cid, width=int(width * scale), minwidth=30,
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
        self.invalidate()
        self.open_report("tps")

    def change_zoom(self, delta):
        self.zoom = max(50, min(150, self.zoom + delta))
        self.zoom_label.config(text=f"{self.zoom}%")
        style = ttk.Style(self)
        style.configure("Treeview", font=(FONT, max(7, round(11*self.zoom/100)), "bold"),
                        rowheight=max(18, round(30*self.zoom/100)))
        style.configure("Treeview.Heading", font=(FONT, max(7, round(11*self.zoom/100)), "bold"))
        if self.current and self.current.get("sheet") and not self.busy:
            self.show_sheet(self.current["sheet"])

    def toggle_fit(self):
        self.fit_columns = not self.fit_columns
        self.change_zoom(0)

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
            rows.append([r] + values)
            tags.append(tag)
        return ("page", name, page, ncols, total, rows, tags)

    def present_page(self, msg):
        _, name, page, ncols, total, rows, tags = msg
        self.draw(["#"] + [calc_letter(i+1) for i in range(ncols)], rows, tags=tags)
        self.page_label.config(text=f"Page {page+1} / {max(1, (total+self.page_size-1)//self.page_size)}")
        self.status(f"{name}: rows {page*self.page_size+1:,}–{min(total,(page+1)*self.page_size):,} of {total:,}. All rows available using Next / Previous.")

    def set_picker(self, names, active):
        for w in self.picker.winfo_children():
            w.destroy()
        if len(names) < 2:
            return
        for n in names:
            b = Btn(self.picker, n, lambda n=n: self.show_sheet(n), color=ACCENT if n == active else "#9fb2cc",
                    hover=NAVY, padx=12, pady=4)
            b.pack(side="left", padx=(0, 6))

    def show_sheet(self, name):
        if self.busy:
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

    # ------------------------------------------------------------ opening a report
    def open_report(self, key):
        if self.busy:
            self.status("Please wait, still working…")
            return
        label = next(m[2] for m in MENU if m[1] == key)
        self.mark_active(key)
        self.title_lbl.config(text=label.strip())
        self.sub_lbl.config(text=SUBTITLE.get(key, ""))
        self.current = {"key": key}
        self.page_label.config(text="")
        self.tax_options.pack_forget()
        if key == "tps":
            self.tax_options.pack(fill="x", before=self.picker, pady=4)
        self.btn_export.set_enabled(key in engine.SECTIONS)
        if key == "sources":
            if self.need_folder():
                self.run_bg(self.job_inventory)
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
    def ensure_built(self):
        key = self.job_settings
        if self.built is None or self.built[0] != key:
            wb, log = engine.build_workbook(key[0], key[1], key[2])
            self.page_cache.clear()
            self.sheet_pages.clear()
            self.sheet_dimensions.clear()
            self.built = (key, wb, log, calc.Calc(wb))
        return self.built

    def job_inventory(self):
        folder = self.job_settings[0]
        if self.inventory is None or self.inventory[0] != folder:
            self.inventory = (folder, source_inventory(folder))
        return ("table", ["Source", "Year / scope", "Files", "Status"], self.inventory[1],
                "File inventory only. Use Validate tax files to check headings, values and formula caches.", False)

    def job_check_tax(self):
        rows = engine.check_sources(self.job_settings[0])
        bad = sum(1 for r in rows if r[4] != "READY")
        return ("table", ["Tax", "Year", "Source file", "Rows", "Result"], [list(r) for r in rows],
                f"{len(rows) - bad} of {len(rows)} items ready.", False)

    def job_finance(self, key):
        summary, files = engine.check_finance(self.job_settings[0])
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
        built = self.ensure_built()
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
        tables = dict.fromkeys(names)
        return ("sheets", key, tables)

    # ------------------------------------------------------------ queue results -> screen
    def poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                self.busy = False
                self.config(cursor="")
                kind = msg[0]
                if kind == "page":
                    self.page_cache[(msg[1], msg[2])] = msg
                    while len(self.page_cache) > 32:
                        self.page_cache.popitem(last=False)
                    self.present_page(msg)
                elif kind == "table":
                    _k, cols, rows, note, _exp = msg
                    self.set_picker([], "")
                    self.draw(cols, rows)
                    self.status(note)
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
