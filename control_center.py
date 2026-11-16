"""Tax and Financial Control Center: one window for tax audit and financial audit.

Click a report to VIEW it. Nothing is written to disk until you press an Export button.
"""
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

NAVY, SIDE, SIDE_HOVER, ACCENT = "#12294a", "#16335c", "#21487a", "#2f80d0"
GREEN, BG, CARD, TEXT, MUTED = "#1a7f5a", "#eef2f7", "#ffffff", "#1b2733", "#5d6b7e"
LINE = "#d3dbe6"
FONT = "Helvetica"
NUMBER = re.compile(r"^-?[\d,]+(\.\d+)?%?$")

# (kind, key, label)   kind: group | item | sub
MENU = [
    ("group", "", "TAX AUDIT"),
    ("item", "check_tax", "Check source files"),
    ("item", "build_all", "Build all data"),
    ("item", "consolidate", "Consolidated data"),
    ("item", "summary", "Summary MTR"),
    ("item", "toi", "TOI compare"),
    ("item", "fulltoi", "FULL TOI"),
    ("item", "mtr", "MTR detail  (all four)"),
    ("sub", "mtr_pr", "PR"),
    ("sub", "mtr_sr", "SR"),
    ("sub", "mtr_tos", "TOS"),
    ("sub", "mtr_wht", "WHT"),
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
        self.setup_styles()
        self.build_header()
        self.build_body()
        self.status("Select the company folder to begin.")
        self.after(120, self.poll)

    # ------------------------------------------------------------ layout
    def setup_styles(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        st.configure("Treeview", background=CARD, fieldbackground=CARD, foreground=TEXT,
                     rowheight=25, borderwidth=0, font=(FONT, 11))
        st.configure("Treeview.Heading", background="#dce6f1", foreground=NAVY,
                     font=(FONT, 11, "bold"), relief="flat", padding=(6, 6))
        st.map("Treeview", background=[("selected", ACCENT)], foreground=[("selected", "white")])
        st.configure("Vertical.TScrollbar", background="#c4cedb", troughcolor=BG, borderwidth=0)
        st.configure("Horizontal.TScrollbar", background="#c4cedb", troughcolor=BG, borderwidth=0)

    def build_header(self):
        head = tk.Frame(self, bg=NAVY)
        head.pack(fill="x")
        tk.Label(head, text="TAX & FINANCIAL CONTROL CENTER", bg=NAVY, fg="white",
                 font=(FONT, 18, "bold")).pack(side="left", padx=22, pady=(14, 14))
        tk.Label(head, text="Tax audit  •  Financial audit  •  Offline", bg=NAVY, fg="#9db4d3",
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
        tk.Label(opt, text="One folder holds both: PR, SR, TOS, WHT, TOI, Info  and  GL, TB, Bank, Sales",
                 bg=CARD, fg=MUTED, font=(FONT, 10)).pack(side="left")
        tk.Checkbutton(opt, text="Fill ADT rows V5.1 leaves blank", variable=self.link_all, bg=CARD, fg=TEXT,
                       activebackground=CARD, activeforeground=TEXT, selectcolor=CARD, font=(FONT, 10),
                       command=self.invalidate).pack(side="right")
        tk.Spinbox(opt, from_=1, to=12, width=3, textvariable=self.months, bg="#f7f9fc", fg=TEXT,
                   buttonbackground="#f7f9fc", insertbackground=TEXT, font=(FONT, 10),
                   command=self.invalidate).pack(side="right", padx=(6, 18))
        tk.Label(opt, text="Months of earlier year (Tax Paid Summary)", bg=CARD, fg=TEXT,
                 font=(FONT, 10)).pack(side="right")

    def build_body(self):
        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True)
        side = tk.Frame(body, bg=SIDE, width=250)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        canvas = tk.Canvas(side, bg=SIDE, highlightthickness=0, width=250)
        inner = tk.Frame(canvas, bg=SIDE)
        canvas.create_window((0, 0), window=inner, anchor="nw", width=250)
        canvas.pack(fill="both", expand=True)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(-1 if e.delta > 0 else 1, "units"))
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
        main.pack(side="left", fill="both", expand=True, padx=18, pady=14)
        top = tk.Frame(main, bg=BG)
        top.pack(fill="x")
        titles = tk.Frame(top, bg=BG)
        titles.pack(side="left", fill="x", expand=True)
        self.title_lbl = tk.Label(titles, text="Welcome", bg=BG, fg=NAVY, font=(FONT, 20, "bold"), anchor="w")
        self.title_lbl.pack(fill="x")
        self.sub_lbl = tk.Label(titles, text="Choose a report on the left. Reports open for viewing; export when you want a file.",
                                bg=BG, fg=MUTED, font=(FONT, 11), anchor="w", justify="left", wraplength=700)
        self.sub_lbl.pack(fill="x")
        tools = tk.Frame(top, bg=BG)
        tools.pack(side="right")
        self.btn_export = Btn(tools, "Export this report", self.export_this, color=GREEN, hover="#14664a")
        self.btn_export.pack(side="left", padx=4)
        Btn(tools, "Export everything", self.export_all).pack(side="left", padx=4)
        Btn(tools, "Open last export", lambda: self.open_if(self.last_export), color="#5d7391").pack(side="left", padx=4)

        self.picker = tk.Frame(main, bg=BG)
        self.picker.pack(fill="x", pady=(10, 0))
        card = tk.Frame(main, bg=CARD, highlightthickness=1, highlightbackground=LINE)
        card.pack(fill="both", expand=True, pady=(8, 8))
        self.tree = ttk.Treeview(card, show="headings", selectmode="browse")
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
        if d:
            self.folder.set(d)
            self.invalidate()
            self.status(f"Folder selected: {Path(d).name}")

    def need_folder(self):
        if not Path(self.folder.get()).is_dir():
            messagebox.showwarning("Select folder", "Click Select folder and choose the company folder first.")
            return False
        return True

    def invalidate(self):
        self.built = None

    def refresh(self):
        self.invalidate()
        if self.current and self.current.get("key"):
            self.open_report(self.current["key"])

    # ------------------------------------------------------------ table drawing
    def draw(self, columns, rows, align_right=True):
        self.tree.delete(*self.tree.get_children())
        ids = [f"c{i}" for i in range(len(columns))]
        self.tree["columns"] = ids
        sample = rows[:300]
        for i, (cid, name) in enumerate(zip(ids, columns)):
            cells = [str(r[i]) for r in sample if i < len(r) and str(r[i]) != ""]
            numeric = align_right and cells and sum(bool(NUMBER.match(c)) for c in cells) / len(cells) > 0.6
            longest = max([len(name)] + [len(c) for c in cells]) if cells else len(name)
            self.tree.heading(cid, text=name)
            self.tree.column(cid, width=max(48, min(longest * 8 + 22, 420)), minwidth=40,
                             anchor="e" if numeric else "w", stretch=False)
        for r in rows:
            self.tree.insert("", "end", values=[("" if v is None else v) for v in r])

    def draw_sheet(self, rows, ncols):
        letters = [calc_letter(i + 1) for i in range(ncols)]
        self.draw(["#"] + letters, [[n + 1] + r for n, r in enumerate(rows)])

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
        tables = self.current["tables"]
        rows, ncols = tables[name]
        self.current["sheet"] = name
        self.draw_sheet(rows, ncols)
        self.set_picker(list(tables), name)
        self.status(f"{name}: {len(rows)} rows. Press Export to save as Excel.")

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
        self.btn_export.set_enabled(key in engine.SECTIONS)
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
        key = (self.folder.get(), self.months.get(), self.link_all.get())
        if self.built is None or self.built[0] != key:
            wb, log = engine.build_workbook(key[0], key[1], key[2])
            self.built = (key, wb, log, calc.Calc(wb))
        return self.built

    def job_check_tax(self):
        rows = engine.check_sources(self.folder.get())
        bad = sum(1 for r in rows if r[4] != "READY")
        return ("table", ["Tax", "Year", "Source file", "Rows", "Result"], [list(r) for r in rows],
                f"{len(rows) - bad} of {len(rows)} items ready.", False)

    def job_finance(self, key):
        summary, files = engine.check_finance(self.folder.get())
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
        tables = {n: calc.sheet_table(wb, n, cal) for n in names}
        return ("sheets", key, tables)

    # ------------------------------------------------------------ queue results -> screen
    def poll(self):
        try:
            while True:
                msg = self.q.get_nowait()
                self.busy = False
                self.config(cursor="")
                kind = msg[0]
                if kind == "table":
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
