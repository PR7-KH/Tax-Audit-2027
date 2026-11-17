"""Excel-style filter popup for a table heading.

Click a column heading: a small window opens under it with
    Sort A → Z / Sort Z → A
    Clear filter from this column
    Search box (words, or numbers such as >1000000  <=500  100..200)
    a tick list of every value in that column with its row count, (Select all)
    OK / Cancel
Only the values that are still ticked (and visible after a search) stay in the table,
exactly like Excel's AutoFilter. Nothing here knows about finance data; the caller passes
the values and gets the result back through callbacks.
"""
import re
import tkinter as tk

NAVY, ACCENT, LINE, MUTED = "#101f35", "#087f8c", "#d7e1eb", "#61748b"
ON, OFF = "☑", "☐"
OPS = re.compile(r"^\s*(>=|<=|>|<|=)?\s*(-?[\d,]*\.?\d+)\s*(?:\.\.\s*(-?[\d,]*\.?\d+))?\s*$")


def numeric_test(query):
    """A test for typed number conditions, or None when the text is not a condition."""
    m = OPS.match(query or "")
    if not m or not (m.group(1) or m.group(3)):
        return None
    a = float(m.group(2).replace(",", ""))
    if m.group(3) is not None:
        lo, hi = sorted((a, float(m.group(3).replace(",", ""))))
        return lambda v: lo <= v <= hi
    return {"=": lambda v: v == a, ">": lambda v: v > a, "<": lambda v: v < a,
            ">=": lambda v: v >= a, "<=": lambda v: v <= a}[m.group(1)]


def matches(key, query):
    """Does one value (as shown) pass the search box text?"""
    query = (query or "").strip()
    if not query:
        return True
    test = numeric_test(query)
    if test is not None:
        try:
            return test(float(key.replace(",", "")))
        except ValueError:
            return False
    k = key.casefold()
    return all(w in k for w in query.casefold().split())


class FilterPopup(tk.Toplevel):
    """values: [(value text, row count)]; selected: set of value texts currently shown, None = all.
    on_apply(values_or_None, text) — values None means no value filter; text is a search used only
    when the list is incomplete (too many values). on_sort(descending). on_clear()."""

    def __init__(self, master, x, y, column, values, selected, text="", complete=True,
                 on_apply=None, on_sort=None, on_clear=None):
        super().__init__(master)
        self.withdraw()
        self.overrideredirect(False)
        self.title(f"Filter: {column}")
        self.transient(master.winfo_toplevel())
        self.resizable(True, True)
        self.configure(bg="white", highlightthickness=1, highlightbackground=LINE)
        self.values, self.complete = values, complete
        self.on_apply, self.on_sort, self.on_clear = on_apply, on_sort, on_clear
        self.state = {k: (selected is None or k in selected) for k, _ in values}
        self.visible = []

        pad = dict(padx=8, pady=2)
        tk.Label(self, text=column, bg="white", fg=NAVY, font=("Helvetica", 12, "bold"), anchor="w").pack(fill="x", padx=8, pady=(8, 2))
        for label, desc in (("↑  Sort A → Z  /  smallest → largest", False), ("↓  Sort Z → A  /  largest → smallest", True)):
            b = tk.Label(self, text=label, bg="white", fg=NAVY, anchor="w", cursor="hand2", padx=4, pady=3)
            b.pack(fill="x", **pad)
            b.bind("<Button-1>", lambda e, d=desc: self._sort(d))
            b.bind("<Enter>", lambda e, w=b: w.config(bg="#e4edf5"))
            b.bind("<Leave>", lambda e, w=b: w.config(bg="white"))
        clear = tk.Label(self, text=f"✕  Clear filter from \"{column}\"", bg="white", fg="#9a293c", anchor="w", cursor="hand2", padx=4, pady=3)
        clear.pack(fill="x", **pad)
        clear.bind("<Button-1>", lambda e: self._clear())
        tk.Frame(self, bg=LINE, height=1).pack(fill="x", padx=8, pady=4)

        self.search = tk.StringVar(value=text)
        entry = tk.Entry(self, textvariable=self.search, relief="solid", borderwidth=1)
        entry.pack(fill="x", padx=8, pady=(2, 1), ipady=3)
        tk.Label(self, text="Search: words, or >1000000  <=500  100..200", bg="white", fg=MUTED,
                 font=("Helvetica", 9), anchor="w").pack(fill="x", padx=8)
        self.search.trace_add("write", lambda *a: self._refill(auto_tick=True))

        box = tk.Frame(self, bg="white")
        box.pack(fill="both", expand=True, padx=8, pady=4)
        self.listbox = tk.Listbox(box, activestyle="none", selectmode="browse", height=14, width=44,
                                  relief="solid", borderwidth=1, exportselection=False, font=("Helvetica", 11))
        ys = tk.Scrollbar(box, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=ys.set)
        ys.pack(side="right", fill="y")
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<ButtonRelease-1>", self._toggle_at)
        self.listbox.bind("<space>", self._toggle_selected)
        self.note = tk.Label(self, text="", bg="white", fg=MUTED, font=("Helvetica", 9), anchor="w", justify="left", wraplength=320)
        self.note.pack(fill="x", padx=8)

        buttons = tk.Frame(self, bg="white")
        buttons.pack(fill="x", padx=8, pady=8)
        for label, cmd, color in (("OK", self._ok, ACCENT), ("Cancel", self.destroy, "#5d7391")):
            b = tk.Label(buttons, text=label, bg=color, fg="white", padx=16, pady=5, cursor="hand2", font=("Helvetica", 11, "bold"))
            b.pack(side="right", padx=3)
            b.bind("<Button-1>", lambda e, c=cmd: c())
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self._refill(auto_tick=False)

        self.update_idletasks()
        w, h = max(340, self.winfo_reqwidth()), self.winfo_reqheight()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{max(0, min(x, sw - w - 10))}+{max(0, min(y, sh - h - 60))}")
        self.deiconify()
        self.lift()
        entry.focus_set()
        try:
            self.grab_set()
        except tk.TclError:
            pass

    # ---------------------------------------------------------------- list
    def _line(self, key, count, on):
        return f"{ON if on else OFF}  {key}    ({count:,})"

    def _refill(self, auto_tick):
        query = self.search.get()
        self.visible = [(k, n) for k, n in self.values if matches(k, query)]
        if auto_tick and query.strip():
            # Excel: typing a search ticks exactly the values it finds
            for k, _n in self.values:
                self.state[k] = False
            for k, _n in self.visible:
                self.state[k] = True
        elif auto_tick:
            for k, _n in self.values:
                self.state[k] = True
        self.listbox.delete(0, "end")
        all_on = bool(self.visible) and all(self.state[k] for k, _ in self.visible)
        self.listbox.insert("end", f"{ON if all_on else OFF}  (Select all{' search results' if query.strip() else ''})")
        for k, n in self.visible:
            self.listbox.insert("end", self._line(k, n, self.state[k]))
        note = f"{len(self.visible):,} of {len(self.values):,} values"
        if not self.complete:
            note += ". This column has too many different values to list them all; the search text is applied to every row."
        self.note.config(text=note)

    def _toggle_at(self, event):
        index = self.listbox.nearest(event.y)
        self._toggle(index)

    def _toggle_selected(self, _event):
        sel = self.listbox.curselection()
        if sel:
            self._toggle(sel[0])

    def _toggle(self, index):
        if index < 0:
            return
        if index == 0:
            new = not all(self.state[k] for k, _ in self.visible)
            for k, _ in self.visible:
                self.state[k] = new
            self._refill(auto_tick=False)
            return
        k, n = self.visible[index - 1]
        self.state[k] = not self.state[k]
        self.listbox.delete(index)
        self.listbox.insert(index, self._line(k, n, self.state[k]))
        all_on = all(self.state[v] for v, _ in self.visible)
        self.listbox.delete(0)
        self.listbox.insert(0, f"{ON if all_on else OFF}  (Select all{' search results' if self.search.get().strip() else ''})")

    # ---------------------------------------------------------------- actions
    def _ok(self):
        query = self.search.get().strip()
        if not self.complete:
            chosen = None if query else {k for k, _ in self.values if self.state[k]}
            if chosen is not None and len(chosen) == len(self.values):
                chosen = None
            result = (chosen, query)
        else:
            shown = {k for k, _ in self.visible} if query else {k for k, _ in self.values}
            chosen = {k for k in shown if self.state[k]}
            if not chosen:
                self.note.config(text="Tick at least one value, or press Cancel.", fg="#9a293c")
                return
            result = (None if len(chosen) == len(self.values) else chosen, "")
        self.destroy()
        if self.on_apply:
            self.on_apply(*result)

    def _sort(self, desc):
        self.destroy()
        if self.on_sort:
            self.on_sort(desc)

    def _clear(self):
        self.destroy()
        if self.on_clear:
            self.on_clear()
