"""
hours_ui.py
===========
Classe HoursUI: costruisce tutti i widget dell'app.
Nessuna logica applicativa, nessun import da db/excel.
"""

import tkinter as tk
from tkinter import ttk

# ── Palette ───────────────────────────────────────────────────────────────────
ACCENT       = "#9a3412"
ACCENT_MID   = "#c2410c"
ACCENT_LIGHT = "#fed7aa"
BG           = "#f4f5f7"
BG_CARD      = "#ffffff"
BG_THEAD     = "#f8f9fb"
TEXT         = "#111827"
TEXT_DIM     = "#6b7280"
TEXT_WHITE   = "#ffffff"
SUCCESS      = "#166534"
BORDER       = "#e2e5ea"
ROW_ODD      = "#ffffff"
ROW_EVEN     = "#f8f9fb"
ERR          = "#991b1b"
WARN         = "#92400e"

COLUMNS = (
    "Resource", "Job", "Phase", "Description",
    "Notes", "Date", "Created On", "Hours",
)
COL_WIDTHS = {
    "Resource": 140, "Job": 120, "Phase": 160,
    "Description": 150, "Notes": 230,
    "Date": 90, "Created On": 90, "Hours": 60,
}


class HoursUI:
    """
    Costruisce l'intera UI dell'app.
    I widget pubblici vengono usati da HoursViewer per agganciare callbacks.

    Attributi pubblici:
        root
        tree
        lbl_status, lbl_update, lbl_summary
        lbl_risorsa_hdr
        lbl_check, btn_detail
        cmb_gruppo, cmb_name
        ent_cdcom, ent_from, ent_to
        btn_cerca, btn_reload
        frm_banner, lbl_banner
    """

    def __init__(self, parent_root, title: str = "Engineering Hours Viewer"):
        if parent_root:
            self.root = tk.Toplevel(parent_root)
        else:
            self.root = tk.Tk()

        self.root.title(title)
        self.root.geometry("1200x760")
        self.root.minsize(900, 560)
        self.root.configure(bg=BG)

        self._setup_styles()
        self._build_header()
        self._build_toolbar()
        self._build_filters()
        self._build_check_bar()
        self._build_banner()
        self._build_summary()
        self._build_table()
        self._build_footer()

    # ── Styles ───────────────────────────────────────────────────────────────

    def _setup_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("PT.Treeview",
            background=BG_CARD, foreground=TEXT,
            fieldbackground=BG_CARD, rowheight=28,
            font=("Segoe UI", 10), borderwidth=0, relief="flat")
        style.configure("PT.Treeview.Heading",
            background=BG_THEAD, foreground=TEXT_DIM,
            font=("Segoe UI", 9, "bold"),
            relief="flat", borderwidth=0, padding=(8, 7))
        style.map("PT.Treeview",
            background=[("selected", ACCENT)],
            foreground=[("selected", TEXT_WHITE)])
        style.map("PT.Treeview.Heading",
            background=[("active", BORDER)],
            relief=[("active", "flat")])

    # ── Header ───────────────────────────────────────────────────────────────

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=ACCENT, height=52)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text=" ENG ", bg=ACCENT_MID, fg=TEXT_WHITE,
                 font=("Segoe UI", 11, "bold")).pack(
                     side="left", padx=(18, 0), pady=11)
        tk.Label(hdr, text="Engineering Hours Viewer",
                 bg=ACCENT, fg=TEXT_WHITE,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=(10, 0))
        self.lbl_risorsa_hdr = tk.Label(hdr, text="· All Groups",
                 bg=ACCENT, fg=ACCENT_LIGHT,
                 font=("Segoe UI", 10))
        self.lbl_risorsa_hdr.pack(side="left", padx=(4, 0))

        self.lbl_status = tk.Label(hdr, text="● Loading...",
                 font=("Segoe UI", 10), fg=TEXT_WHITE, bg=ACCENT)
        self.lbl_status.pack(side="right", padx=(0, 20))

        tk.Frame(self.root, bg=ACCENT_MID, height=1).pack(fill="x")

    # ── Toolbar (refresh + update info) ──────────────────────────────────────

    def _build_toolbar(self):
        bar = tk.Frame(self.root, bg=BG_CARD, height=40)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        self.lbl_status2 = tk.Label(bar, text="",
                 font=("Segoe UI", 9, "bold"), fg=ACCENT_MID, bg=BG_CARD)
        self.lbl_status2.pack(side="left", padx=(16, 12), fill="y")

        self.lbl_update = tk.Label(bar, text="",
                 font=("Segoe UI", 8), fg=TEXT_DIM, bg=BG_CARD)
        self.lbl_update.pack(side="left", fill="y")

        self.btn_reload = tk.Button(bar, text="⟳  Refresh from DB",
                 font=("Segoe UI", 9), bg=ACCENT, fg=TEXT_WHITE,
                 relief="flat", cursor="hand2", padx=12, pady=4,
                 state="disabled")
        self.btn_reload.pack(side="right", padx=16, pady=5)

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

    # ── Filters ──────────────────────────────────────────────────────────────

    def _build_filters(self):
        ff = tk.Frame(self.root, bg=BG_CARD, padx=18, pady=12)
        ff.pack(fill="x")

        def lbl(t, r, c, **kw):
            tk.Label(ff, text=t, font=("Segoe UI", 9),
                     fg=TEXT_DIM, bg=BG_CARD).grid(
                         row=r, column=c, sticky="w", **kw)

        def entry(w):
            return tk.Entry(ff, width=w, font=("Segoe UI", 10),
                            bg=BG_CARD, fg=TEXT, insertbackground=TEXT,
                            relief="solid", highlightthickness=1,
                            highlightbackground=BORDER, highlightcolor=ACCENT)

        def btn(text, bg, fg, r, c, **kw):
            b = tk.Button(ff, text=text, font=("Segoe UI", 9),
                          bg=bg, fg=fg, relief="flat", cursor="hand2",
                          padx=8, pady=4)
            b.grid(row=r, column=c, **kw)
            return b

        # ── Row 0: Group, Resource, Job, Search, Reset, Export ──────────────
        lbl("Group", 0, 0, padx=(0, 6))
        self.cmb_gruppo = ttk.Combobox(ff, state="disabled", width=18,
                                       font=("Segoe UI", 10))
        self.cmb_gruppo.grid(row=0, column=1, sticky="w", padx=(0, 14))

        lbl("Resource", 0, 2, padx=(0, 6))
        self.cmb_name = ttk.Combobox(ff, state="disabled", width=20,
                                     font=("Segoe UI", 10))
        self.cmb_name.grid(row=0, column=3, sticky="w", padx=(0, 14))

        lbl("Job", 0, 4, padx=(0, 4))
        self.ent_cdcom = entry(16)
        self.ent_cdcom.grid(row=0, column=5, sticky="w", padx=(0, 10))

        self.btn_cerca = tk.Button(ff, text="Search",
                 font=("Segoe UI", 10), bg=ACCENT, fg=TEXT_WHITE,
                 relief="flat", cursor="hand2", padx=12, pady=4, state="disabled")
        self.btn_cerca.grid(row=0, column=6, padx=(0, 8))

        self._btn_reset = btn("Reset", "#e5e7eb", TEXT_DIM, 0, 7, padx=(0, 8))
        self._btn_export = btn("Export XLSX", SUCCESS, TEXT_WHITE, 0, 8)

        # ── Row 1: From date, To date ────────────────────────────────────────
        # Usiamo frame dedicati per evitare che le colonne della riga 0
        # creino spazi enormi nella riga 1
        date_frame = tk.Frame(ff, bg=BG_CARD)
        date_frame.grid(row=1, column=0, columnspan=9, sticky="w", pady=(8, 0))

        tk.Label(date_frame, text="From", font=("Segoe UI", 9),
                 fg=TEXT_DIM, bg=BG_CARD).pack(side="left", padx=(0, 4))
        self.ent_from = tk.Entry(date_frame, width=10, font=("Segoe UI", 10),
                                 bg=BG_CARD, fg=TEXT, insertbackground=TEXT,
                                 relief="solid", highlightthickness=1,
                                 highlightbackground=BORDER, highlightcolor=ACCENT)
        self.ent_from.pack(side="left")
        self.btn_cal_from = tk.Button(date_frame, text="📅",
                 font=("Segoe UI", 9), bg="#e5e7eb", fg=TEXT_DIM,
                 relief="flat", cursor="hand2", padx=3, pady=3)
        self.btn_cal_from.pack(side="left", padx=(2, 20))

        tk.Label(date_frame, text="To", font=("Segoe UI", 9),
                 fg=TEXT_DIM, bg=BG_CARD).pack(side="left", padx=(0, 4))
        self.ent_to = tk.Entry(date_frame, width=10, font=("Segoe UI", 10),
                               bg=BG_CARD, fg=TEXT, insertbackground=TEXT,
                               relief="solid", highlightthickness=1,
                               highlightbackground=BORDER, highlightcolor=ACCENT)
        self.ent_to.pack(side="left")
        self.btn_cal_to = tk.Button(date_frame, text="📅",
                 font=("Segoe UI", 9), bg="#e5e7eb", fg=TEXT_DIM,
                 relief="flat", cursor="hand2", padx=3, pady=3)
        self.btn_cal_to.pack(side="left", padx=(2, 0))

        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")

        # ── Hours check bar ───────────────────────────────────────────────────────

    def _build_check_bar(self):
        chk = tk.Frame(self.root, bg=BG_CARD, height=30)
        chk.pack(fill="x", padx=18)
        chk.pack_propagate(False)

        self.lbl_check = tk.Label(chk,
                 text="Select a resource and period to check hours",
                 font=("Segoe UI", 9), fg=TEXT_DIM, bg=BG_CARD, anchor="w")
        self.lbl_check.pack(side="left", fill="y", padx=(2, 10))

        self.btn_detail = tk.Button(chk, text="View Details",
                 font=("Segoe UI", 9), bg=ACCENT, fg=TEXT_WHITE,
                 relief="flat", cursor="hand2", padx=8, pady=2)
        self.btn_detail.pack(side="left")
        self.btn_detail.pack_forget()

    # ── Update banner ─────────────────────────────────────────────────────────

    def _build_banner(self):
        self.frm_banner = tk.Frame(self.root, bg="#fef9c3")
        self.lbl_banner = tk.Label(self.frm_banner, text="",
                 font=("Segoe UI", 9), fg=WARN, bg="#fef9c3", anchor="w")
        self.lbl_banner.pack(side="left", padx=16, pady=6)
        tk.Button(self.frm_banner, text="Update now",
                 font=("Segoe UI", 9), bg=ACCENT, fg=TEXT_WHITE,
                 relief="flat", cursor="hand2", padx=8, pady=2).pack(
                     side="right", padx=16, pady=6)
        self.frm_banner.pack_forget()

    # ── Summary label ─────────────────────────────────────────────────────────

    def _build_summary(self):
        self.lbl_summary = tk.Label(self.root, text="",
                 font=("Segoe UI", 9), fg=TEXT_DIM, bg=BG_CARD, anchor="w")
        self.lbl_summary.pack(fill="x", padx=18, pady=(4, 2))

    # ── Table ─────────────────────────────────────────────────────────────────

    def _build_table(self):
        tbl = tk.Frame(self.root, bg=BG)
        tbl.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.tree = ttk.Treeview(tbl, columns=COLUMNS, show="headings",
                                 style="PT.Treeview", selectmode="browse")

        for c in COLUMNS:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=COL_WIDTHS.get(c, 100),
                             anchor="center" if c in ("Date", "Created On", "Hours") else "w",
                             minwidth=50)

        self.tree.tag_configure("odd",   background=ROW_ODD)
        self.tree.tag_configure("even",  background=ROW_EVEN)
        self.tree.tag_configure("total", background="#dcfce7",
                                foreground=SUCCESS,
                                font=("Segoe UI", 9, "bold"))

        vsb = ttk.Scrollbar(tbl, orient="vertical",   command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl.rowconfigure(0, weight=1)
        tbl.columnconfigure(0, weight=1)

    # ── Footer ────────────────────────────────────────────────────────────────

    def _build_footer(self):
        tk.Frame(self.root, bg=BORDER, height=1).pack(fill="x")
        bar = tk.Frame(self.root, bg=BG_CARD, height=26)
        bar.pack(fill="x", side="bottom")
        bar.pack_propagate(False)
        self.lbl_footer = tk.Label(bar, text="",
                 font=("Segoe UI", 8), fg=TEXT_DIM, bg=BG_CARD, anchor="w")
        self.lbl_footer.pack(side="left", padx=16, fill="y")