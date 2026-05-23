from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from .config import get_column_aliases, save_column_aliases
    from .merger import ExcelMerger
except ImportError:
    from config import get_column_aliases, save_column_aliases
    from merger import ExcelMerger


OUTPUT_COLUMNS = ["SN", "Truck", "Trailer", "Position", "Status", "Type", "Return", "DSJ"]

# ── Palette ────────────────────────────────────────────────────────────────────
C = {
    # Backgrounds
    "bg":           "#0f1117",   # root / sidebar background
    "surface":      "#181c27",   # card surface
    "surface_hi":   "#1e2335",   # elevated surface / header bands
    "surface_lo":   "#13161f",   # sunken input fields
    "border":       "#2a2f47",   # subtle borders
    "border_bright":"#3d4468",   # focused / active borders

    # Brand
    "primary":      "#3b82f6",   # blue accent
    "primary_dk":   "#2563eb",   # hover / pressed
    "primary_glow": "#1d3b6b",   # soft glow bg
    "success":      "#22c55e",   # operational green
    "warn":         "#f59e0b",   # warning amber
    "error":        "#ef4444",   # error red

    # Text
    "text":         "#e8ecf5",   # primary text
    "text_sub":     "#8b91aa",   # secondary / muted
    "text_dim":     "#555b77",   # disabled / very muted

    # Log console
    "log_bg":       "#0a0c12",
    "log_text":     "#8fb4e8",
    "log_ok":       "#22c55e",
    "log_warn":     "#f59e0b",
    "log_err":      "#ef4444",

    # Sidebar
    "sidebar":      "#0d1018",
    "sidebar_hi":   "#181c27",
}

FONT_H1    = ("Segoe UI", 18, "bold")
FONT_H2    = ("Segoe UI", 10, "bold")
FONT_BODY  = ("Segoe UI", 9)
FONT_SMALL = ("Segoe UI", 8)
FONT_MONO  = ("Consolas", 9)
FONT_BADGE = ("Segoe UI", 8, "bold")


class DataDockApp:
    SIDEBAR_W = 230   # fixed pixel width for the left panel

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DataDock")
        self.root.minsize(1000, 660)
        self.root.configure(bg=C["bg"])

        # Maximise without flickering geometry jitter
        try:
            self.root.state("zoomed")
        except tk.TclError:
            self.root.attributes("-zoomed", True)

        # ── State ────────────────────────────────────────────────────────────
        self.file_paths: List[Path] = []
        self.output_path: Optional[Path] = None
        self.tables: Dict[str, pd.DataFrame] = {}
        self.table_display_map: Dict[str, str] = {}
        self.aliases_map = get_column_aliases()

        self.truck_search_var   = tk.StringVar()
        self.filter_vars:    Dict[str, tk.StringVar]  = {}
        self.filter_widgets: Dict[str, ttk.Combobox]  = {}
        self.report_vars: Dict[str, tk.BooleanVar] = {}
        self.report_summary_var = tk.StringVar(value="All reports")
        self.report_popup: Optional[tk.Toplevel] = None
        self._report_hover_job: Optional[str] = None
        self.output_path_var = tk.StringVar(value="No output selected")
        self.status_var      = tk.StringVar(value="Ready")
        self.file_count_var  = tk.StringVar(value="0 files")
        self.row_count_var   = tk.StringVar(value="")
        self.log_buffer: List[Tuple[str, str]] = []
        self.log_window: Optional[tk.Toplevel] = None
        self.log_text: Optional[tk.Text] = None

        self._configure_style()
        self._build_ui()
        self._bind_events()

    # ── ttk Style ─────────────────────────────────────────────────────────────
    def _configure_style(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")

        # Primary button
        style.configure("Primary.TButton",
            background=C["primary"], foreground="#ffffff",
            padding=(14, 8), borderwidth=0, focusthickness=0,
            font=FONT_H2, relief="flat")
        style.map("Primary.TButton",
            background=[("active", C["primary_dk"]), ("pressed", C["primary_dk"])])

        # Ghost / secondary button
        style.configure("Ghost.TButton",
            background=C["surface_hi"], foreground=C["text_sub"],
            padding=(10, 6), borderwidth=1, focusthickness=0,
            font=FONT_BODY, relief="flat")
        style.map("Ghost.TButton",
            background=[("active", C["border"])],
            foreground=[("active", C["text"])])

        # Danger (link-style clear button)
        style.configure("Danger.TButton",
            background=C["surface_hi"], foreground=C["error"],
            padding=(8, 4), borderwidth=0, focusthickness=0,
            font=FONT_SMALL, relief="flat")
        style.map("Danger.TButton",
            background=[("active", C["surface"])])

        # Treeview
        style.configure("Custom.Treeview",
            background=C["surface"], fieldbackground=C["surface"],
            foreground=C["text"], rowheight=30,
            borderwidth=0, relief="flat",
            font=FONT_BODY)
        style.configure("Custom.Treeview.Heading",
            font=FONT_BADGE,
            background=C["surface_hi"], foreground=C["text_sub"],
            borderwidth=0, relief="flat", padding=(6, 6))
        style.map("Custom.Treeview",
            background=[("selected", C["primary_glow"])],
            foreground=[("selected", C["primary"])])
        style.map("Custom.Treeview.Heading",
            background=[("active", C["border"])])

        # Combobox
        style.configure("TCombobox",
            fieldbackground=C["surface_lo"], background=C["surface_lo"],
            foreground=C["text"], arrowcolor=C["text_sub"],
            bordercolor=C["border"], lightcolor=C["border"],
            darkcolor=C["border"], selectbackground=C["primary_glow"],
            selectforeground=C["text"], padding=(6, 5),
            font=FONT_BODY)
        style.map("TCombobox",
            fieldbackground=[("readonly", C["surface_lo"])],
            bordercolor=[("focus", C["primary"])])

        # Entry
        style.configure("TEntry",
            fieldbackground=C["surface_lo"], foreground=C["text"],
            insertcolor=C["text"], bordercolor=C["border"],
            lightcolor=C["border"], darkcolor=C["border"],
            padding=(8, 5), font=FONT_BODY)
        style.map("TEntry",
            bordercolor=[("focus", C["primary"])])

        # Scrollbar
        style.configure("Thin.Vertical.TScrollbar",
            background=C["surface_hi"], troughcolor=C["bg"],
            arrowcolor=C["border"], borderwidth=0, width=8)
        style.configure("Thin.Horizontal.TScrollbar",
            background=C["surface_hi"], troughcolor=C["bg"],
            arrowcolor=C["border"], borderwidth=0, width=8)

    # ── Top-level layout ──────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        self._build_header()

        body = tk.Frame(self.root, bg=C["bg"])
        body.pack(fill="both", expand=True)

        # Left: fixed-width sidebar
        self.sidebar = tk.Frame(body, bg=C["sidebar"], width=self.SIDEBAR_W)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # 1-px divider between sidebar and content
        tk.Frame(body, bg=C["border"], width=1).pack(side="left", fill="y")

        # Right: main content fills the rest
        content = tk.Frame(body, bg=C["bg"])
        content.pack(side="left", fill="both", expand=True)
        content.rowconfigure(0, weight=1)
        content.columnconfigure(0, weight=1)

        self._build_sidebar()
        self._build_viewer(content)
        self._build_footer()

    # ── Header ────────────────────────────────────────────────────────────────
    def _build_header(self) -> None:
        bar = tk.Frame(self.root, bg=C["surface"], height=52)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        # Bottom border line
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")

        # Logo
        logo_frame = tk.Frame(bar, bg=C["surface"])
        logo_frame.pack(side="left", padx=(18, 0))

        tk.Label(logo_frame, text="●", font=("Segoe UI", 14),
                 fg=C["primary"], bg=C["surface"]).pack(side="left", padx=(0, 6))
        tk.Label(logo_frame, text="DataDock", font=FONT_H1,
                 fg=C["text"], bg=C["surface"]).pack(side="left")

        # Right: status + date
        right = tk.Frame(bar, bg=C["surface"])
        right.pack(side="right", padx=18)

        ttk.Button(right, text="Process Logs", style="Ghost.TButton",
                   command=self._open_log_window).pack(side="right", padx=(0, 12))

        self._kv_badge(right, "System Status", "Operational", C["success"]).pack(
            side="right", padx=(12, 0))
        tk.Frame(right, bg=C["border"], width=1).pack(side="right", fill="y", pady=8)
        self._kv_badge(right, "Current Date",
                       date.today().strftime("%B %d, %Y"), C["text"]).pack(
            side="right", padx=(0, 12))

    def _kv_badge(self, parent, label: str, value: str, value_color: str) -> tk.Frame:
        f = tk.Frame(parent, bg=C["surface"])
        tk.Label(f, text=label, font=FONT_SMALL,
                 fg=C["text_sub"], bg=C["surface"]).pack(anchor="e")
        tk.Label(f, text=value, font=("Segoe UI", 10, "bold"),
                 fg=value_color, bg=C["surface"]).pack(anchor="e")
        return f

    # ── Sidebar ───────────────────────────────────────────────────────────────
    def _build_sidebar(self) -> None:
        sb = self.sidebar
        self._section_label(sb, "SOURCE FILES")
        self._build_files_section(sb)
        self._hsep(sb)
        self._section_label(sb, "SCHEMA CONFIG")
        self._build_config_section(sb)
        self._hsep(sb)
        self._section_label(sb, "GENERATION")
        self._build_output_section(sb)

    def _section_label(self, parent, text: str) -> None:
        tk.Label(parent, text=text, font=("Segoe UI", 7, "bold"),
                 fg=C["text_dim"], bg=C["sidebar"],
                 padx=14).pack(fill="x", anchor="w", pady=(10, 4))

    def _hsep(self, parent) -> None:
        tk.Frame(parent, bg=C["border"], height=1).pack(fill="x", padx=14, pady=6)

    def _build_files_section(self, parent: tk.Widget) -> None:
        frame = tk.Frame(parent, bg=C["sidebar"])
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 4))

        ttk.Button(frame, text="＋  Select Excel Files",
                   style="Primary.TButton",
                   command=self._select_files).pack(fill="x", pady=(0, 6))

        # File list
        list_outer = tk.Frame(frame, bg=C["surface_lo"],
                              highlightbackground=C["border"],
                              highlightthickness=1)
        list_outer.pack(fill="both", expand=True)
        list_outer.rowconfigure(0, weight=1)
        list_outer.columnconfigure(0, weight=1)

        self.file_listbox = tk.Listbox(
            list_outer, bg=C["surface_lo"], fg=C["text_sub"],
            selectbackground=C["primary_glow"], selectforeground=C["primary"],
            activestyle="none", borderwidth=0, highlightthickness=0,
            font=FONT_SMALL, height=7, relief="flat")
        self.file_listbox.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)

        sb = ttk.Scrollbar(list_outer, orient="vertical",
                           command=self.file_listbox.yview,
                           style="Thin.Vertical.TScrollbar")
        sb.grid(row=0, column=1, sticky="ns")
        self.file_listbox.configure(yscrollcommand=sb.set)

        # Bottom actions
        btns = tk.Frame(frame, bg=C["sidebar"])
        btns.pack(fill="x", pady=(4, 0))
        ttk.Button(btns, text="Remove selected", style="Ghost.TButton",
                   command=self._remove_selected_file).pack(side="left", fill="x", expand=True)
        ttk.Button(btns, text="✕ Clear", style="Danger.TButton",
                   command=self._clear_files).pack(side="right")

    def _build_config_section(self, parent: tk.Widget) -> None:
        frame = tk.Frame(parent, bg=C["sidebar"])
        frame.pack(fill="x", padx=10, pady=(0, 4))

        ttk.Button(frame, text="Configure schema", style="Ghost.TButton",
                   command=self._open_config_editor).pack(fill="x", pady=(0, 8))

        self.config_grid = tk.Frame(frame, bg=C["sidebar"])
        self.config_grid.pack(fill="x")
        self.config_grid.columnconfigure(0, weight=1)
        self.config_grid.columnconfigure(1, weight=1)
        self._render_config_chips()

    def _build_output_section(self, parent: tk.Widget) -> None:
        frame = tk.Frame(parent, bg=C["sidebar"])
        frame.pack(fill="x", padx=10, pady=(0, 10))

        # Output path display
        path_box = tk.Frame(frame, bg=C["surface_lo"],
                            highlightbackground=C["border"],
                            highlightthickness=1)
        path_box.pack(fill="x", pady=(0, 6))

        tk.Label(path_box, text="Output Path", font=FONT_BADGE,
                 fg=C["text_dim"], bg=C["surface_lo"],
                 padx=8).pack(anchor="w", pady=(6, 0))
        self._out_label = tk.Label(path_box, textvariable=self.output_path_var,
                 font=FONT_SMALL, fg=C["primary"], bg=C["surface_lo"],
                 padx=8, wraplength=self.SIDEBAR_W - 40, justify="left")
        self._out_label.pack(anchor="w", pady=(0, 6))

        ttk.Button(frame, text="Save As…", style="Ghost.TButton",
                   command=self._select_output).pack(fill="x", pady=(0, 6))
        ttk.Button(frame, text="⚡  Generate Report",
                   style="Primary.TButton",
                   command=self._generate_report).pack(fill="x")

    # ── Output Viewer ─────────────────────────────────────────────────────────
    def _build_viewer(self, parent: tk.Widget) -> None:
        outer = tk.Frame(parent, bg=C["surface"], height=100,
                 highlightbackground=C["border"],
                 highlightthickness=1)
        outer.grid_propagate(False)
        outer.grid(row=0, column=0, sticky="nsew", padx=12, pady=(10, 4))
        outer.rowconfigure(4, weight=1)
        outer.columnconfigure(0, weight=1)

        # ── Viewer top bar ────────────────────────────────────────────────
        top_bar = tk.Frame(outer, bg=C["surface_hi"])
        top_bar.grid(row=0, column=0, sticky="ew")

        tk.Label(top_bar, text="Output Viewer", font=FONT_H2,
                 fg=C["text"], bg=C["surface_hi"],
                 padx=14, pady=10).pack(side="left")

        # Row count badge
        self.row_badge = tk.Label(top_bar, textvariable=self.row_count_var,
                 font=FONT_BADGE, fg=C["text_sub"], bg=C["surface_hi"],
                 padx=10)
        self.row_badge.pack(side="right")

        tk.Frame(outer, bg=C["border"], height=1).grid(row=1, column=0, sticky="ew")

        # ── Filter row ─────────────────────────────────────────────────────
        filter_bar = tk.Frame(outer, bg=C["surface"])
        filter_bar.grid(row=2, column=0, sticky="ew", padx=12, pady=8)

        # Report selector
        report_wrap = tk.Frame(filter_bar, bg=C["surface"])
        report_wrap.pack(side="left", padx=(0, 10))
        tk.Label(report_wrap, text="REPORTS", font=("Segoe UI", 7, "bold"),
                 fg=C["text_dim"], bg=C["surface"]).pack(anchor="w")

        self.report_select = tk.Frame(report_wrap, bg=C["surface_lo"],
                          highlightbackground=C["border"],
                          highlightthickness=1, cursor="hand2")
        self.report_select.pack(anchor="w")

        self.report_label = tk.Label(self.report_select,
                         textvariable=self.report_summary_var,
                         fg=C["text"], bg=C["surface_lo"],
                         font=FONT_BODY, padx=8, pady=5,
                         cursor="hand2")
        self.report_label.pack(side="left")
        self.report_caret = tk.Label(self.report_select, text="▾",
                         fg=C["text_dim"], bg=C["surface_lo"],
                         padx=6, cursor="hand2")
        self.report_caret.pack(side="right")
        self._bind_report_select_events()

        # Truck search
        search_wrap = tk.Frame(filter_bar, bg=C["surface"])
        search_wrap.pack(side="right")
        tk.Label(search_wrap, text="TRUCK / TRAILER", font=("Segoe UI", 7, "bold"),
                 fg=C["text_dim"], bg=C["surface"]).pack(anchor="w")
        search_frame = tk.Frame(search_wrap, bg=C["surface_lo"],
                                highlightbackground=C["border"],
                                highlightthickness=1)
        search_frame.pack(anchor="e")
        tk.Label(search_frame, text="⌕", font=("Segoe UI", 11),
                 fg=C["text_dim"], bg=C["surface_lo"], padx=6).pack(side="left")
        ttk.Entry(search_frame, textvariable=self.truck_search_var,
                  width=22, style="TEntry").pack(side="left")

        for col in ["Truck", "Trailer", "Position", "Status", "Type", "Return"]:
            var = tk.StringVar(value="All")
            self.filter_vars[col] = var

            col_wrap = tk.Frame(filter_bar, bg=C["surface"])
            col_wrap.pack(side="left", padx=(0, 10))

            tk.Label(col_wrap, text=col.upper(), font=("Segoe UI", 7, "bold"),
                     fg=C["text_dim"], bg=C["surface"]).pack(anchor="w")
            combo = ttk.Combobox(col_wrap, textvariable=var,
                                 state="readonly", values=["All"], width=11)
            combo.pack()
            self.filter_widgets[col] = combo
            combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_view())

        tk.Frame(outer, bg=C["border"], height=1).grid(row=3, column=0, sticky="ew")

        # ── Treeview ───────────────────────────────────────────────────────
        tree_frame = tk.Frame(outer, bg=C["surface"])
        tree_frame.grid(row=4, column=0, sticky="nsew", padx=0, pady=0)
        outer.rowconfigure(4, weight=1)
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            tree_frame, columns=OUTPUT_COLUMNS,
            show="headings", style="Custom.Treeview",
            selectmode="browse")

        for col in OUTPUT_COLUMNS:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=100, minwidth=60, stretch=True, anchor="w")
        # Narrow fixed columns
        self.tree.column("SN",     width=50,  minwidth=40,  stretch=False)
        self.tree.column("DSJ",    width=70,  minwidth=50,  stretch=False)
        self.tree.column("Type",   width=90,  minwidth=70,  stretch=False)
        self.tree.column("Return", width=80,  minwidth=60,  stretch=False)

        self.tree.grid(row=0, column=0, sticky="nsew")

        # Alternating row colours
        self.tree.tag_configure("even", background=C["surface"])
        self.tree.tag_configure("odd",  background=C["surface_hi"])
        self.tree.tag_configure("group_header",
                    background=C["surface_lo"],
                    foreground=C["text"],
                    font=FONT_BADGE)
        self.tree.tag_configure("spacer",
                    background=C["surface"],
                    foreground=C["surface"])

        y_scroll = ttk.Scrollbar(tree_frame, orient="vertical",
                                 command=self.tree.yview,
                                 style="Thin.Vertical.TScrollbar")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ttk.Scrollbar(tree_frame, orient="horizontal",
                                 command=self.tree.xview,
                                 style="Thin.Horizontal.TScrollbar")
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=y_scroll.set,
                            xscrollcommand=x_scroll.set)

    # ── Log window ───────────────────────────────────────────────────────────
    def _open_log_window(self) -> None:
        if self.log_window and self.log_window.winfo_exists():
            self.log_window.lift()
            return

        win = tk.Toplevel(self.root)
        win.title("Process Logs")
        win.configure(bg=C["surface"])
        win.minsize(600, 300)

        outer = tk.Frame(win, bg=C["surface"],
                         highlightbackground=C["border"],
                         highlightthickness=1)
        outer.pack(fill="both", expand=True, padx=12, pady=12)
        outer.rowconfigure(1, weight=1)
        outer.columnconfigure(0, weight=1)

        header = tk.Frame(outer, bg=C["surface_hi"])
        header.grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="Process Logs", font=FONT_H2,
                 fg=C["text"], bg=C["surface_hi"],
                 padx=14, pady=8).pack(side="left")

        ttk.Button(header, text="Clear", style="Ghost.TButton",
                   command=self._clear_log).pack(side="right", padx=10, pady=6)

        self.log_text = tk.Text(
            outer, bg=C["log_bg"], fg=C["log_text"],
            borderwidth=0, highlightthickness=0,
            font=FONT_MONO, wrap="word", state="disabled",
            insertbackground=C["text"], padx=10, pady=8)
        self.log_text.grid(row=1, column=0, sticky="nsew")

        log_scroll = ttk.Scrollbar(outer, orient="vertical",
                                   command=self.log_text.yview,
                                   style="Thin.Vertical.TScrollbar")
        log_scroll.grid(row=1, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=log_scroll.set)

        # Tag colours for log levels
        self.log_text.tag_configure("INFO",    foreground=C["log_text"])
        self.log_text.tag_configure("SUCCESS", foreground=C["log_ok"])
        self.log_text.tag_configure("WARN",    foreground=C["log_warn"])
        self.log_text.tag_configure("ERROR",   foreground=C["log_err"])
        self.log_text.tag_configure("ts",      foreground=C["text_dim"])
        self.log_text.tag_configure("bracket", foreground=C["text_dim"])

        for level, message in self.log_buffer:
            self._append_log(message, level)

        win.protocol("WM_DELETE_WINDOW", self._close_log_window)
        self.log_window = win

    def _close_log_window(self) -> None:
        if self.log_window and self.log_window.winfo_exists():
            self.log_window.destroy()
        self.log_window = None
        self.log_text = None

    # ── Footer ────────────────────────────────────────────────────────────────
    def _build_footer(self) -> None:
        tk.Frame(self.root, bg=C["border"], height=1).pack(fill="x")
        bar = tk.Frame(self.root, bg=C["surface_hi"], height=30)
        bar.pack(fill="x")
        bar.pack_propagate(False)

        left = tk.Frame(bar, bg=C["surface_hi"])
        left.pack(side="left", padx=16)

        self._status_dot = tk.Label(left, text="●", font=("Segoe UI", 8),
                                    fg=C["success"], bg=C["surface_hi"])
        self._status_dot.pack(side="left", padx=(0, 6))

        tk.Label(left, textvariable=self.status_var, font=FONT_SMALL,
                 fg=C["text_sub"], bg=C["surface_hi"]).pack(side="left")

        tk.Label(left, textvariable=self.file_count_var, font=FONT_SMALL,
                 fg=C["text_dim"], bg=C["surface_hi"], padx=14).pack(side="left")

        right = tk.Frame(bar, bg=C["surface_hi"])
        right.pack(side="right", padx=16)
        tk.Label(right, text="v1.0.0", font=FONT_SMALL,
                 fg=C["text_dim"], bg=C["surface_hi"]).pack()

    # ── Events ────────────────────────────────────────────────────────────────
    def _bind_events(self) -> None:
        self.truck_search_var.trace_add("write", lambda *_: self._refresh_view())
        

    # ── Logging ───────────────────────────────────────────────────────────────
    def _log(self, message: str, level: str = "INFO") -> None:
        self.log_buffer.append((level, message))
        if self.log_text is None:
            return
        self._append_log(message, level)

    def _append_log(self, message: str, level: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{ts}] ", "ts")
        self.log_text.insert("end", f"[{level}] ", level)
        self.log_text.insert("end", f"{message}\n", level)
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    def _clear_log(self) -> None:
        self.log_buffer.clear()
        if self.log_text is None:
            return
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # ── File actions ──────────────────────────────────────────────────────────
    def _select_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select Excel files",
            filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if not paths:
            return
        added = 0
        for p in paths:
            cand = Path(p)
            if cand not in self.file_paths:
                self.file_paths.append(cand)
                added += 1
        self._refresh_file_list()
        self._log(f"Added {added} file(s). Total: {len(self.file_paths)}.")

    def _clear_files(self) -> None:
        self.file_paths.clear()
        self._refresh_file_list()
        self._log("File list cleared.")

    def _remove_selected_file(self) -> None:
        for idx in reversed(list(self.file_listbox.curselection())):
            removed = self.file_paths.pop(idx)
            self._log(f"Removed: {removed.name}")
        self._refresh_file_list()

    def _refresh_file_list(self) -> None:
        self.file_listbox.delete(0, "end")
        for p in self.file_paths:
            self.file_listbox.insert("end", f"  {p.name}")
        n = len(self.file_paths)
        self.file_count_var.set(f"{n} file{'s' if n != 1 else ''} selected")

    # ── Config ────────────────────────────────────────────────────────────────
    def _render_config_chips(self) -> None:
        for child in self.config_grid.winfo_children():
            child.destroy()

        columns = list(self.aliases_map.keys())
        for i, col in enumerate(columns):
            chip = tk.Frame(self.config_grid, bg=C["primary_glow"],
                            highlightbackground=C["primary"],
                            highlightthickness=1)
            chip.grid(row=i // 2, column=i % 2,
                      padx=3, pady=3, sticky="ew")
            tk.Label(chip, text=col, font=FONT_BADGE,
                     fg=C["primary"], bg=C["primary_glow"],
                     padx=6, pady=5).pack()

    def _open_config_editor(self) -> None:
        editor = tk.Toplevel(self.root)
        editor.title("Schema Config")
        editor.configure(bg=C["surface"])
        editor.minsize(560, 420)
        editor.transient(self.root)
        editor.grab_set()

        working = {k: list(v) for k, v in self.aliases_map.items()}
        selected_target = tk.StringVar()

        def refresh_targets() -> None:
            target_list.delete(0, "end")
            for name in working.keys():
                target_list.insert("end", name)

            if working:
                if selected_target.get() not in working:
                    selected_target.set(next(iter(working)))
            else:
                selected_target.set("")

            target_list.selection_clear(0, "end")
            if selected_target.get() in working:
                idx = list(working.keys()).index(selected_target.get())
                target_list.selection_set(idx)
                target_list.see(idx)

            refresh_aliases()

        def refresh_aliases() -> None:
            alias_list.delete(0, "end")
            target = selected_target.get()
            if not target or target not in working:
                return
            for alias in working[target]:
                alias_list.insert("end", alias)

        def on_target_select(_event=None) -> None:
            selection = target_list.curselection()
            if selection:
                selected_target.set(target_list.get(selection[0]))
                refresh_aliases()

        def add_target() -> None:
            name = target_entry.get().strip()
            if not name:
                return
            if name in working:
                messagebox.showwarning("Schema Config", "Target already exists.")
                return
            working[name] = []
            target_entry.delete(0, "end")
            selected_target.set(name)
            refresh_targets()

        def remove_target() -> None:
            target = selected_target.get()
            if not target:
                return
            if not messagebox.askyesno(
                "Schema Config",
                f"Remove target '{target}'?",
            ):
                return
            working.pop(target, None)
            selected_target.set("")
            refresh_targets()

        def add_alias() -> None:
            target = selected_target.get()
            if not target:
                messagebox.showwarning(
                    "Schema Config", "Select a target column first."
                )
                return
            alias = alias_entry.get().strip()
            if not alias:
                return
            if alias in working[target]:
                return
            working[target].append(alias)
            alias_entry.delete(0, "end")
            refresh_aliases()

        def remove_alias() -> None:
            target = selected_target.get()
            if not target:
                return
            selection = alias_list.curselection()
            if not selection:
                return
            for idx in reversed(selection):
                alias = alias_list.get(idx)
                if alias in working[target]:
                    working[target].remove(alias)
            refresh_aliases()

        def normalize_map() -> Dict[str, List[str]]:
            cleaned: Dict[str, List[str]] = {}
            for key, aliases in working.items():
                name = key.strip()
                if not name:
                    continue
                seen = set()
                clean_aliases: List[str] = []
                for alias in aliases:
                    alias = alias.strip()
                    if not alias or alias in seen:
                        continue
                    seen.add(alias)
                    clean_aliases.append(alias)
                cleaned[name] = clean_aliases
            return cleaned

        def save_changes() -> None:
            cleaned = normalize_map()
            if not cleaned:
                messagebox.showerror(
                    "Schema Config", "At least one target column is required."
                )
                return
            try:
                save_column_aliases(cleaned)
            except Exception as exc:
                messagebox.showerror(
                    "Schema Config", f"Failed to save config:\n{exc}"
                )
                return

            self.aliases_map = cleaned
            self._render_config_chips()
            self._log("Schema config updated.")
            editor.destroy()

        header = tk.Frame(editor, bg=C["surface_hi"])
        header.pack(fill="x")
        tk.Label(header, text="Schema Config", font=FONT_H2,
                 fg=C["text"], bg=C["surface_hi"],
                 padx=14).pack(anchor="w", pady=8)

        body = tk.Frame(editor, bg=C["surface"])
        body.pack(fill="both", expand=True, padx=14, pady=12)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(1, weight=1)

        left = tk.Frame(body, bg=C["surface"])
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        tk.Label(left, text="TARGET COLUMNS", font=("Segoe UI", 7, "bold"),
                 fg=C["text_dim"], bg=C["surface"]).pack(anchor="w")

        target_box = tk.Frame(left, bg=C["surface_lo"],
                              highlightbackground=C["border"],
                              highlightthickness=1)
        target_box.pack(fill="both", expand=True, pady=(6, 8))
        target_box.rowconfigure(0, weight=1)
        target_box.columnconfigure(0, weight=1)
        target_list = tk.Listbox(
            target_box, bg=C["surface_lo"], fg=C["text"],
            selectbackground=C["primary_glow"], selectforeground=C["primary"],
            activestyle="none", borderwidth=0, highlightthickness=0,
            font=FONT_BODY)
        target_list.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        target_scroll = ttk.Scrollbar(
            target_box, orient="vertical", command=target_list.yview,
            style="Thin.Vertical.TScrollbar")
        target_scroll.grid(row=0, column=1, sticky="ns")
        target_list.configure(yscrollcommand=target_scroll.set)
        target_list.bind("<<ListboxSelect>>", on_target_select)

        target_entry = ttk.Entry(left, style="TEntry")
        target_entry.pack(fill="x", pady=(0, 6))
        target_actions = tk.Frame(left, bg=C["surface"])
        target_actions.pack(fill="x")
        ttk.Button(target_actions, text="Add target", style="Ghost.TButton",
                   command=add_target).pack(side="left", fill="x", expand=True)
        ttk.Button(target_actions, text="Remove", style="Danger.TButton",
                   command=remove_target).pack(side="right")

        right = tk.Frame(body, bg=C["surface"])
        right.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        tk.Label(right, text="ALIASES", font=("Segoe UI", 7, "bold"),
                 fg=C["text_dim"], bg=C["surface"]).pack(anchor="w")

        alias_box = tk.Frame(right, bg=C["surface_lo"],
                             highlightbackground=C["border"],
                             highlightthickness=1)
        alias_box.pack(fill="both", expand=True, pady=(6, 8))
        alias_box.rowconfigure(0, weight=1)
        alias_box.columnconfigure(0, weight=1)
        alias_list = tk.Listbox(
            alias_box, bg=C["surface_lo"], fg=C["text"],
            selectbackground=C["primary_glow"], selectforeground=C["primary"],
            activestyle="none", borderwidth=0, highlightthickness=0,
            font=FONT_BODY)
        alias_list.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        alias_scroll = ttk.Scrollbar(
            alias_box, orient="vertical", command=alias_list.yview,
            style="Thin.Vertical.TScrollbar")
        alias_scroll.grid(row=0, column=1, sticky="ns")
        alias_list.configure(yscrollcommand=alias_scroll.set)

        alias_entry = ttk.Entry(right, style="TEntry")
        alias_entry.pack(fill="x", pady=(0, 6))
        alias_actions = tk.Frame(right, bg=C["surface"])
        alias_actions.pack(fill="x")
        ttk.Button(alias_actions, text="Add alias", style="Ghost.TButton",
                   command=add_alias).pack(side="left", fill="x", expand=True)
        ttk.Button(alias_actions, text="Remove", style="Danger.TButton",
                   command=remove_alias).pack(side="right")

        footer = tk.Frame(editor, bg=C["surface_hi"])
        footer.pack(fill="x")
        ttk.Button(footer, text="Cancel", style="Ghost.TButton",
                   command=editor.destroy).pack(side="right", padx=(0, 10), pady=8)
        ttk.Button(footer, text="Save changes", style="Primary.TButton",
                   command=save_changes).pack(side="right", padx=10, pady=8)

        refresh_targets()

    # ── Output path ───────────────────────────────────────────────────────────
    def _select_output(self) -> None:
        default = f"IMPORT_REPORT_{date.today().isoformat()}.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save merged report",
            defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        self.output_path = Path(path)
        # Truncate for display
        display = str(self.output_path)
        if len(display) > 34:
            display = "…" + display[-34:]
        self.output_path_var.set(display)
        self._log(f"Output: {self.output_path.name}")

    # ── Generate ──────────────────────────────────────────────────────────────
    def _generate_report(self) -> None:
        if not self.file_paths:
            messagebox.showwarning("DataDock", "Select Excel files first.")
            return

        if self.output_path is None:
            self._select_output()
            if self.output_path is None:
                return

        self._log("Merging files…")
        self.status_var.set("Processing…")
        self._status_dot.configure(fg=C["warn"])
        self.root.update_idletasks()

        merger = ExcelMerger()
        merger.set_files(self.file_paths)
        tables = merger.merge_files(aliases_map=self.aliases_map)
        tables = merger.format_tables(tables)

        for msg in merger.errors:
            self._log(msg, "WARN")

        if not tables:
            self._log("No tables generated.", "ERROR")
            self.status_var.set("Failed")
            self._status_dot.configure(fg=C["error"])
            messagebox.showerror("DataDock", "No tables were generated.")
            return

        self.tables = tables
        self._update_table_selector()
        self._refresh_view()

        if self.output_path.exists():
            if not messagebox.askyesno("Overwrite?",
                    "File already exists. Replace it?"):
                return
            try:
                self.output_path.unlink()
            except PermissionError:
                messagebox.showerror("File in use",
                    "Close the Excel file and try again.")
                return

        try:
            merger.export_excel_tables(tables, self.output_path)
        except Exception as exc:
            self._log(str(exc), "ERROR")
            self.status_var.set("Export failed")
            self._status_dot.configure(fg=C["error"])
            messagebox.showerror("DataDock", f"Export failed:\n{exc}")
            return

        total_rows = sum(len(t) for t in tables.values())
        self._log(
            f"Done — {len(tables)} table(s), {total_rows} rows → {self.output_path.name}",
            "SUCCESS")
        self.status_var.set("Report generated")
        self._status_dot.configure(fg=C["success"])

    # ── Table selector ────────────────────────────────────────────────────────
    def _update_table_selector(self) -> None:
        self.table_display_map = {}
        for name in self.tables:
            display = Path(name).stem
            self.table_display_map[display] = name

        self.report_vars = {}
        for display in self.table_display_map.keys():
            self.report_vars[display] = tk.BooleanVar(value=True)

        self._update_report_summary()
        self._refresh_view()

    def _update_report_summary(self) -> None:
        if not self.table_display_map:
            self.report_summary_var.set("No reports")
            return

        selected = [
            name
            for name, var in self.report_vars.items()
            if var.get()
        ]
        if not selected or len(selected) == len(self.report_vars):
            self.report_summary_var.set("All reports")
            return
        if len(selected) == 1:
            self.report_summary_var.set(selected[0])
            return
        self.report_summary_var.set(f"{len(selected)} selected")

    def _bind_report_select_events(self) -> None:
        widgets = [self.report_select, self.report_label, self.report_caret]
        for widget in widgets:
            widget.bind("<Button-1>", self._toggle_report_popup)
            widget.bind("<Enter>", self._schedule_report_popup)
            widget.bind("<Leave>", self._cancel_report_popup)

    def _schedule_report_popup(self, _event=None) -> None:
        if self.report_popup and self.report_popup.winfo_exists():
            return
        if self._report_hover_job:
            self.root.after_cancel(self._report_hover_job)
        self._report_hover_job = self.root.after(200, self._open_report_popup)

    def _cancel_report_popup(self, _event=None) -> None:
        if self.report_select is not None:
            widget = self.report_select.winfo_containing(
                self.root.winfo_pointerx(),
                self.root.winfo_pointery(),
            )
            if widget and (widget == self.report_select or widget.master == self.report_select):
                return
        if self._report_hover_job:
            self.root.after_cancel(self._report_hover_job)
            self._report_hover_job = None

    def _toggle_report_popup(self, _event=None) -> None:
        if self.report_popup and self.report_popup.winfo_exists():
            self._close_report_popup()
        else:
            self._open_report_popup()

    def _open_report_popup(self) -> None:
        self._cancel_report_popup()
        if not self.table_display_map:
            return
        if self.report_popup and self.report_popup.winfo_exists():
            return

        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.configure(bg=C["surface"])
        popup.attributes("-topmost", True)

        x = self.report_select.winfo_rootx()
        y = self.report_select.winfo_rooty() + self.report_select.winfo_height()
        width = max(220, self.report_select.winfo_width() + 24)
        height = 170
        popup.geometry(f"{width}x{height}+{x}+{y}")

        outer = tk.Frame(popup, bg=C["surface"],
                         highlightbackground=C["border"],
                         highlightthickness=1)
        outer.pack(fill="both", expand=True)
        outer.rowconfigure(0, weight=1)
        outer.columnconfigure(0, weight=1)

        canvas = tk.Canvas(outer, bg=C["surface"], highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview,
                               style="Thin.Vertical.TScrollbar")
        scroll.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scroll.set)

        frame = tk.Frame(canvas, bg=C["surface"])
        canvas.create_window((0, 0), window=frame, anchor="nw")

        def update_scroll(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        frame.bind("<Configure>", update_scroll)

        for display, var in self.report_vars.items():
            tk.Checkbutton(
                frame,
                text=display,
                variable=var,
                command=self._on_report_toggle,
                fg=C["text_sub"],
                bg=C["surface"],
                activebackground=C["surface"],
                activeforeground=C["text"],
                selectcolor=C["surface"],
                font=FONT_BODY,
                anchor="w",
            ).pack(fill="x", padx=8, pady=3)

        popup.bind("<FocusOut>", lambda _e: self._close_report_popup())
        popup.bind("<Escape>", lambda _e: self._close_report_popup())
        popup.focus_set()
        self.report_popup = popup

    def _close_report_popup(self) -> None:
        if self.report_popup and self.report_popup.winfo_exists():
            self.report_popup.destroy()
        self.report_popup = None

    def _on_report_toggle(self) -> None:
        self._update_report_summary()
        self._refresh_view()

    def _get_selected_table(self) -> pd.DataFrame:
        if not self.table_display_map:
            return pd.DataFrame()

        selected = [
            key
            for display, key in self.table_display_map.items()
            if self.report_vars.get(display) and self.report_vars[display].get()
        ]
        if not selected:
            selected = list(self.table_display_map.values())

        if len(selected) == 1:
            return self.tables.get(selected[0], pd.DataFrame())

        parts = [
            self.tables[name].assign(Source=Path(name).stem)
            for name in selected
            if name in self.tables
        ]
        if not parts:
            return pd.DataFrame()
        df = pd.concat(parts, ignore_index=True)
        cols = ["Source"] + OUTPUT_COLUMNS
        return df.loc[:, [c for c in cols if c in df.columns]]

    # ── Filters ───────────────────────────────────────────────────────────────
    def _refresh_filters(self, df: pd.DataFrame) -> None:
        for col, var in self.filter_vars.items():
            values = ["All"]
            if col in df.columns:
                values += sorted({str(v) for v in df[col].dropna()})
            widget = self.filter_widgets.get(col)
            if widget:
                widget.configure(values=values)
            if var.get() not in values:
                var.set("All")

    def _apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
        out = df
        for col, var in self.filter_vars.items():
            v = var.get()
            if v != "All" and col in out.columns:
                out = out[out[col].astype(str) == v]
        q = self.truck_search_var.get().strip().lower()
        if q:
            cols = [c for c in ["Truck", "Trailer"] if c in out.columns]
            if cols:
                mask = out[cols].astype(str).apply(
                    lambda c: c.str.lower().str.contains(q, na=False))
                out = out[mask.any(axis=1)]
        return out

    def _refresh_view(self) -> None:
        df = self._get_selected_table()
        if df.empty:
            self._update_treeview(pd.DataFrame())
            self.row_count_var.set("")
            return
        self._refresh_filters(df)
        filtered = self._apply_filters(df)
        self._update_treeview(filtered)
        n = len(filtered)
        self.row_count_var.set(f"{n:,} row{'s' if n != 1 else ''}")

    def _update_treeview(self, df: pd.DataFrame) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

        display_cols = list(df.columns) if not df.empty else OUTPUT_COLUMNS
        if "Source" in display_cols:
            display_cols = [c for c in display_cols if c != "Source"]
        if not display_cols:
            display_cols = OUTPUT_COLUMNS

        self.tree.configure(columns=display_cols)
        for col in display_cols:
            self.tree.heading(col, text=col.upper())
            self.tree.column(col, width=100, minwidth=60, stretch=True, anchor="w")
        for narrow, w in [("SN", 50), ("DSJ", 70), ("Type", 90), ("Return", 80)]:
            if narrow in display_cols:
                self.tree.column(narrow, width=w, minwidth=w - 10, stretch=False)

        if df.empty:
            return

        if "Source" in df.columns:
            row_index = 0
            header_col = "Position" if "Position" in display_cols else display_cols[0]
            header_index = display_cols.index(header_col)
            current_width = self.tree.column(header_col, "width")
            if current_width < 200:
                self.tree.column(header_col, width=200, minwidth=120, stretch=True)
            first_group = True
            for source in df["Source"].dropna().astype(str).unique():
                spacer_values = [""] * len(display_cols)
                header_values = [""] * len(display_cols)
                header_values[header_index] = f"--- {source} ---"
                if not first_group:
                    self.tree.insert("", "end", values=spacer_values, tags=("spacer",))
                self.tree.insert("", "end", values=header_values, tags=("group_header",))

                subset = df[df["Source"].astype(str) == source]
                for _, row in subset.iterrows():
                    tag = "even" if row_index % 2 == 0 else "odd"
                    self.tree.insert("", "end",
                                     values=[row.get(c, "") for c in display_cols],
                                     tags=(tag,))
                    row_index += 1
                first_group = False
            return

        for i, (_, row) in enumerate(df.iterrows()):
            tag = "even" if i % 2 == 0 else "odd"
            self.tree.insert("", "end",
                             values=[row.get(c, "") for c in display_cols],
                             tags=(tag,))


def run_app() -> None:
    root = tk.Tk()
    DataDockApp(root)
    root.mainloop()