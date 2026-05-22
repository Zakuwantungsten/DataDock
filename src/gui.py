from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
	from .config import get_column_aliases
	from .merger import ExcelMerger
except ImportError:  # pragma: no cover - supports running as a script
	from config import get_column_aliases
	from merger import ExcelMerger


OUTPUT_COLUMNS = [
	"SN",
	"Truck",
	"Trailer",
	"Position",
	"Status",
	"Type",
	"Return",
	"DSJ",
]


class DataDockApp:
	def __init__(self, root: tk.Tk) -> None:
		self.root = root
		self.root.title("DataDock")
		self.root.geometry("1280x800")
		self.root.minsize(1100, 700)
		try:
			self.root.state("zoomed")
		except tk.TclError:
			self.root.attributes("-zoomed", True)

		self.colors = {
			"bg": "#f7f9fb",
			"card": "#ffffff",
			"card_alt": "#f2f4f6",
			"border": "#c3c6d7",
			"primary": "#004ac6",
			"primary_hover": "#003ea8",
			"text": "#191c1e",
			"text_muted": "#434655",
			"tertiary": "#006243",
			"error": "#ba1a1a",
			"log_bg": "#1e1e1e",
			"log_text": "#b7c8e1",
		}

		self.root.configure(bg=self.colors["bg"])
		self.file_paths: List[Path] = []
		self.output_path: Optional[Path] = None
		self.tables: Dict[str, pd.DataFrame] = {}
		self.table_display_map: Dict[str, str] = {}

		self.selected_table_var = tk.StringVar(value="All Tables")
		self.search_var = tk.StringVar()
		self.filter_vars: Dict[str, tk.StringVar] = {}
		self.filter_widgets: Dict[str, ttk.Combobox] = {}

		self.output_path_var = tk.StringVar(value="No output selected")
		self.status_var = tk.StringVar(value="Ready")
		self.file_count_var = tk.StringVar(value="0 files selected")
		self.row_count_var = tk.StringVar(value="")

		self._configure_style()
		self._build_ui()
		self._bind_events()

	def _configure_style(self) -> None:
		style = ttk.Style()
		style.theme_use("clam")

		style.configure(
			"Primary.TButton",
			background=self.colors["primary"],
			foreground="white",
			padding=(12, 8),
			borderwidth=0,
			focusthickness=0,
		)
		style.map(
			"Primary.TButton",
			background=[("active", self.colors["primary_hover"])],
		)

		style.configure(
			"Secondary.TButton",
			background=self.colors["card_alt"],
			foreground=self.colors["text"],
			padding=(10, 6),
			borderwidth=1,
		)

		style.configure(
			"Treeview",
			background=self.colors["card"],
			fieldbackground=self.colors["card"],
			foreground=self.colors["text"],
			rowheight=28,
			bordercolor=self.colors["border"],
		)
		style.configure(
			"Treeview.Heading",
			font=("Segoe UI", 9, "bold"),
			background=self.colors["card_alt"],
			foreground=self.colors["text_muted"],
		)

	def _build_ui(self) -> None:
		self._build_header()
		self._build_main()
		self._build_footer()

	def _build_header(self) -> None:
		header = tk.Frame(self.root, bg=self.colors["card"], height=56)
		header.pack(fill="x")

		left = tk.Frame(header, bg=self.colors["card"])
		left.pack(side="left", padx=16, pady=10)

		tk.Label(
			left,
			text="DataDock",
			font=("Segoe UI", 20, "bold"),
			fg=self.colors["text"],
			bg=self.colors["card"],
		).pack(side="left")

		right = tk.Frame(header, bg=self.colors["card"])
		right.pack(side="right", padx=16, pady=8)

		status_box = tk.Frame(right, bg=self.colors["card"])
		status_box.pack(side="left", padx=12)
		tk.Label(
			status_box,
			text="System Status",
			font=("Segoe UI", 8, "bold"),
			fg=self.colors["text_muted"],
			bg=self.colors["card"],
		).pack(anchor="e")
		tk.Label(
			status_box,
			text="Operational",
			font=("Segoe UI", 11, "bold"),
			fg=self.colors["tertiary"],
			bg=self.colors["card"],
		).pack(anchor="e")

		divider = tk.Frame(right, width=1, bg=self.colors["border"])
		divider.pack(side="left", fill="y", padx=12)

		date_box = tk.Frame(right, bg=self.colors["card"])
		date_box.pack(side="left")
		tk.Label(
			date_box,
			text="Current Date",
			font=("Segoe UI", 8, "bold"),
			fg=self.colors["text_muted"],
			bg=self.colors["card"],
		).pack(anchor="e")
		tk.Label(
			date_box,
			text=date.today().strftime("%B %d, %Y"),
			font=("Segoe UI", 11, "bold"),
			fg=self.colors["text"],
			bg=self.colors["card"],
		).pack(anchor="e")

	def _build_main(self) -> None:
		main = tk.Frame(self.root, bg=self.colors["bg"])
		main.pack(fill="both", expand=True, padx=0, pady=0)

		main.columnconfigure(0, weight=5)
		main.columnconfigure(1, weight=7)
		main.rowconfigure(0, weight=1)

		left = tk.Frame(main, bg=self.colors["bg"])
		right = tk.Frame(main, bg=self.colors["bg"])
		left.grid(row=0, column=0, sticky="nsew", padx=(0, 0))
		right.grid(row=0, column=1, sticky="nsew", padx=(0, 0))

		left.rowconfigure(0, weight=4)
		left.rowconfigure(1, weight=2)
		left.rowconfigure(2, weight=2)

		self._build_files_card(left)
		self._build_config_card(left)
		self._build_output_card(left)

		right.rowconfigure(0, weight=7)
		right.rowconfigure(1, weight=3)
		self._build_viewer_card(right)
		self._build_log_card(right)

	def _build_footer(self) -> None:
		footer = tk.Frame(self.root, bg=self.colors["card_alt"], height=36)
		footer.pack(fill="x")

		left = tk.Frame(footer, bg=self.colors["card_alt"])
		left.pack(side="left", padx=24)
		tk.Label(
			left,
			textvariable=self.status_var,
			font=("Segoe UI", 9, "bold"),
			fg=self.colors["text_muted"],
			bg=self.colors["card_alt"],
		).pack(side="left")
		tk.Label(
			left,
			textvariable=self.file_count_var,
			font=("Segoe UI", 9),
			fg=self.colors["text_muted"],
			bg=self.colors["card_alt"],
			padx=12,
		).pack(side="left")
		tk.Label(
			left,
			textvariable=self.row_count_var,
			font=("Segoe UI", 9),
			fg=self.colors["text_muted"],
			bg=self.colors["card_alt"],
		).pack(side="left")

		right = tk.Frame(footer, bg=self.colors["card_alt"])
		right.pack(side="right", padx=24)
		tk.Label(
			right,
			text="v1.0.0",
			font=("Segoe UI", 9, "bold"),
			fg=self.colors["text"],
			bg=self.colors["card_alt"],
		).pack(side="right")

	def _build_card(self, parent: tk.Widget, title: str) -> tk.Frame:
		card = tk.Frame(
			parent,
			bg=self.colors["card"],
			highlightbackground=self.colors["border"],
			highlightthickness=1,
		)
		header = tk.Frame(card, bg=self.colors["card_alt"])
		header.pack(fill="x")
		tk.Label(
			header,
			text=title,
			font=("Segoe UI", 11, "bold"),
			fg=self.colors["text"],
			bg=self.colors["card_alt"],
			padx=12,
			pady=8,
		).pack(side="left")
		return card

	def _build_files_card(self, parent: tk.Widget) -> None:
		card = self._build_card(parent, "Source Files")
		card.grid(row=0, column=0, sticky="nsew", pady=0)
		card.rowconfigure(1, weight=1)
		card.columnconfigure(0, weight=1)

		header_actions = tk.Frame(card, bg=self.colors["card_alt"])
		header_actions.pack(fill="x")
		tk.Button(
			header_actions,
			text="Clear list",
			command=self._clear_files,
			bg=self.colors["card_alt"],
			fg=self.colors["error"],
			bd=0,
			padx=12,
			pady=6,
		).pack(side="right", padx=12, pady=6)

		content = tk.Frame(card, bg=self.colors["card"], padx=8, pady=8)
		content.pack(fill="both", expand=True)

		select_btn = ttk.Button(
			content,
			text="Select Excel Files",
			style="Primary.TButton",
			command=self._select_files,
		)
		select_btn.pack(fill="x")

		list_frame = tk.Frame(content, bg=self.colors["card"])
		list_frame.pack(fill="both", expand=True, pady=(8, 0))
		list_frame.rowconfigure(0, weight=1)
		list_frame.columnconfigure(0, weight=1)

		self.file_listbox = tk.Listbox(
			list_frame,
			bg=self.colors["card"],
			fg=self.colors["text"],
			highlightthickness=1,
			highlightbackground=self.colors["border"],
			selectbackground=self.colors["card_alt"],
			activestyle="none",
		)
		self.file_listbox.grid(row=0, column=0, sticky="nsew")

		scrollbar = ttk.Scrollbar(
			list_frame,
			orient="vertical",
			command=self.file_listbox.yview,
		)
		scrollbar.grid(row=0, column=1, sticky="ns")
		self.file_listbox.configure(yscrollcommand=scrollbar.set)

		remove_btn = ttk.Button(
			content,
			text="Remove selected",
			style="Secondary.TButton",
			command=self._remove_selected_file,
		)
		remove_btn.pack(fill="x", pady=(8, 0))

	def _build_config_card(self, parent: tk.Widget) -> None:
		card = self._build_card(parent, "Schema Config")
		card.grid(row=1, column=0, sticky="nsew", pady=0)

		content = tk.Frame(card, bg=self.colors["card"], padx=8, pady=8)
		content.pack(fill="both", expand=True)

		ttk.Button(
			content,
			text="Open config",
			style="Secondary.TButton",
			command=self._open_config,
		).pack(anchor="e")

		columns = list(get_column_aliases().keys())
		grid = tk.Frame(content, bg=self.colors["card"])
		grid.pack(fill="x", pady=(8, 0))
		for index, column in enumerate(columns):
			label = tk.Label(
				grid,
				text=column,
				font=("Segoe UI", 9, "bold"),
				fg=self.colors["text_muted"],
				bg=self.colors["card_alt"],
				padx=8,
				pady=6,
			)
			label.grid(row=index // 2, column=index % 2, padx=4, pady=4, sticky="ew")
			grid.columnconfigure(index % 2, weight=1)

	def _build_output_card(self, parent: tk.Widget) -> None:
		card = self._build_card(parent, "Generation Parameters")
		card.grid(row=2, column=0, sticky="nsew", pady=0)

		content = tk.Frame(card, bg=self.colors["card"], padx=8, pady=8)
		content.pack(fill="both", expand=True)

		output_row = tk.Frame(content, bg=self.colors["card_alt"], padx=8, pady=8)
		output_row.pack(fill="x")
		tk.Label(
			output_row,
			text="Output Path",
			font=("Segoe UI", 8, "bold"),
			fg=self.colors["text_muted"],
			bg=self.colors["card_alt"],
		).pack(anchor="w")
		tk.Label(
			output_row,
			textvariable=self.output_path_var,
			font=("Segoe UI", 9),
			fg=self.colors["primary"],
			bg=self.colors["card_alt"],
		).pack(anchor="w")
		ttk.Button(
			output_row,
			text="Save As...",
			style="Secondary.TButton",
			command=self._select_output,
		).pack(anchor="e", pady=(6, 0))

		ttk.Button(
			content,
			text="Generate Report",
			style="Primary.TButton",
			command=self._generate_report,
		).pack(fill="x", pady=(12, 0))

	def _build_viewer_card(self, parent: tk.Widget) -> None:
		card = self._build_card(parent, "Output Viewer")
		card.grid(row=0, column=0, sticky="nsew", pady=0)
		card.rowconfigure(2, weight=1)
		card.columnconfigure(0, weight=1)

		controls = tk.Frame(card, bg=self.colors["card_alt"], padx=8, pady=6)
		controls.pack(fill="x")

		self.table_selector = ttk.Combobox(
			controls,
			textvariable=self.selected_table_var,
			state="readonly",
			values=["All Tables"],
			width=26,
		)
		self.table_selector.pack(side="left")

		search_entry = ttk.Entry(controls, textvariable=self.search_var, width=30)
		search_entry.pack(side="left", padx=12)

		filter_frame = tk.Frame(card, bg=self.colors["card"], padx=8, pady=6)
		filter_frame.pack(fill="x")

		for column in ["Truck", "Trailer", "Position", "Status", "Type", "Return"]:
			var = tk.StringVar(value="All")
			self.filter_vars[column] = var
			col_frame = tk.Frame(filter_frame, bg=self.colors["card"])
			col_frame.pack(side="left", padx=6)
			tk.Label(
				col_frame,
				text=column,
				font=("Segoe UI", 8, "bold"),
				fg=self.colors["text_muted"],
				bg=self.colors["card"],
			).pack(anchor="w")
			combo = ttk.Combobox(
				col_frame,
				textvariable=var,
				state="readonly",
				values=["All"],
				width=14,
			)
			combo.pack()
			self.filter_widgets[column] = combo
			combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_view())

		table_frame = tk.Frame(card, bg=self.colors["card"])
		table_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
		table_frame.rowconfigure(0, weight=1)
		table_frame.columnconfigure(0, weight=1)

		self.tree = ttk.Treeview(table_frame, columns=OUTPUT_COLUMNS, show="headings")
		for col in OUTPUT_COLUMNS:
			self.tree.heading(col, text=col.upper())
			self.tree.column(col, width=120, anchor="w")
		self.tree.grid(row=0, column=0, sticky="nsew")

		y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
		y_scroll.grid(row=0, column=1, sticky="ns")
		x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
		x_scroll.grid(row=1, column=0, sticky="ew")
		self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

	def _build_log_card(self, parent: tk.Widget) -> None:
		card = self._build_card(parent, "Process Logs")
		card.grid(row=1, column=0, sticky="nsew", pady=0)
		card.rowconfigure(1, weight=1)
		card.columnconfigure(0, weight=1)

		self.log_text = tk.Text(
			card,
			bg=self.colors["log_bg"],
			fg=self.colors["log_text"],
			height=8,
			borderwidth=0,
			highlightthickness=0,
			font=("Consolas", 9),
		)
		self.log_text.pack(fill="both", expand=True, padx=8, pady=8)
		self.log_text.insert("end", "[READY] Awaiting input...\n")
		self.log_text.configure(state="disabled")

	def _bind_events(self) -> None:
		self.search_var.trace_add("write", lambda *_: self._refresh_view())
		self.table_selector.bind("<<ComboboxSelected>>", lambda _e: self._refresh_view())

	def _log(self, message: str, level: str = "INFO") -> None:
		timestamp = datetime.now().strftime("%H:%M:%S")
		line = f"[{timestamp}] [{level}] {message}\n"
		self.log_text.configure(state="normal")
		self.log_text.insert("end", line)
		self.log_text.configure(state="disabled")
		self.log_text.see("end")

	def _select_files(self) -> None:
		file_paths = filedialog.askopenfilenames(
			title="Select Excel files",
			filetypes=[("Excel files", "*.xlsx *.xlsm")],
		)
		if not file_paths:
			return

		for path in file_paths:
			candidate = Path(path)
			if candidate not in self.file_paths:
				self.file_paths.append(candidate)

		self._refresh_file_list()

	def _clear_files(self) -> None:
		self.file_paths.clear()
		self._refresh_file_list()

	def _remove_selected_file(self) -> None:
		selection = list(self.file_listbox.curselection())
		for index in reversed(selection):
			del self.file_paths[index]
		self._refresh_file_list()

	def _refresh_file_list(self) -> None:
		self.file_listbox.delete(0, "end")
		for path in self.file_paths:
			self.file_listbox.insert("end", path.name)
		self.file_count_var.set(f"{len(self.file_paths)} files selected")

	def _open_config(self) -> None:
		config_path = Path(__file__).resolve().parent / "config.py"
		if config_path.exists():
			os.startfile(config_path)

	def _select_output(self) -> None:
		default_name = f"IMPORT_REPORT_{date.today().isoformat()}.xlsx"
		file_path = filedialog.asksaveasfilename(
			title="Save merged report",
			defaultextension=".xlsx",
			initialfile=default_name,
			filetypes=[("Excel files", "*.xlsx")],
		)
		if not file_path:
			return
		self.output_path = Path(file_path)
		self.output_path_var.set(str(self.output_path))

	def _generate_report(self) -> None:
		if not self.file_paths:
			messagebox.showwarning("DataDock", "Please select Excel files first.")
			return

		if self.output_path is None:
			self._select_output()
			if self.output_path is None:
				return

		merger = ExcelMerger()
		merger.set_files(self.file_paths)
		aliases_map = get_column_aliases()
		tables = merger.merge_files(aliases_map=aliases_map)
		tables = merger.format_tables(tables)

		if merger.errors:
			for message in merger.errors:
				self._log(message, "WARN")

		if not tables:
			self._log("No tables were generated.", "ERROR")
			messagebox.showerror("DataDock", "No tables were generated.")
			return

		self.tables = tables
		self._update_table_selector()
		self._refresh_view()

		if self.output_path.exists():
			overwrite = messagebox.askyesno(
				"Overwrite file",
				"The selected file already exists. Replace it?",
			)
			if not overwrite:
				return
			try:
				self.output_path.unlink()
			except PermissionError:
				messagebox.showerror(
					"File in use",
					"Close the Excel file and try again.",
				)
				return

		try:
			merger.export_excel_tables(tables, self.output_path)
		except Exception as exc:
			self._log(str(exc), "ERROR")
			messagebox.showerror("DataDock", f"Export failed: {exc}")
			return

		self._log("Report generated successfully.", "SUCCESS")
		self.status_var.set("Report generated")

	def _update_table_selector(self) -> None:
		self.table_display_map = {"All Tables": "__all__"}
		for name in self.tables.keys():
			display = Path(name).stem
			self.table_display_map[display] = name
		values = list(self.table_display_map.keys())
		self.table_selector.configure(values=values)
		self.selected_table_var.set(values[0])

	def _get_selected_table(self) -> pd.DataFrame:
		selection = self.selected_table_var.get()
		key = self.table_display_map.get(selection)
		if key == "__all__":
			combined = []
			for name, table in self.tables.items():
				source = Path(name).stem
				combined.append(table.assign(Source=source))
			if not combined:
				return pd.DataFrame()
			df = pd.concat(combined, ignore_index=True)
			columns = ["Source"] + OUTPUT_COLUMNS
			return df.loc[:, columns]
		if key and key in self.tables:
			return self.tables[key]
		return pd.DataFrame()

	def _refresh_filters(self, df: pd.DataFrame) -> None:
		for column, var in self.filter_vars.items():
			values = ["All"]
			if column in df.columns:
				uniques = sorted({str(value) for value in df[column].dropna()})
				values.extend(uniques)
			widget = self.filter_widgets.get(column)
			if widget is not None:
				widget.configure(values=values)
			if var.get() not in values:
				var.set("All")

	def _apply_filters(self, df: pd.DataFrame) -> pd.DataFrame:
		filtered = df
		for column, var in self.filter_vars.items():
			value = var.get()
			if value != "All" and column in filtered.columns:
				filtered = filtered[filtered[column].astype(str) == value]

		search = self.search_var.get().strip().lower()
		if search:
			mask = filtered.astype(str).apply(
				lambda col: col.str.lower().str.contains(search, na=False)
			)
			filtered = filtered[mask.any(axis=1)]

		return filtered

	def _refresh_view(self) -> None:
		df = self._get_selected_table()
		if df.empty:
			self._update_treeview(pd.DataFrame())
			self.row_count_var.set("")
			return

		self._refresh_filters(df)
		filtered = self._apply_filters(df)
		self._update_treeview(filtered)
		self.row_count_var.set(f"Showing {len(filtered)} rows")

	def _update_treeview(self, df: pd.DataFrame) -> None:
		for item in self.tree.get_children():
			self.tree.delete(item)

		columns = list(df.columns) if not df.empty else OUTPUT_COLUMNS
		self.tree.configure(columns=columns)
		for col in columns:
			self.tree.heading(col, text=col.upper())
			self.tree.column(col, width=120, anchor="w")

		for _, row in df.iterrows():
			values = [row.get(col, "") for col in columns]
			self.tree.insert("", "end", values=values)


def run_app() -> None:
	root = tk.Tk()
	app = DataDockApp(root)
	root.mainloop()
