import os
import uuid
from dataclasses import dataclass, field

import numpy as np
import pandas as pd
import tkinter as tk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
from matplotlib.lines import Line2D
from tkinter import filedialog, messagebox, ttk


STAGE_COLORS = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728",
    "#9467bd", "#8c564b", "#e377c2", "#17becf",
]


def default_config() -> dict:
    return {
        "format": "col_per_dim",
        "header_row": 0,
        "skip_cols": 0,
        "nominal_row": None,
        "usl_row": None,
        "lsl_row": None,
        "upper_tol_row": None,
        "lower_tol_row": None,
        "data_start_row": 1,
        "name_col": 0,
        "first_data_row": 1,
        "nominal_col": None,
        "usl_col": None,
        "lsl_col": None,
        "upper_tol_col": None,
        "lower_tol_col": None,
        "data_start_col": 4,
    }


def _safe_float(value):
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _cell_display(value):
    if value is None:
        return ""
    if isinstance(value, float):
        if np.isnan(value):
            return ""
        return round(value, 6)
    return str(value)


def _read_excel(path: str, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet_name, header=None)


def load_preview_rows(path: str, sheet_name: str, max_rows: int = 20) -> list[list[str]]:
    df = _read_excel(path, sheet_name)
    rows = []
    for row_index in range(min(max_rows, len(df))):
        rows.append([_cell_display(df.iloc[row_index, column_index]) for column_index in range(len(df.columns))])
    return rows


def extract_dimensions(path: str, sheet_name: str, config: dict) -> list[str]:
    df = _read_excel(path, sheet_name)
    dims = []
    fmt = config.get("format", "col_per_dim")

    if fmt == "col_per_dim":
        header_row = int(config.get("header_row", 0))
        skip_cols = int(config.get("skip_cols", 0))
        for column_index in range(skip_cols, len(df.columns)):
            value = df.iloc[header_row, column_index]
            name = str(value).strip() if pd.notna(value) else ""
            if name and name.lower() != "nan":
                dims.append(name)
    else:
        name_col = int(config.get("name_col", 0))
        first_data_row = int(config.get("first_data_row", 1))
        for row_index in range(first_data_row, len(df)):
            value = df.iloc[row_index, name_col]
            name = str(value).strip() if pd.notna(value) else ""
            if name and name.lower() != "nan":
                dims.append(name)

    return dims


def build_stage_dimensions(path: str, sheet_name: str, config: dict, selected_dims: set[str]) -> list[dict]:
    df = _read_excel(path, sheet_name)
    fmt = config.get("format", "col_per_dim")
    dimensions = []

    if fmt == "col_per_dim":
        header_row = int(config.get("header_row", 0))
        skip_cols = int(config.get("skip_cols", 0))
        nominal_row = config.get("nominal_row")
        usl_row = config.get("usl_row")
        lsl_row = config.get("lsl_row")
        upper_tol_row = config.get("upper_tol_row")
        lower_tol_row = config.get("lower_tol_row")
        data_start_row = int(config.get("data_start_row", 1))

        def row_value(row_index, column_index):
            if row_index is None:
                return None
            row_number = int(row_index)
            if 0 <= row_number < len(df):
                return _safe_float(df.iloc[row_number, column_index])
            return None

        for column_index in range(skip_cols, len(df.columns)):
            raw_name = df.iloc[header_row, column_index]
            name = str(raw_name).strip() if pd.notna(raw_name) else ""
            if not name or name.lower() == "nan":
                continue
            if selected_dims and name not in selected_dims:
                continue

            nominal = row_value(nominal_row, column_index)
            usl = row_value(usl_row, column_index)
            lsl = row_value(lsl_row, column_index)
            upper_tol = row_value(upper_tol_row, column_index)
            lower_tol = row_value(lower_tol_row, column_index)

            if usl is None and upper_tol is not None and nominal is not None:
                usl = nominal + upper_tol
            if lsl is None and lower_tol is not None and nominal is not None:
                lsl = nominal - abs(lower_tol)

            measurements = []
            for row_index in range(data_start_row, len(df)):
                value = _safe_float(df.iloc[row_index, column_index])
                if value is not None:
                    measurements.append(value)

            dimensions.append(
                {
                    "name": name,
                    "nominal": nominal,
                    "usl": usl,
                    "lsl": lsl,
                    "measurements": measurements,
                }
            )
    else:
        name_col = int(config.get("name_col", 0))
        nominal_col = config.get("nominal_col")
        usl_col = config.get("usl_col")
        lsl_col = config.get("lsl_col")
        upper_tol_col = config.get("upper_tol_col")
        lower_tol_col = config.get("lower_tol_col")
        data_start_col = int(config.get("data_start_col", 4))
        first_data_row = int(config.get("first_data_row", 1))

        def col_value(row_index, column_index):
            if column_index is None:
                return None
            column_number = int(column_index)
            if 0 <= column_number < len(df.columns):
                return _safe_float(df.iloc[row_index, column_number])
            return None

        for row_index in range(first_data_row, len(df)):
            raw_name = df.iloc[row_index, name_col]
            name = str(raw_name).strip() if pd.notna(raw_name) else ""
            if not name or name.lower() == "nan":
                continue
            if selected_dims and name not in selected_dims:
                continue

            nominal = col_value(row_index, nominal_col)
            usl = col_value(row_index, usl_col)
            lsl = col_value(row_index, lsl_col)
            upper_tol = col_value(row_index, upper_tol_col)
            lower_tol = col_value(row_index, lower_tol_col)

            if usl is None and upper_tol is not None and nominal is not None:
                usl = nominal + upper_tol
            if lsl is None and lower_tol is not None and nominal is not None:
                lsl = nominal - abs(lower_tol)

            measurements = []
            for column_index in range(data_start_col, len(df.columns)):
                value = _safe_float(df.iloc[row_index, column_index])
                if value is not None:
                    measurements.append(value)

            dimensions.append(
                {
                    "name": name,
                    "nominal": nominal,
                    "usl": usl,
                    "lsl": lsl,
                    "measurements": measurements,
                }
            )

    return dimensions


@dataclass
class FileEntry:
    file_id: str
    path: str
    filename: str
    stage_name: str
    sheet_names: list[str]
    selected_sheet: str
    config: dict = field(default_factory=default_config)
    dimensions: list[str] = field(default_factory=list)


class ConfigDialog(tk.Toplevel):
    def __init__(self, master, entry: FileEntry, on_apply):
        super().__init__(master)
        self.entry = entry
        self.on_apply = on_apply
        self.title(f"配置数据格式 - {entry.filename}")
        self.geometry("1050x720")
        self.transient(master)
        self.grab_set()

        self.format_var = tk.StringVar(value=entry.config.get("format", "col_per_dim"))
        self.vars = self._build_vars(entry.config)

        self._build_ui()
        self._toggle_format_frames()
        self._load_preview()

    def _build_vars(self, config: dict) -> dict[str, tk.StringVar]:
        def display_value(value):
            if value is None:
                return ""
            return str(int(value) + 1)

        return {
            "header_row": tk.StringVar(value=display_value(config.get("header_row", 0))),
            "skip_cols": tk.StringVar(value=str(config.get("skip_cols", 0))),
            "nominal_row": tk.StringVar(value=display_value(config.get("nominal_row"))),
            "usl_row": tk.StringVar(value=display_value(config.get("usl_row"))),
            "lsl_row": tk.StringVar(value=display_value(config.get("lsl_row"))),
            "upper_tol_row": tk.StringVar(value=display_value(config.get("upper_tol_row"))),
            "lower_tol_row": tk.StringVar(value=display_value(config.get("lower_tol_row"))),
            "data_start_row": tk.StringVar(value=display_value(config.get("data_start_row", 1))),
            "name_col": tk.StringVar(value=display_value(config.get("name_col", 0))),
            "first_data_row": tk.StringVar(value=display_value(config.get("first_data_row", 1))),
            "nominal_col": tk.StringVar(value=display_value(config.get("nominal_col"))),
            "usl_col": tk.StringVar(value=display_value(config.get("usl_col"))),
            "lsl_col": tk.StringVar(value=display_value(config.get("lsl_col"))),
            "upper_tol_col": tk.StringVar(value=display_value(config.get("upper_tol_col"))),
            "lower_tol_col": tk.StringVar(value=display_value(config.get("lower_tol_col"))),
            "data_start_col": tk.StringVar(value=display_value(config.get("data_start_col", 4))),
        }

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        header = ttk.Frame(self, padding=12)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text=self.entry.filename, font=("Segoe UI", 12, "bold")).pack(anchor="w")
        ttk.Label(header, text=f"当前 Sheet: {self.entry.selected_sheet}").pack(anchor="w", pady=(4, 0))

        format_frame = ttk.LabelFrame(self, text="数据格式", padding=12)
        format_frame.grid(row=1, column=0, sticky="ew", padx=12)
        ttk.Radiobutton(
            format_frame,
            text="每列一个尺寸",
            value="col_per_dim",
            variable=self.format_var,
            command=self._toggle_format_frames,
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))
        ttk.Radiobutton(
            format_frame,
            text="每行一个尺寸",
            value="row_per_dim",
            variable=self.format_var,
            command=self._toggle_format_frames,
        ).grid(row=0, column=1, sticky="w")

        content = ttk.Frame(self, padding=(12, 8, 12, 0))
        content.grid(row=2, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(2, weight=1)

        self.col_frame = ttk.LabelFrame(content, text="列模式参数", padding=12)
        self.col_frame.grid(row=0, column=0, sticky="ew")
        self.row_frame = ttk.LabelFrame(content, text="行模式参数", padding=12)
        self.row_frame.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        self._build_form(self.col_frame, [
            ("尺寸名行号", "header_row"),
            ("跳过前 N 列", "skip_cols"),
            ("理论值行号", "nominal_row"),
            ("USL 行号", "usl_row"),
            ("LSL 行号", "lsl_row"),
            ("上公差行号", "upper_tol_row"),
            ("下公差行号", "lower_tol_row"),
            ("数据起始行号", "data_start_row"),
        ])
        self._build_form(self.row_frame, [
            ("尺寸名列号", "name_col"),
            ("首行数据行号", "first_data_row"),
            ("理论值列号", "nominal_col"),
            ("USL 列号", "usl_col"),
            ("LSL 列号", "lsl_col"),
            ("上公差列号", "upper_tol_col"),
            ("下公差列号", "lower_tol_col"),
            ("测量数据起始列号", "data_start_col"),
        ])

        preview_frame = ttk.LabelFrame(content, text="数据预览（前 20 行）", padding=8)
        preview_frame.grid(row=2, column=0, sticky="nsew", pady=(8, 0))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        self.preview_tree = ttk.Treeview(preview_frame, show="headings")
        preview_scroll_y = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        preview_scroll_y.grid(row=0, column=1, sticky="ns")
        preview_scroll_x.grid(row=1, column=0, sticky="ew")

        footer = ttk.Frame(self, padding=12)
        footer.grid(row=3, column=0, sticky="ew")
        ttk.Button(footer, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(footer, text="提取尺寸", command=self._apply).pack(side="right", padx=(0, 8))

    def _build_form(self, parent, fields: list[tuple[str, str]]):
        for column_index in range(4):
            parent.columnconfigure(column_index, weight=1)
        for row_index, (label, key) in enumerate(fields):
            row = row_index // 2
            column = (row_index % 2) * 2
            ttk.Label(parent, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=4)
            ttk.Entry(parent, textvariable=self.vars[key], width=12).grid(row=row, column=column + 1, sticky="ew", pady=4)

    def _toggle_format_frames(self):
        if self.format_var.get() == "col_per_dim":
            self.col_frame.grid()
            self.row_frame.grid_remove()
        else:
            self.row_frame.grid()
            self.col_frame.grid_remove()

    def _load_preview(self):
        rows = load_preview_rows(self.entry.path, self.entry.selected_sheet)
        column_count = max((len(row) for row in rows), default=0)
        columns = [str(index + 1) for index in range(column_count)]
        self.preview_tree.configure(columns=columns)
        for column in columns:
            self.preview_tree.heading(column, text=column)
            self.preview_tree.column(column, width=110, anchor="center")
        self.preview_tree.delete(*self.preview_tree.get_children())
        for row in rows:
            values = row + [""] * (column_count - len(row))
            self.preview_tree.insert("", "end", values=values)

    def _one_based_to_zero_based(self, value: str):
        stripped = value.strip()
        if not stripped:
            return None
        return int(stripped) - 1

    def _apply(self):
        try:
            if self.format_var.get() == "col_per_dim":
                config = {
                    "format": "col_per_dim",
                    "header_row": self._one_based_to_zero_based(self.vars["header_row"].get()) or 0,
                    "skip_cols": int(self.vars["skip_cols"].get().strip() or "0"),
                    "nominal_row": self._one_based_to_zero_based(self.vars["nominal_row"].get()),
                    "usl_row": self._one_based_to_zero_based(self.vars["usl_row"].get()),
                    "lsl_row": self._one_based_to_zero_based(self.vars["lsl_row"].get()),
                    "upper_tol_row": self._one_based_to_zero_based(self.vars["upper_tol_row"].get()),
                    "lower_tol_row": self._one_based_to_zero_based(self.vars["lower_tol_row"].get()),
                    "data_start_row": self._one_based_to_zero_based(self.vars["data_start_row"].get()) or 1,
                }
            else:
                config = {
                    "format": "row_per_dim",
                    "name_col": self._one_based_to_zero_based(self.vars["name_col"].get()) or 0,
                    "first_data_row": self._one_based_to_zero_based(self.vars["first_data_row"].get()) or 1,
                    "nominal_col": self._one_based_to_zero_based(self.vars["nominal_col"].get()),
                    "usl_col": self._one_based_to_zero_based(self.vars["usl_col"].get()),
                    "lsl_col": self._one_based_to_zero_based(self.vars["lsl_col"].get()),
                    "upper_tol_col": self._one_based_to_zero_based(self.vars["upper_tol_col"].get()),
                    "lower_tol_col": self._one_based_to_zero_based(self.vars["lower_tol_col"].get()),
                    "data_start_col": self._one_based_to_zero_based(self.vars["data_start_col"].get()) or 4,
                }

            dimensions = extract_dimensions(self.entry.path, self.entry.selected_sheet, config)
        except Exception as exc:
            messagebox.showerror("配置失败", f"提取尺寸时出错：\n{exc}", parent=self)
            return

        self.entry.config = config
        self.entry.dimensions = dimensions
        self.on_apply()
        self.destroy()


class BoxPlotApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Box Plot 分析工具")
        self.geometry("1480x900")
        self.minsize(1200, 760)

        self.files: list[FileEntry] = []
        self.current_figure = None
        self._refreshing_selection = False

        self.stage_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.title_var = tk.StringVar(value="尺寸测量 Box Plot")
        self.y_axis_var = tk.StringVar(value="测量值")
        self.normalize_var = tk.BooleanVar(value=False)
        self.show_spec_var = tk.BooleanVar(value=True)
        self.show_points_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="请选择 Excel 文件开始分析。")

        self._build_ui()
        self.stage_var.trace_add("write", self._on_stage_changed)

    def _build_ui(self):
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")

        container = ttk.PanedWindow(self, orient="horizontal")
        container.pack(fill="both", expand=True)

        left = ttk.Frame(container, padding=12)
        right = ttk.Frame(container, padding=(0, 12, 12, 12))
        container.add(left, weight=2)
        container.add(right, weight=5)

        left.columnconfigure(0, weight=1)
        left.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=3)
        right.rowconfigure(1, weight=2)

        file_actions = ttk.Frame(left)
        file_actions.grid(row=0, column=0, sticky="ew")
        ttk.Button(file_actions, text="添加 Excel", command=self.add_files).pack(side="left")
        ttk.Button(file_actions, text="删除所选", command=self.remove_selected_file).pack(side="left", padx=(8, 0))
        ttk.Button(file_actions, text="配置所选文件", command=self.open_config_dialog).pack(side="left", padx=(8, 0))

        file_frame = ttk.LabelFrame(left, text="文件列表", padding=8)
        file_frame.grid(row=1, column=0, sticky="nsew", pady=(10, 10))
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(0, weight=1)

        self.file_tree = ttk.Treeview(
            file_frame,
            columns=("stage", "sheet", "dimensions"),
            show="tree headings",
            selectmode="browse",
            height=12,
        )
        self.file_tree.heading("#0", text="文件")
        self.file_tree.heading("stage", text="阶段")
        self.file_tree.heading("sheet", text="Sheet")
        self.file_tree.heading("dimensions", text="尺寸数")
        self.file_tree.column("#0", width=220, anchor="w")
        self.file_tree.column("stage", width=110, anchor="center")
        self.file_tree.column("sheet", width=130, anchor="center")
        self.file_tree.column("dimensions", width=70, anchor="center")
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        self.file_tree.bind("<<TreeviewSelect>>", self._on_file_selected)
        file_scroll = ttk.Scrollbar(file_frame, orient="vertical", command=self.file_tree.yview)
        file_scroll.grid(row=0, column=1, sticky="ns")
        self.file_tree.configure(yscrollcommand=file_scroll.set)

        detail_frame = ttk.LabelFrame(left, text="当前文件设置", padding=10)
        detail_frame.grid(row=2, column=0, sticky="ew")
        detail_frame.columnconfigure(1, weight=1)
        ttk.Label(detail_frame, text="阶段名称").grid(row=0, column=0, sticky="w")
        ttk.Entry(detail_frame, textvariable=self.stage_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Label(detail_frame, text="Sheet").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.sheet_combo = ttk.Combobox(detail_frame, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        self.sheet_combo.bind("<<ComboboxSelected>>", self._on_sheet_changed)

        dim_frame = ttk.LabelFrame(left, text="尺寸选择", padding=10)
        dim_frame.grid(row=3, column=0, sticky="nsew", pady=(10, 10))
        dim_frame.columnconfigure(0, weight=1)
        dim_frame.rowconfigure(0, weight=1)
        self.dim_listbox = tk.Listbox(dim_frame, selectmode=tk.EXTENDED, exportselection=False, height=10)
        self.dim_listbox.grid(row=0, column=0, sticky="nsew")
        dim_scroll = ttk.Scrollbar(dim_frame, orient="vertical", command=self.dim_listbox.yview)
        dim_scroll.grid(row=0, column=1, sticky="ns")
        self.dim_listbox.configure(yscrollcommand=dim_scroll.set)
        dim_actions = ttk.Frame(dim_frame)
        dim_actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(dim_actions, text="全选", command=self.select_all_dimensions).pack(side="left")
        ttk.Button(dim_actions, text="清空", command=self.clear_dimension_selection).pack(side="left", padx=(8, 0))

        options = ttk.LabelFrame(left, text="图表设置", padding=10)
        options.grid(row=4, column=0, sticky="ew")
        options.columnconfigure(1, weight=1)
        ttk.Label(options, text="图表标题").grid(row=0, column=0, sticky="w")
        ttk.Entry(options, textvariable=self.title_var).grid(row=0, column=1, sticky="ew", padx=(8, 0))
        ttk.Label(options, text="Y 轴标题").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(options, textvariable=self.y_axis_var).grid(row=1, column=1, sticky="ew", padx=(8, 0), pady=(8, 0))
        ttk.Checkbutton(options, text="偏差显示（减去理论值）", variable=self.normalize_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Checkbutton(options, text="显示规格线（USL / LSL / 理论值）", variable=self.show_spec_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4, 0))
        ttk.Checkbutton(options, text="显示散点", variable=self.show_points_var).grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        plot_actions = ttk.Frame(left)
        plot_actions.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(plot_actions, text="生成 Box Plot", command=self.generate_plot).pack(side="left")
        ttk.Button(plot_actions, text="导出 PNG", command=self.export_png).pack(side="left", padx=(8, 0))

        plot_frame = ttk.LabelFrame(right, text="图表", padding=8)
        plot_frame.grid(row=0, column=0, sticky="nsew")
        plot_frame.columnconfigure(0, weight=1)
        plot_frame.rowconfigure(0, weight=1)

        self.figure = Figure(figsize=(10, 5), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=plot_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.grid(row=0, column=0, sticky="nsew")
        self.toolbar = NavigationToolbar2Tk(self.canvas, plot_frame, pack_toolbar=False)
        self.toolbar.update()
        self.toolbar.grid(row=1, column=0, sticky="ew")

        stats_frame = ttk.LabelFrame(right, text="统计摘要", padding=8)
        stats_frame.grid(row=1, column=0, sticky="nsew", pady=(12, 0))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        stats_columns = ("stage", "dimension", "nominal", "usl", "lsl", "n", "min", "max", "mean", "std", "cpk")
        self.stats_tree = ttk.Treeview(stats_frame, columns=stats_columns, show="headings")
        headings = {
            "stage": "阶段",
            "dimension": "尺寸",
            "nominal": "理论值",
            "usl": "USL",
            "lsl": "LSL",
            "n": "N",
            "min": "Min",
            "max": "Max",
            "mean": "Mean",
            "std": "Std",
            "cpk": "Cpk",
        }
        for key in stats_columns:
            self.stats_tree.heading(key, text=headings[key])
            self.stats_tree.column(key, width=92, anchor="center")
        self.stats_tree.column("dimension", width=150, anchor="w")
        self.stats_tree.grid(row=0, column=0, sticky="nsew")
        stats_scroll = ttk.Scrollbar(stats_frame, orient="vertical", command=self.stats_tree.yview)
        stats_scroll.grid(row=0, column=1, sticky="ns")
        self.stats_tree.configure(yscrollcommand=stats_scroll.set)

        status = ttk.Label(self, textvariable=self.status_var, anchor="w", padding=(12, 6))
        status.pack(fill="x")

        self._render_empty_plot()

    def _render_empty_plot(self):
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        ax.axis("off")
        ax.text(0.5, 0.5, "请先添加 Excel 文件并配置尺寸。", ha="center", va="center", fontsize=14)
        self.canvas.draw_idle()

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel files", "*.xlsx *.xls")],
        )
        if not paths:
            return

        added = 0
        for path in paths:
            try:
                sheet_names = pd.ExcelFile(path).sheet_names
            except Exception as exc:
                messagebox.showerror("无法读取文件", f"{os.path.basename(path)}\n\n{exc}")
                continue

            entry = FileEntry(
                file_id=str(uuid.uuid4()),
                path=path,
                filename=os.path.basename(path),
                stage_name=f"阶段 {len(self.files) + 1}",
                sheet_names=sheet_names,
                selected_sheet=sheet_names[0] if sheet_names else "",
            )
            self.files.append(entry)
            added += 1

        if added:
            self._refresh_file_tree()
            self._refresh_dimension_list()
            self.status_var.set(f"已添加 {added} 个文件。请选择文件并配置数据格式。")

    def _refresh_file_tree(self):
        selection = self.file_tree.selection()
        selected_id = selection[0] if selection else None
        self.file_tree.delete(*self.file_tree.get_children())
        for entry in self.files:
            self.file_tree.insert(
                "",
                "end",
                iid=entry.file_id,
                text=entry.filename,
                values=(entry.stage_name, entry.selected_sheet, len(entry.dimensions)),
            )

        self._refreshing_selection = True
        try:
            if selected_id and self.file_tree.exists(selected_id):
                self.file_tree.selection_set(selected_id)
                self.file_tree.focus(selected_id)
            elif self.files:
                self.file_tree.selection_set(self.files[0].file_id)
                self.file_tree.focus(self.files[0].file_id)
            else:
                self.file_tree.selection_remove(self.file_tree.selection())
        finally:
            self._refreshing_selection = False
        self._on_file_selected()

    def _get_selected_entry(self):
        selection = self.file_tree.selection()
        if not selection:
            return None
        file_id = selection[0]
        return next((entry for entry in self.files if entry.file_id == file_id), None)

    def _on_file_selected(self, event=None):
        entry = self._get_selected_entry()
        if not entry:
            self.stage_var.set("")
            self.sheet_var.set("")
            self.sheet_combo["values"] = []
            return
        self.stage_var.set(entry.stage_name)
        self.sheet_combo["values"] = entry.sheet_names
        self.sheet_var.set(entry.selected_sheet)

    def _on_stage_changed(self, *args):
        if self._refreshing_selection:
            return
        entry = self._get_selected_entry()
        if not entry:
            return
        new_value = self.stage_var.get().strip()
        if not new_value or new_value == entry.stage_name:
            return
        entry.stage_name = new_value
        self._refresh_file_tree()

    def _on_sheet_changed(self, event=None):
        entry = self._get_selected_entry()
        if not entry:
            return
        entry.selected_sheet = self.sheet_var.get()
        entry.dimensions = []
        self._refresh_file_tree()
        self._refresh_dimension_list()
        self.status_var.set(f"{entry.filename} 已切换到 Sheet: {entry.selected_sheet}，请重新提取尺寸。")

    def remove_selected_file(self):
        entry = self._get_selected_entry()
        if not entry:
            messagebox.showinfo("提示", "请先选择一个文件。")
            return
        self.files = [item for item in self.files if item.file_id != entry.file_id]
        self._refresh_file_tree()
        self._refresh_dimension_list()
        if not self.files:
            self._render_empty_plot()
            self._clear_stats_table()
        self.status_var.set(f"已删除文件：{entry.filename}")

    def open_config_dialog(self):
        entry = self._get_selected_entry()
        if not entry:
            messagebox.showinfo("提示", "请先选择一个文件。")
            return
        ConfigDialog(self, entry, self._on_config_applied)

    def _on_config_applied(self):
        self._refresh_file_tree()
        self._refresh_dimension_list()
        total_dims = len(self._all_dimensions())
        self.status_var.set(f"尺寸提取完成，当前可选尺寸数：{total_dims}")

    def _all_dimensions(self) -> list[str]:
        return sorted({dim for entry in self.files for dim in entry.dimensions})

    def _refresh_dimension_list(self):
        current_selection = set(self.get_selected_dimensions())
        all_dims = self._all_dimensions()
        self.dim_listbox.delete(0, tk.END)
        for dim in all_dims:
            self.dim_listbox.insert(tk.END, dim)
        if not all_dims:
            return
        if not current_selection:
            self.select_all_dimensions()
            return
        for index, dim in enumerate(all_dims):
            if dim in current_selection:
                self.dim_listbox.selection_set(index)

    def get_selected_dimensions(self) -> list[str]:
        return [self.dim_listbox.get(index) for index in self.dim_listbox.curselection()]

    def select_all_dimensions(self):
        self.dim_listbox.selection_set(0, tk.END)

    def clear_dimension_selection(self):
        self.dim_listbox.selection_clear(0, tk.END)

    def _collect_plot_data(self, selected_dims: list[str]) -> list[dict]:
        result = []
        selected_set = set(selected_dims)
        for entry in self.files:
            if not entry.selected_sheet:
                continue
            try:
                dimensions = build_stage_dimensions(entry.path, entry.selected_sheet, entry.config, selected_set)
            except Exception as exc:
                raise RuntimeError(f"读取 {entry.filename} 失败：{exc}") from exc
            result.append({"stage": entry.stage_name, "dimensions": dimensions})
        return result

    def generate_plot(self):
        if not self.files:
            messagebox.showinfo("提示", "请先添加 Excel 文件。")
            return

        selected_dims = self.get_selected_dimensions()
        if not selected_dims:
            messagebox.showinfo("提示", "请至少选择一个尺寸。")
            return

        if not any(entry.dimensions for entry in self.files):
            messagebox.showinfo("提示", "请先为至少一个文件配置数据格式并提取尺寸。")
            return

        try:
            plot_data = self._collect_plot_data(selected_dims)
        except RuntimeError as exc:
            messagebox.showerror("生成图表失败", str(exc))
            return

        if not any(stage["dimensions"] for stage in plot_data):
            messagebox.showinfo("提示", "当前选择下没有可绘制的数据。")
            return

        self._render_plot(plot_data, selected_dims)
        self._render_stats_table(plot_data, selected_dims)
        self.status_var.set("图表已生成。")

    def _render_plot(self, plot_data: list[dict], selected_dims: list[str]):
        self.figure.clear()
        ax = self.figure.add_subplot(111)

        normalize = self.normalize_var.get()
        show_spec = self.show_spec_var.get()
        show_points = self.show_points_var.get()
        stage_count = max(1, len(plot_data))
        width = 0.72 / stage_count
        centers = np.arange(1, len(selected_dims) + 1)
        rng = np.random.default_rng(42)
        spec_info = {}
        legend_handles = []

        for stage_index, stage in enumerate(plot_data):
            color = STAGE_COLORS[stage_index % len(STAGE_COLORS)]
            offset = (stage_index - (stage_count - 1) / 2) * width
            positions = []
            datasets = []

            for dim_index, dim_name in enumerate(selected_dims):
                dimension = next((item for item in stage["dimensions"] if item["name"] == dim_name), None)
                if not dimension or not dimension["measurements"]:
                    continue
                if dim_name not in spec_info:
                    spec_info[dim_name] = {
                        "nominal": dimension["nominal"],
                        "usl": dimension["usl"],
                        "lsl": dimension["lsl"],
                    }

                measurements = list(dimension["measurements"])
                nominal = dimension["nominal"]
                if normalize and nominal is not None:
                    measurements = [value - nominal for value in measurements]

                datasets.append(measurements)
                positions.append(centers[dim_index] + offset)

            if not datasets:
                continue

            boxplot = ax.boxplot(
                datasets,
                positions=positions,
                widths=width * 0.82,
                patch_artist=True,
                showfliers=show_points,
                manage_ticks=False,
            )
            for box in boxplot["boxes"]:
                box.set(facecolor=color, alpha=0.24, edgecolor=color, linewidth=1.3)
            for median in boxplot["medians"]:
                median.set(color=color, linewidth=1.8)
            for line in boxplot["whiskers"] + boxplot["caps"]:
                line.set(color=color, linewidth=1.1)

            if show_points:
                for pos, values in zip(positions, datasets):
                    jitter = rng.uniform(-width * 0.18, width * 0.18, len(values))
                    ax.scatter(np.full(len(values), pos) + jitter, values, s=16, alpha=0.55, color=color, zorder=3)

            legend_handles.append(Line2D([0], [0], color=color, linewidth=8, alpha=0.6, label=stage["stage"]))

        if show_spec:
            for center, dim_name in zip(centers, selected_dims):
                spec = spec_info.get(dim_name)
                if not spec:
                    continue
                nominal = spec.get("nominal")
                usl = spec.get("usl")
                lsl = spec.get("lsl")
                xmin = center - 0.42
                xmax = center + 0.42
                if normalize and nominal is not None:
                    if usl is not None:
                        usl = usl - nominal
                    if lsl is not None:
                        lsl = lsl - nominal
                    nominal = 0.0
                if nominal is not None:
                    ax.hlines(nominal, xmin, xmax, colors="#fd7e14", linewidth=2.0)
                if usl is not None:
                    ax.hlines(usl, xmin, xmax, colors="#d62728", linewidth=2.0)
                if lsl is not None:
                    ax.hlines(lsl, xmin, xmax, colors="#2ca02c", linewidth=2.0)

        ax.set_title(self.title_var.get().strip() or "Box Plot", fontsize=15)
        ax.set_ylabel("偏差值" if normalize else (self.y_axis_var.get().strip() or "测量值"))
        ax.set_xticks(centers)
        ax.set_xticklabels(selected_dims, rotation=20, ha="right")
        ax.grid(axis="y", linestyle="--", linewidth=0.7, alpha=0.35)
        ax.margins(x=0.04)

        if legend_handles:
            spec_handles = []
            if show_spec:
                spec_handles = [
                    Line2D([0], [0], color="#fd7e14", linewidth=2, label="理论值"),
                    Line2D([0], [0], color="#d62728", linewidth=2, label="USL"),
                    Line2D([0], [0], color="#2ca02c", linewidth=2, label="LSL"),
                ]
            ax.legend(handles=legend_handles + spec_handles, loc="upper right", fontsize=9)

        self.figure.tight_layout()
        self.current_figure = self.figure
        self.canvas.draw_idle()

    def _render_stats_table(self, plot_data: list[dict], selected_dims: list[str]):
        self._clear_stats_table()
        selected_set = set(selected_dims)
        for stage in plot_data:
            for dim in stage["dimensions"]:
                if dim["name"] not in selected_set:
                    continue
                measurements = dim["measurements"]
                if not measurements:
                    continue
                mean = float(np.mean(measurements))
                std = float(np.std(measurements))
                cpk = "-"
                if dim["usl"] is not None and dim["lsl"] is not None and std > 0:
                    cpu = (dim["usl"] - mean) / (3 * std)
                    cpl = (mean - dim["lsl"]) / (3 * std)
                    cpk = f"{min(cpu, cpl):.3f}"
                self.stats_tree.insert(
                    "",
                    "end",
                    values=(
                        stage["stage"],
                        dim["name"],
                        self._fmt(dim["nominal"]),
                        self._fmt(dim["usl"]),
                        self._fmt(dim["lsl"]),
                        len(measurements),
                        self._fmt(min(measurements)),
                        self._fmt(max(measurements)),
                        self._fmt(mean),
                        self._fmt(std),
                        cpk,
                    ),
                )

    def _clear_stats_table(self):
        self.stats_tree.delete(*self.stats_tree.get_children())

    def _fmt(self, value):
        if value is None:
            return "-"
        return f"{float(value):.4f}"

    def export_png(self):
        if self.current_figure is None:
            messagebox.showinfo("提示", "请先生成图表。")
            return
        path = filedialog.asksaveasfilename(
            title="导出 PNG",
            defaultextension=".png",
            filetypes=[("PNG image", "*.png")],
        )
        if not path:
            return
        self.current_figure.savefig(path, dpi=200, bbox_inches="tight")
        self.status_var.set(f"图表已导出到：{path}")


def main():
    app = BoxPlotApp()
    app.mainloop()


if __name__ == "__main__":
    main()
