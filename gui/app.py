# gui/app.py

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os

from gui.ui.layout import build_layout
from gui.core.presets import load_presets, apply_preset
from gui.core.formula_engine import evaluate_formula
from gui.core.exporter import export_to_excel

class ExcelTransformerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Matrix Ingestor")

        self.df = None
        self.ref_df = None
        self.file_path = None
        self.config_path = "config.json"
        self.presets_path = "presets.json"
        self.header_row = tk.IntVar(value=0)

        self.column_vars = {}
        self.formula_vars = {}
        self.preview_labels = {}

        self.presets = load_presets(self.presets_path)

        build_layout(self)

    def load_main_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        self.file_path = path
        xls = pd.ExcelFile(self.file_path)
        self.sheet_dropdown["values"] = xls.sheet_names
        self.sheet_dropdown.current(0)
        self.load_sheet()

    def load_sheet(self, *args):
        self.df = pd.read_excel(self.file_path, sheet_name=self.sheet_dropdown.get(), header=self.header_row.get())
        self.column_vars.clear()
        self.formula_vars.clear()
        self.preview_labels.clear()
        for widget in self.columns_frame.winfo_children():
            widget.destroy()

        for col in self.df.columns:
            self.add_column_row(col)

    def add_column_row(self, col):
        row_frame = tk.Frame(self.columns_frame)
        row_frame.pack(anchor="w", pady=1, padx=5)

        var = tk.BooleanVar()
        chk = tk.Checkbutton(row_frame, text=col, variable=var)
        chk.grid(row=0, column=0, sticky="w")
        self.column_vars[col] = var

        formula_var = tk.StringVar()
        entry = tk.Entry(row_frame, textvariable=formula_var, width=35)
        entry.grid(row=0, column=1, padx=4)
        entry.bind("<KeyRelease>", lambda e, c=col: self.update_formula_preview(c))
        self.formula_vars[col] = formula_var

        preview_label = tk.Label(row_frame, text="", font=("Courier", 9), fg="gray")
        preview_label.grid(row=1, column=1, sticky="w")
        self.preview_labels[col] = preview_label

    def update_formula_preview(self, col):
        result = evaluate_formula(self.df, col, self.formula_vars[col].get().strip())
        self.preview_labels[col].config(text=result)

    def apply_preset(self):
        apply_preset(
            self.presets, self.preset_var.get(),
            self.column_vars, self.formula_vars, self.update_formula_preview
        )

    def load_reference_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.ref_df = pd.read_excel(path)

    def export(self):
        export_to_excel(
            self.df,
            self.column_vars,
            self.formula_vars,
            self.ref_df,
            self.main_key_entry.get(),
            self.ref_key_entry.get(),
            self.ref_cols_entry.get(),
            self.file_path
        )

    def save_config(self):
        config = {
            "selected": [col for col, var in self.column_vars.items() if var.get()],
            "formulas": {col: var.get() for col, var in self.formula_vars.items()},
            "header_row": self.header_row.get()
        }
        with open(self.config_path, "w") as f:
            json.dump(config, f)
        messagebox.showinfo("Saved", f"Config saved to {self.config_path}")

    def load_config(self):
        try:
            with open(self.config_path, "r") as f:
                config = json.load(f)
            self.header_row.set(config.get("header_row", 0))
            for col in config.get("selected", []):
                if col in self.column_vars:
                    self.column_vars[col].set(True)
            for col, formula in config.get("formulas", {}).items():
                if col in self.formula_vars:
                    self.formula_vars[col].set(formula)
                    self.update_formula_preview(col)
            messagebox.showinfo("Loaded", "Config loaded.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load config: {e}")
