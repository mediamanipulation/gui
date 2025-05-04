# gui/app.py

import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import json

class ExcelTransformerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Transformer Advanced")

        self.df = None
        self.ref_df = None
        self.file_path = None
        self.config_path = "config.json"
        self.sheet_name = None
        self.header_row = tk.IntVar(value=0)

        self.column_vars = {}
        self.formula_vars = {}

        # GUI elements
        tk.Button(root, text="Load Main Excel File", command=self.load_main_file).pack()
        tk.Label(root, text="Header Row:").pack()
        tk.Entry(root, textvariable=self.header_row).pack()
        self.sheet_dropdown = ttk.Combobox(root, state="readonly")
        self.sheet_dropdown.pack()
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.load_sheet)

        self.columns_frame = tk.Frame(root)
        self.columns_frame.pack()

        tk.Button(root, text="Load Reference File", command=self.load_reference_file).pack()

        self.ref_key_frame = tk.Frame(root)
        self.ref_key_frame.pack()
        tk.Label(self.ref_key_frame, text="Main Key:").grid(row=0, column=0)
        tk.Label(self.ref_key_frame, text="Ref Key:").grid(row=0, column=2)
        tk.Label(self.ref_key_frame, text="Ref Columns:").grid(row=0, column=4)

        self.main_key_entry = tk.Entry(self.ref_key_frame)
        self.main_key_entry.grid(row=0, column=1)
        self.ref_key_entry = tk.Entry(self.ref_key_frame)
        self.ref_key_entry.grid(row=0, column=3)
        self.ref_cols_entry = tk.Entry(self.ref_key_frame)
        self.ref_cols_entry.grid(row=0, column=5)

        tk.Button(root, text="Save Config", command=self.save_config).pack()
        tk.Button(root, text="Load Config", command=self.load_config).pack()
        tk.Button(root, text="Export", command=self.export).pack()

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
        sheet = self.sheet_dropdown.get()
        if not sheet:
            return
        df = pd.read_excel(self.file_path, sheet_name=sheet, header=self.header_row.get())
        self.df = df
        self.column_vars.clear()
        self.formula_vars.clear()
        for widget in self.columns_frame.winfo_children():
            widget.destroy()
        for i, col in enumerate(df.columns):
            var = tk.BooleanVar()
            self.column_vars[col] = var
            tk.Checkbutton(self.columns_frame, text=col, variable=var).grid(row=i, column=0, sticky='w')
            formula_var = tk.StringVar()
            self.formula_vars[col] = formula_var
            tk.Entry(self.columns_frame, textvariable=formula_var, width=30).grid(row=i, column=1)

    def load_reference_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.ref_df = pd.read_excel(path)

    def export(self):
        selected_cols = [col for col, var in self.column_vars.items() if var.get()]
        if not selected_cols:
            messagebox.showwarning("No columns selected", "Select at least one column.")
            return
        result = pd.DataFrame()
        for col in selected_cols:
            formula = self.formula_vars[col].get().strip()
            if formula:
                try:
                    result[col] = self.df.eval(formula)
                except Exception as e:
                    result[col] = f"#ERR {e}"
            else:
                result[col] = self.df[col]

        if self.ref_df is not None:
            main_key = self.main_key_entry.get().strip()
            ref_key = self.ref_key_entry.get().strip()
            ref_cols = [c.strip() for c in self.ref_cols_entry.get().split(',') if c.strip()]
            if main_key and ref_key and ref_cols:
                ref_slice = self.ref_df[[ref_key] + ref_cols]
                result = result.merge(ref_slice, left_on=main_key, right_on=ref_key, how="left")

        out_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if out_path:
            result.to_excel(out_path, index=False)
            messagebox.showinfo("Done", f"Saved to {out_path}")

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
            messagebox.showinfo("Loaded", "Config loaded.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load config: {e}")
