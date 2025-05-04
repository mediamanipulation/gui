# 🧰 Excel Transformer GUI

A local-first, interactive Excel transformation tool built with Python and Tkinter.

This app empowers users to visually:
- Load and preview Excel sheets
- Select specific columns to keep
- Apply custom formulas to data
- Join in reference data (VLOOKUP-style)
- Export clean, transformed spreadsheets

---

## 🚀 Features

### ✅ Load Excel Files
- Select any `.xlsx` or `.xls` file
- Choose which sheet to work with
- Specify which row contains the header (default: 0)

### ✅ Visual Column Picker
- Interactive checkboxes for column inclusion
- Optional formula field per column using `pandas.eval()`
  - Example: `Salary * 0.25`, `round(Salary - Bonus, 2)`

### ✅ Reference File Join
- Load a secondary Excel file
- Define:
  - Key column from main file
  - Key column from reference file
  - Columns to pull from reference
- Automatically performs a **left join**

### ✅ Save / Load Configurations
- Automatically stores your:
  - Selected columns
  - Formula mappings
  - Header row index
- Reload instantly for repeated workflows

### ✅ Export Transformed Excel
- Outputs a clean `.xlsx` file
- Includes selected columns, calculated values, and joined reference data

---

## 🛠 Requirements

Tested on Python 3.11 with Conda.

```bash
conda install pandas openpyxl
