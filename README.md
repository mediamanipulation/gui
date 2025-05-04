# 🧰 Matrix Ingestor

A powerful, local-first Excel transformation toolkit with a modern UI built using Python and Tkinter.

Matrix Ingestor empowers users to visually manipulate, transform and join Excel data with a polished, intuitive interface, prioritizing safety with controlled formula evaluation.

#### Matrix Ingestor Interface Screenshot
<img src="./docs/scrn01.jpg" alt="Matrix Ingestor Interface Screenshot" width="300" />
<br/>

#### Matrix Ingestor Interface Screenshot Source Loaded
<img src="./docs/scrn02.jpg" alt="Matrix Ingestor Interface Screenshot" width="400" />

## 🚀 Features

### ✅ Interactive Excel Loading
- Open any `.xlsx` or `.xls` file with intuitive file browser
- Seamlessly navigate multi-sheet workbooks
- Flexible header row configuration
- Automatic column detection and preview

### ✅ Visual Column Management
- Select columns with intuitive checkboxes
- Apply powerful transformations with the enhanced formula editor
- Real-time formula validation and previews
- **Support for Python expressions via `pandas.eval()` with a safer `asteval` fallback** (ensures secure evaluation)

### ✅ Advanced Reference Data Joining
- Load and link secondary Excel files
- Define relationships with intuitive key mapping
- Selective column inclusion from reference data
- Left-join operation with clear visual feedback

### ✅ Smart Preset System
- Save frequently used transformations as presets
- Quick-apply column configurations
- Persistent workspace settings
- Streamlined workflow for repeated tasks

### ✅ Enhanced Export Capabilities
- One-click export to clean Excel files
- Visual progress tracking (where available)
- Optimized output files
- Comprehensive error handling

### ✨ Modern UI Improvements
- Professional, material design-inspired interface
- Responsive layout that adapts to window size
- Intuitive tooltips and help features
- Streamlined workflow with clear visual hierarchy
- Enhanced status updates with success/error indicators

---

## 🖥️ Technical Details

### Project Structure
matrix-ingestor/
├── main.py             # Entry point
├── gui/
│   ├── app.py          # Main application logic & UI coordination
│   ├── core/
│   │   ├── exporter.py       # Export functionality & formula application logic
│   │   ├── formula_engine.py # Formula preview logic
│   │   ├── loader.py         # Excel file loading
│   │   └── presets.py        # Preset management
│   └── ui/
│       ├── layout.py         # UI layout builder
│       ├── theme.py          # Theming system
│       ├── utils.py          # UI utilities (tooltips, messages, etc.)
│       ├── custom_widgets.py # Enhanced widgets (FormulaEntry, TooltipButton, etc.)
│       └── init.py       # Package initialization
├── assets/             # Optional: Application icons/images
├── presets.json        # Default location for saved transformation presets
├── config.json         # Default location for saved application configuration
├── .gitignore          # Specifies intentionally untracked files
└── README.md           # This file

*(Note: Add a `LICENSE` file if applicable)*

### Core Components
- **Formula Engine**: Leverages pandas' efficient `eval()` capabilities, **falling back to the secure `asteval` library for complex or row-wise formulas.**
- **Theme System**: Consistent, customizable visual styling using `tkinter.ttk`.
- **Custom Widgets**: Enhanced UI controls (like `TooltipButton`, `ProgressDialog`, and a potential `FormulaEntry`) for improved user experience.
- **Export Pipeline**: Handles data transformation, joining, and output generation to Excel format.

---

## 🛠️ Installation

### Prerequisites
- Python 3.7 or higher
- Required packages: `pandas`, `openpyxl`, `asteval`
- `tkinter` (usually included with standard Python distributions)

### Setup

1.  Clone the repository:
    ```bash
    # Replace 'yourusername' with the actual GitHub username/organization
    git clone [https://github.com/yourusername/matrix-ingestor.git](https://github.com/yourusername/matrix-ingestor.git)
    cd matrix-ingestor
    ```

2.  Install required packages:
    ```bash
    pip install pandas openpyxl asteval
    ```
    *(Consider creating a virtual environment first)*

3.  Run the application:
    ```bash
    python main.py
    ```

---

## 📝 Usage Guide

### Working with Excel Files
1.  Click "Load Excel File" to select your main file.
2.  Choose the appropriate sheet from the dropdown menu.
3.  Set the header row index if needed (default: 0, meaning the first row is the header).
4.  The columns will appear in the main panel for configuration.

### Transforming Data
-   **Include/Exclude Columns**: Toggle checkboxes next to column names.
-   **Apply Formulas**: Enter transformations in the formula field next to the column name.
    -   Reference other columns using their names directly within pandas/asteval compatible expressions: `Price * Quantity` or `[Column Name with Spaces]` (though direct names are usually better if no spaces).
    -   Example: `Price * 1.15` (applies 15% markup).
    -   Example: `FirstName.str.upper() + " " + LastName.str.upper()` (combines names in uppercase using pandas string methods - preferred over row-wise).
    -   Example: `"Yes" if Age > 18 else "No"` (conditional logic, likely evaluated by `asteval` row-wise if `pandas.eval` fails).
    -   **Note**: Formulas are primarily evaluated using pandas' efficient engine. More complex Python logic or functions not directly supported by pandas may use the `asteval` library for safe, row-by-row evaluation. Check the preview text for hints (`Preview (pandas): ...` or `Preview (asteval): ...`).

### Working with Reference Data
1.  Click "Load Reference File" to select a secondary Excel file.
2.  Enter the column name from your main file (after transformations) to join on in "Main Key".
3.  Enter the corresponding column name from your reference file in "Ref Key".
4.  List the columns you want to bring over from the reference file in "Ref Columns" (comma-separated).

### Using Presets & Config
1.  Configure your column selections and formulas as needed.
2.  Click "Save Config" to save the current setup (header row, selections, formulas) to `config.json`.
3.  Click "Load Config" to restore a previously saved setup.
4.  Use the "Column Presets" dropdown (populated from `presets.json`) and "Apply Preset" for quick access to predefined column/formula sets.

### Exporting Transformed Data
1.  Ensure all desired columns are selected, formulas are correct, and reference joins are configured.
2.  Click "EXPORT DATA".
3.  Choose a name and location to save the output `.xlsx` file.
4.  The processed data will be exported.

---

## 🤝 Contributing

Contributions are welcome! Please feel free to submit issues or pull requests to help improve Matrix Ingestor.

## 📄 License

This project is licensed under the MIT License - see the `LICENSE` file for details. *(