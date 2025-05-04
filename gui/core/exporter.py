import pandas as pd
from tkinter import filedialog, messagebox

def export_to_excel(df, column_vars, formula_vars, ref_df, main_key, ref_key, ref_cols_input, file_path):
    selected_cols = [col for col, var in column_vars.items() if var.get()]
    if not selected_cols:
        messagebox.showwarning("No columns selected", "Select at least one column.")
        return

    result = pd.DataFrame()
    for col in selected_cols:
        formula = formula_vars[col].get().strip()
        if not formula:
            result[col] = df[col]
            continue
        try:
            result[col] = df.eval(formula)
        except:
            try:
                result[col] = df.apply(lambda row: eval(formula, {}, row.to_dict()), axis=1)
            except Exception as e:
                result[col] = f"#ERR {e}"

    if ref_df is not None:
        ref_cols = [c.strip() for c in ref_cols_input.split(',') if c.strip()]
        if main_key and ref_key and ref_cols:
            try:
                ref_slice = ref_df[[ref_key] + ref_cols]
                result = result.merge(ref_slice, left_on=main_key, right_on=ref_key, how="left")
            except Exception as e:
                messagebox.showerror("Join Error", str(e))

    out_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
    if out_path:
        result.to_excel(out_path, index=False)
        messagebox.showinfo("Done", f"Saved to {out_path}")
