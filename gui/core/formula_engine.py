import numpy as np
import pandas as pd

def evaluate_formula(df, col, formula):
    if not formula:
        return ""
    try:
        value = df.eval(formula).dropna().iloc[0]
        return f"Preview: {value}"
    except Exception:
        try:
            result = df.apply(lambda row: eval(formula, {}, row.to_dict()), axis=1)
            preview = result.dropna().iloc[0]
            return f"Preview: {preview}"
        except:
            return "Error"
