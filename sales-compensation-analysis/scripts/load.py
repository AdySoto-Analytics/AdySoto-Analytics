"""Load the original case-study workbook into pandas DataFrames (read-only)."""
import pandas as pd
from pathlib import Path
SRC = Path(__file__).resolve().parent.parent / "data" / "Sales_Compensation_Analyst_Case_Study_ORIGINAL.xlsx"
def load():
    x = pd.read_excel(SRC, sheet_name=None)
    out = {}
    for k, df in x.items():
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        df.insert(0, "src_row", df.index + 2)  # Excel row number for traceability
        out[k] = df
    return out
