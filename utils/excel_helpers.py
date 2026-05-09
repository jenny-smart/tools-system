from __future__ import annotations

from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string


def safe_excel_value(value):
    """Convert pandas / numpy empty values to a clean Excel blank."""
    try:
        import pandas as pd

        if pd.isna(value):
            return None
    except Exception:
        pass
    return value


def write_dataframe_to_sheet(ws, df, start_row: int = 2, start_col: int = 1) -> None:
    for r_idx, row in enumerate(df.itertuples(index=False, name=None), start=start_row):
        for c_idx, value in enumerate(row, start=start_col):
            ws.cell(r_idx, c_idx, safe_excel_value(value))


def fill_formula_down(ws, column_letter: str, formula_in_row_2: str, row_count: int) -> None:
    if row_count <= 0:
        return

    column_letter = str(column_letter).strip().upper()
    formula = str(formula_in_row_2).strip()
    if not formula:
        return
    if not formula.startswith("="):
        formula = "=" + formula

    col_num = column_index_from_string(column_letter)
    origin = f"{column_letter}2"

    for excel_row in range(2, row_count + 2):
        target = f"{column_letter}{excel_row}"
        ws.cell(excel_row, col_num, Translator(formula, origin=origin).translate_formula(target))
