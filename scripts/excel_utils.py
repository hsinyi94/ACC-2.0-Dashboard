"""Shared Excel workbook helpers."""
from pathlib import Path
from collections.abc import Iterable

import pandas as pd


def find_sheet_with_columns(path: Path, required_columns: Iterable[str]) -> str:
    """Return the first worksheet containing every required column."""
    required = set(required_columns)
    workbook = pd.ExcelFile(path, engine="openpyxl")
    preferred = ["raw", "Raw", "Sheet1"]
    candidates = [name for name in preferred if name in workbook.sheet_names]
    candidates.extend(name for name in workbook.sheet_names if name not in candidates)

    for sheet in candidates:
        header = pd.read_excel(
            path, sheet_name=sheet, engine="openpyxl", nrows=0
        )
        if required.issubset(header.columns):
            return sheet

    missing = ", ".join(sorted(required))
    raise ValueError(f"{path.name} 找不到包含必要欄位的工作表: {missing}")
