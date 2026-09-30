"""Shared Excel workbook helpers."""
from pathlib import Path
from collections.abc import Iterable

import pandas as pd

# P0 有時是 xlsx,有時是資料庫直接匯出的 csv/tsv
P0_SUFFIXES = {".xlsx", ".csv", ".txt", ".tsv"}
_TEXT_SUFFIXES = {".csv", ".txt", ".tsv"}


def is_p0_file(path: Path, exclude_ab_data: bool = True) -> bool:
    """P0 開頭且格式支援;預設排除 ab_data 附檔。"""
    if not (path.is_file()
            and path.suffix.lower() in P0_SUFFIXES
            and path.name.lower().startswith("p0")):
        return False
    return not (exclude_ab_data and "ab_data" in path.name.lower())


def read_p0(path: Path, columns: Iterable[str]) -> pd.DataFrame:
    """讀 P0 檔的指定欄位,自動處理 xlsx / csv / tsv。"""
    columns = list(columns)
    suffix = path.suffix.lower()
    if suffix in _TEXT_SUFFIXES:
        separator = "\t" if suffix in (".txt", ".tsv") else ","
        # utf-8-sig:部分匯出檔帶 BOM,會讓第一個欄位名對不上
        return pd.read_csv(path, sep=separator, usecols=columns, encoding="utf-8-sig")
    sheet = find_sheet_with_columns(path, columns)
    return pd.read_excel(path, sheet_name=sheet, engine="openpyxl", usecols=columns)


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
