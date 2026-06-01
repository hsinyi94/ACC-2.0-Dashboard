"""只確認能不能讀到兩個 Excel 檔,列出工作表名稱與大小。不做任何分析。"""
import pandas as pd
from pathlib import Path

BASE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0")
FILES = [
    "ACC 1.0分析.xlsx",
    "ACC 2.0 最終錄取結果 (20260105).xlsx",
]

for fname in FILES:
    fpath = BASE / fname
    print(f"[檔案] {fname}")
    print(f"  路徑存在: {fpath.exists()}")
    if not fpath.exists():
        continue
    print(f"  大小: {fpath.stat().st_size / 1024:.1f} KB")
    try:
        xls = pd.ExcelFile(fpath, engine="openpyxl")
        print(f"  工作表 ({len(xls.sheet_names)} 個): {xls.sheet_names}")
    except Exception as e:
        print(f"  讀取失敗: {e}")
    print()
