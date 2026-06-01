"""檢視 ACC 2.0 資料夾中兩個 Excel 檔的結構。"""
import pandas as pd
from pathlib import Path

BASE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0")
FILES = [
    "ACC 1.0分析.xlsx",
    "ACC 2.0 最終錄取結果 (20260105).xlsx",
]

for fname in FILES:
    fpath = BASE / fname
    print("=" * 80)
    print(f"檔案: {fname}")
    print("=" * 80)
    try:
        xls = pd.ExcelFile(fpath, engine="openpyxl")
        print(f"工作表數量: {len(xls.sheet_names)}")
        for sheet in xls.sheet_names:
            df = pd.read_excel(fpath, sheet_name=sheet, engine="openpyxl")
            print(f"\n-- 工作表: {sheet}")
            print(f"   形狀: {df.shape[0]} rows x {df.shape[1]} cols")
            print(f"   欄位: {list(df.columns)}")
            # 顯示前 3 列(截斷避免太長)
            preview = df.head(3).to_string(max_cols=8, max_colwidth=20)
            print("   預覽:")
            for line in preview.split("\n"):
                print(f"     {line}")
    except Exception as e:
        print(f"讀取失敗: {e}")
    print()
