"""檢視三個分析需要的工作表結構。只看欄位與少量樣本,不做計算。"""
import pandas as pd
from pathlib import Path

BASE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0")

TARGETS = [
    ("ACC 2.0 最終錄取結果 (20260105).xlsx", "最終錄取115位 + 候補18位"),
    ("ACC 2.0 最終錄取結果 (20260105).xlsx", "2.0分析"),
    ("ACC 2.0 最終錄取結果 (20260105).xlsx", "1.0分析"),
    ("ACC 1.0分析.xlsx", "最終錄取名單"),
    ("ACC 1.0分析.xlsx", "分析"),
]

for fname, sheet in TARGETS:
    fpath = BASE / fname
    print("=" * 90)
    print(f"[檔案] {fname}")
    print(f"[工作表] {sheet}")
    print("=" * 90)
    try:
        # 讀原始 header (不自動推斷) 看最上面幾列,因為 Excel 常有合併標題列
        raw = pd.read_excel(fpath, sheet_name=sheet, header=None, nrows=6, engine="openpyxl")
        print(f"原始前 6 列 (無 header):")
        print(raw.to_string(max_cols=15, max_colwidth=25))
        print()

        # 再用預設 header=0 看標準讀法下的欄位與形狀
        df = pd.read_excel(fpath, sheet_name=sheet, engine="openpyxl")
        print(f"預設讀法形狀: {df.shape[0]} rows x {df.shape[1]} cols")
        print(f"欄位 ({len(df.columns)}):")
        for i, c in enumerate(df.columns):
            print(f"  [{i}] {c!r}")
    except Exception as e:
        print(f"讀取失敗: {e}")
    print()
