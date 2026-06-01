"""把 2.0分析 工作表整張倒出來,找出所有分析項目。"""
import pandas as pd
from pathlib import Path

BASE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0")
FPATH = BASE / "ACC 2.0 最終錄取結果 (20260105).xlsx"

df = pd.read_excel(FPATH, sheet_name="2.0分析", header=None, engine="openpyxl")
print(f"形狀: {df.shape}")
print()

# 把所有非空儲存格,依 (row, col) 印出來
pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 200)
pd.set_option("display.max_colwidth", 40)

# 格式化輸出:每列顯示 row_idx | col0 | col1 | ...
for i, row in df.iterrows():
    cells = []
    for j, v in enumerate(row):
        if pd.notna(v):
            s = str(v)
            if len(s) > 30:
                s = s[:27] + "..."
            cells.append(f"[{j}]{s}")
    if cells:
        print(f"R{i:02d}: " + " | ".join(cells))
