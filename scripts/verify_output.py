"""快速檢查輸出檔的關鍵數字。"""
import pandas as pd
from pathlib import Path

OUT = Path(__file__).resolve().parent.parent / "output" / f"ACC_2.0_分析_20260430.xlsx"
df = pd.read_excel(OUT, sheet_name="ACC 2.0 分析", header=None, engine="openpyxl")

# 列出所有非空儲存格
for i, row in df.iterrows():
    cells = []
    for j, v in enumerate(row):
        if pd.notna(v):
            s = str(v)
            if len(s) > 40:
                s = s[:37] + "..."
            cells.append(f"[{j}]{s}")
    if cells:
        print(f"R{i:03d}: " + " | ".join(cells))
