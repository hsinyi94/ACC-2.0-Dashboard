"""把 115 seller performance summary 上半部 (row 0~22) 全欄位倒出來,
看 P2:Z5 周邊有沒有 NSR 的 2025 / Goal 資料。"""
import pandas as pd
from pathlib import Path

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
df = pd.read_excel(ACC_FILE, sheet_name="115 seller performance summary",
                   header=None, engine="openpyxl")
print(f"形狀: {df.shape}")

# 列出 row 0..22 的非空儲存格
for i in range(0, 23):
    cells = []
    for j in range(df.shape[1]):
        v = df.iat[i, j]
        if pd.notna(v):
            s = str(v)
            if len(s) > 30:
                s = s[:27] + "..."
            cells.append(f"[{j}]{s}")
    if cells:
        print(f"R{i:02d}: " + " | ".join(cells))
