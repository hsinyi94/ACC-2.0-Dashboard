"""檢視新的 ACC 2.0 最終賣家名單 (110) 工作表結構。"""
import pandas as pd
from pathlib import Path

ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

df = pd.read_excel(ACC, sheet_name="最終賣家名單 (110)", engine="openpyxl")
print(f"形狀: {df.shape}")
print(f"欄位 ({len(df.columns)}):")
for i, c in enumerate(df.columns):
    print(f"  [{i}] {c!r}")

print(f"\n前 5 列:")
print(df.head().to_string(max_cols=12, max_colwidth=25))

# MCID 確認
mcid_col = None
for c in df.columns:
    if "MCID" in str(c).upper():
        mcid_col = c
        break

if mcid_col:
    print(f"\nMCID 欄位名: {mcid_col!r}")
    print(f"MCID 非空: {df[mcid_col].notna().sum()}")
    print(f"MCID 唯一: {df[mcid_col].nunique()}")
    print(f"MCID 前 5: {df[mcid_col].head().tolist()}")
    print(f"MCID dtype: {df[mcid_col].dtype}")
else:
    print("\n找不到 MCID 欄位!")
