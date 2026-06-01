"""看 1.0 最終錄取名單 站點 欄的原始值樣貌。"""
import pandas as pd
from pathlib import Path

FPATH = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")
df = pd.read_excel(FPATH, sheet_name="最終錄取名單", engine="openpyxl")
df = df[df["公司名稱"].notna()]

print(f"總筆數: {len(df)}")
print(f"站點空值: {df['站點'].isna().sum()}")
print()
print("站點 value_counts (原始):")
for v, c in df["站點"].value_counts(dropna=False).items():
    print(f"  {c:>3}  {v!r}")
