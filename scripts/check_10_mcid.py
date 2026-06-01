"""確認 ACC 1.0 的 MCID 來源。"""
import pandas as pd
from pathlib import Path

F = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")
df = pd.read_excel(F, sheet_name="GMS by seller", engine="openpyxl")
print(f"形狀: {df.shape}")
print(f"MCID 欄位: {'MCID' in df.columns}")
if "MCID" in df.columns:
    print(f"MCID 非空: {df['MCID'].notna().sum()}")
    print(f"MCID dtype: {df['MCID'].dtype}")
    print(f"前 5: {df['MCID'].head().tolist()}")
