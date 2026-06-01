"""確認 ACC 1.0 2025 全年 GMS 精確數字。"""
import pandas as pd
from pathlib import Path

F = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")

# 讀「分析」工作表,看 raw cell,看 R07 col 4 實際值
df = pd.read_excel(F, sheet_name="分析", header=None, engine="openpyxl")
print("R07 實際值:")
for j in range(df.shape[1]):
    v = df.iat[7, j]
    print(f"  col {j}: {v!r}")

# 也看 R11 'GMS Goal' = 8MM
print("\nR11 實際值:")
for j in range(df.shape[1]):
    v = df.iat[11, j]
    print(f"  col {j}: {v!r}")

# GMS by seller 的 Dec. YTD GMS 明顯是數字,看它的實際分布
df2 = pd.read_excel(F, sheet_name="GMS by seller", engine="openpyxl")
print(f"\nGMS by seller 形狀: {df2.shape}")
print(f"公司數 (有 MCID 的): {df2['MCID'].notna().sum()}")
print(f"Aug. YTD GMS sum: {pd.to_numeric(df2['Aug. YTD GMS'], errors='coerce').sum():,.2f}")
print(f"Dec. YTD GMS sum: {pd.to_numeric(df2['Dec. YTD GMS'], errors='coerce').sum():,.2f}")
print()
# 分開看 — 是否有重複 MCID
mcid_counts = df2['MCID'].value_counts()
print(f"重複 MCID 數: {(mcid_counts > 1).sum()}")

# 看前 10 筆 Dec. YTD GMS
print("\n前 10 筆 Dec. YTD GMS:")
for i, row in df2.head(10).iterrows():
    print(f"  {row['公司名稱']}: MCID={row['MCID']}, Dec. YTD GMS={row['Dec. YTD GMS']}")
