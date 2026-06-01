"""驗證 No. 欄非空 = 115 筆,且 MCID 可用。"""
import pandas as pd
from pathlib import Path

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
df = pd.read_excel(ACC_FILE, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")

recruited = df[df["No."].notna()].copy()
print(f"No. 非空筆數: {len(recruited)}")
print(f"其中 MCID 非空: {recruited['MCID'].notna().sum()}")
print(f"MCID 空 的筆: {recruited['MCID'].isna().sum()}")

# MCID 型別
print(f"\nMCID 型別: {recruited['MCID'].dtype}")
print(f"前 5 個: {recruited['MCID'].head().tolist()}")

# 用 Jan/Feb GMS 橫比對 P24:Y29 的 82798.99 / 202450.69
print(f"\nJan. GMS 合計: {recruited['Jan. GMS'].sum():.2f}")
print(f"Feb. GMS 合計: {recruited['Feb. GMS'].sum():.2f}")
print("(P24:Y29 顯示 82798.99 / 202450.69)")
