"""看 2026 Raw 的 MCID 型別、筆數、以及 6 個 adoption 欄位的值分布。"""
import pandas as pd
from pathlib import Path

F = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wk21\NSR Launch Tracker_20260523.xlsx")
COLS = ["merchant_customer_id", "launch_channel",
        "is_pl", "is_fba_adopt_by_seller", "is_sp_adopt_by_seller",
        "is_deal_adopt_by_seller", "is_brand_rep_by_seller", "is_b2b_adopt_by_seller"]

df = pd.read_excel(F, sheet_name="2026 Raw", engine="openpyxl", usecols=COLS)
print(f"筆數: {len(df)}")
print(f"唯一 MCID: {df['merchant_customer_id'].nunique()}")
print(f"MCID dtype: {df['merchant_customer_id'].dtype}")
print(f"MCID 前 5: {df['merchant_customer_id'].head().tolist()}")
print(f"\nlaunch_channel 分布:")
print(df["launch_channel"].value_counts(dropna=False))

for col in ["is_pl", "is_fba_adopt_by_seller", "is_sp_adopt_by_seller",
            "is_deal_adopt_by_seller", "is_brand_rep_by_seller", "is_b2b_adopt_by_seller"]:
    print(f"\n{col} 分布:")
    print(df[col].value_counts(dropna=False))
