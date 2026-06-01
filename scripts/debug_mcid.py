"""Debug:為何 WBR P0 的 MCID 對不上 ACC 2.0 的 MCID。"""
import pandas as pd
from pathlib import Path

P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wk17\P0_20260425.xlsx")
ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

acc_df = pd.read_excel(ACC, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
acc_df = acc_df[acc_df["No."].notna()]
print(f"ACC MCID dtype: {acc_df['MCID'].dtype}")
print(f"ACC MCID 前 5: {acc_df['MCID'].head().tolist()}")
print(f"ACC MCID 型別 per 值: {set(type(v).__name__ for v in acc_df['MCID'].head())}")
print(f"ACC MCID 字串化前 5: {acc_df['MCID'].astype(str).head().tolist()}")

cols = ["reporting_year", "reporting_week_of_year", "launch_channel",
        "merchant_customer_id", "ytd_ord_gms"]
p0 = pd.read_excel(P0, sheet_name="raw", engine="openpyxl", usecols=cols)
p0 = p0[(p0["launch_channel"] == "DSR") & (p0["reporting_year"] == 2026)
        & (p0["reporting_week_of_year"] == 17)]
print(f"\nP0 latest merchant_customer_id dtype: {p0['merchant_customer_id'].dtype}")
print(f"P0 MCID 前 5: {p0['merchant_customer_id'].head().tolist()}")
print(f"P0 MCID 字串化前 5: {p0['merchant_customer_id'].astype(str).head().tolist()}")

# 交集試
set_acc_str = set(acc_df["MCID"].astype(str).str.strip())
set_p0_str = set(p0["merchant_customer_id"].astype(str).str.strip())
print(f"\nACC 集合大小: {len(set_acc_str)}")
print(f"P0 集合大小: {len(set_p0_str)}")
print(f"交集: {len(set_acc_str & set_p0_str)}")

# 試試把 .0 拿掉 (pandas float -> str 會變 '1.23e+12' 或 '1.23E+12' 或純字串)
sample_acc = list(set_acc_str)[:3]
sample_p0 = list(set_p0_str)[:3]
print(f"\nACC 樣本: {sample_acc}")
print(f"P0 樣本: {sample_p0}")
