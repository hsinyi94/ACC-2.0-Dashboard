"""檢視 NSR Launch Tracker 的 2026 raw 工作表結構。"""
import pandas as pd
from pathlib import Path

F = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wk21\NSR Launch Tracker_20260523.xlsx")
print(f"檔案: {F.name}, 大小: {F.stat().st_size / 1024 / 1024:.1f} MB")

xls = pd.ExcelFile(F, engine="openpyxl")
print(f"工作表: {xls.sheet_names}")

df = pd.read_excel(F, sheet_name="2026 Raw", engine="openpyxl", nrows=5)
print(f"\n-- 2026 raw ({df.shape[1]} 欄) --")
print(f"欄位: {list(df.columns)}")

# 找目標欄位
targets = ["merchant_customer_id", "is_pl", "is_fba_adopt_by_seller",
           "is_sp_adopt_by_seller", "is_deal_adopt_by_seller",
           "is_brand_rep_by_seller", "is_b2b_adopt_by_seller",
           "launch_channel"]
found = [c for c in df.columns if c in targets]
missing = [c for c in targets if c not in df.columns]
print(f"\n找到的目標欄位: {found}")
print(f"缺少的目標欄位: {missing}")

# 看前 5 列的目標欄位值
if found:
    print(f"\n前 5 列 (目標欄位):")
    print(df[found].to_string())

# 看總筆數
df_full = pd.read_excel(F, sheet_name="2026 Raw", engine="openpyxl",
                        usecols=found[:2] if found else None)
print(f"\n總筆數: {len(df_full)}")
