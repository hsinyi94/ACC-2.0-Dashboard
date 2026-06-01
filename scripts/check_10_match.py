"""檢查兩個 2025 檔案各能 match 到幾位 ACC 1.0 MCID。"""
import pandas as pd
from pathlib import Path

ACC_10_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")

# 讀 ACC 1.0 MCID
df_10 = pd.read_excel(ACC_10_FILE, sheet_name="GMS by seller", engine="openpyxl")
mcids_10 = set(df_10["MCID"].dropna().astype(str).str.strip())
print(f"ACC 1.0 MCID 數: {len(mcids_10)}")
print(f"前 5: {list(mcids_10)[:5]}")

# 檔案 1: wk45 NSR Launch Tracker
f1 = Path(r"W:\Team Spaces\TWGS\00_Commonly Accessed\2025 Weekly Business Report\wk45\NSR Launch Tracker_20251108.xlsx")
print(f"\n=== 檔案 1: {f1.name} ===")
print(f"存在: {f1.exists()}")
if f1.exists():
    xls = pd.ExcelFile(f1, engine="openpyxl")
    print(f"工作表: {xls.sheet_names}")
    # 嘗試 2025 Raw
    for sh in ["2025 Raw", "2025 raw", "Raw"]:
        if sh in xls.sheet_names:
            df = pd.read_excel(f1, sheet_name=sh, engine="openpyxl",
                               usecols=["merchant_customer_id"])
            df = df.dropna(subset=["merchant_customer_id"])
            df["merchant_customer_id"] = df["merchant_customer_id"].apply(
                lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x).strip()
            )
            all_mcids = set(df["merchant_customer_id"])
            matched = mcids_10 & all_mcids
            print(f"  工作表 '{sh}': 總筆數={len(df)}, 唯一 MCID={len(all_mcids)}, matched 1.0={len(matched)}")
            break
    else:
        print("  找不到 2025 Raw 工作表")

# 檔案 2: wk35 WBR Page0
f2 = Path(r"W:\Team Spaces\TWGS\00_Commonly Accessed\2025 Weekly Business Report\wk35\WBR Page0 W35_v2.xlsx")
print(f"\n=== 檔案 2: {f2.name} ===")
print(f"存在: {f2.exists()}")
if f2.exists():
    xls2 = pd.ExcelFile(f2, engine="openpyxl")
    print(f"工作表: {xls2.sheet_names}")
    # 看哪個工作表有 merchant_customer_id
    for sh in xls2.sheet_names:
        try:
            df = pd.read_excel(f2, sheet_name=sh, engine="openpyxl", nrows=2)
            if "merchant_customer_id" in df.columns:
                df_full = pd.read_excel(f2, sheet_name=sh, engine="openpyxl",
                                        usecols=["merchant_customer_id"])
                df_full = df_full.dropna(subset=["merchant_customer_id"])
                df_full["merchant_customer_id"] = df_full["merchant_customer_id"].apply(
                    lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x).strip()
                )
                all_mcids = set(df_full["merchant_customer_id"])
                matched = mcids_10 & all_mcids
                print(f"  工作表 '{sh}': 總筆數={len(df_full)}, 唯一 MCID={len(all_mcids)}, matched 1.0={len(matched)}")
        except Exception as e:
            pass
