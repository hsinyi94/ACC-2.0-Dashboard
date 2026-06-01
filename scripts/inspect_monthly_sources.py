"""檢視月度分析所需的三個資料來源結構。"""
import pandas as pd
from pathlib import Path
import re

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
MBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR")

print("=" * 80)
print("1. 115 seller performance summary — P24:Y29 區域")
print("=" * 80)
# P = 第 16 欄 (0-index 15), Y = 第 25 欄 (0-index 24)
# 列 24-29 (0-index 23-28)
df_sum = pd.read_excel(ACC_FILE, sheet_name="115 seller performance summary",
                       header=None, engine="openpyxl")
print(f"工作表形狀: {df_sum.shape}")
# P24:Y29 即 col 15..24, row 23..28
sub = df_sum.iloc[23:29, 15:25]
print("\n--- P24:Y29 內容 ---")
print(sub.to_string())
print()

# 同時看一下 P 欄前後幾列,找到這塊的標題
print("\n--- P20:Y30 (前後幾列) 供對照 ---")
print(df_sum.iloc[19:30, 15:25].to_string(max_colwidth=25))

print()
print("=" * 80)
print("2. 115 seller performance summary — MCID 欄位")
print("=" * 80)
df_sum_head = pd.read_excel(ACC_FILE, sheet_name="115 seller performance summary", engine="openpyxl")
print(f"頂層形狀: {df_sum_head.shape}")
print(f"欄位前 15: {list(df_sum_head.columns[:15])}")

print()
print("=" * 80)
print("3. ACC 2.0 錄取結果 — MCID 位置 (最終錄取115位 + 候補18位)")
print("=" * 80)
df_r = pd.read_excel(ACC_FILE, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
print(f"形狀: {df_r.shape}")
print(f"欄位: {list(df_r.columns)}")
# 查看 MCID 的樣本
if "MCID" in df_r.columns:
    print("\nMCID 前 5 個樣本:")
    print(df_r["MCID"].head(10).tolist())
    print(f"MCID 非空數量: {df_r['MCID'].notna().sum()}")
if "錄取國家" in df_r.columns:
    print("\n錄取國家 分布:")
    print(df_r["錄取國家"].value_counts(dropna=False))

print()
print("=" * 80)
print("4. P0_20260331.xlsx 結構")
print("=" * 80)
p0 = MBR_BASE / "3. Mar" / "P0_20260331.xlsx"
print(f"檔案: {p0}")
print(f"存在: {p0.exists()}, 大小: {p0.stat().st_size / 1024 / 1024:.1f} MB")

xls = pd.ExcelFile(p0, engine="openpyxl")
print(f"工作表: {xls.sheet_names}")

for sh in xls.sheet_names:
    df_h = pd.read_excel(p0, sheet_name=sh, engine="openpyxl", nrows=3)
    print(f"\n-- {sh} (前 3 列,{df_h.shape[1]} 欄)")
    print(f"   欄位: {list(df_h.columns)}")
