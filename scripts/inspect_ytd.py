"""檢視 115 seller performance summary 的 P2:Z5 區域,以及 WBR 最新一週 P0 結構。"""
import pandas as pd
from pathlib import Path
import re

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
WBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")

print("=" * 80)
print("1. 115 seller performance summary — P2:Z5 區域")
print("=" * 80)
# P=col 15 (0-index), Z=col 25, rows 2-5 -> 0-index 1..4
df = pd.read_excel(ACC_FILE, sheet_name="115 seller performance summary",
                   header=None, engine="openpyxl")
print(f"工作表形狀: {df.shape}")
sub = df.iloc[0:8, 15:26]
print("\n--- 前 8 列 × P-Z 欄 ---")
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 240)
pd.set_option("display.max_colwidth", 25)
print(sub.to_string())

print()
print("=" * 80)
print("2. WBR 最新一週資料夾")
print("=" * 80)
pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
latest = None
for child in WBR_BASE.iterdir():
    if not child.is_dir():
        continue
    m = pattern.match(child.name)
    if m:
        n = int(m.group(1))
        if latest is None or n > latest[0]:
            latest = (n, child)
print(f"最新週次資料夾: {latest[1].name if latest else '無'}")

if latest:
    p0_files = sorted(
        [f for f in latest[1].iterdir()
         if f.is_file() and f.suffix.lower() == ".xlsx"
         and f.name.lower().startswith("p0")],
        key=lambda f: f.stat().st_mtime, reverse=True
    )
    print(f"P0 檔: {p0_files[0].name}" if p0_files else "無 P0 檔")

    print("\n=== P0 工作表與欄位 ===")
    p0 = p0_files[0]
    xls = pd.ExcelFile(p0, engine="openpyxl")
    print(f"工作表: {xls.sheet_names}")
    for sh in xls.sheet_names:
        df_h = pd.read_excel(p0, sheet_name=sh, engine="openpyxl", nrows=2)
        print(f"\n-- {sh} ({df_h.shape[1]} 欄)")
        if df_h.shape[1] > 30:
            # 找關鍵字欄位
            cols = list(df_h.columns)
            key_cols = [c for c in cols if any(k in str(c).lower()
                for k in ["year", "week", "month", "launch", "merchant", "ytd", "mtd"])]
            print(f"   關鍵欄位: {key_cols}")
        else:
            print(f"   欄位: {list(df_h.columns)}")
