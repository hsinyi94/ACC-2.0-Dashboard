"""從 ACC 1.0 分析.xlsx 找 ACC 1.0 2025 全年 GMS 總額。"""
import pandas as pd
from pathlib import Path

F = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")

# 1. 分析工作表 (之前看到 ACC 1.0 YTD GMS = 3.6MM、DSR total GMS = 26,960,190)
print("=" * 60)
print("[分析] 工作表全欄位")
print("=" * 60)
df = pd.read_excel(F, sheet_name="分析", header=None, engine="openpyxl")
for i, row in df.iterrows():
    cells = [f"[{j}]{v}" for j, v in enumerate(row) if pd.notna(v)]
    if cells:
        print(f"R{i:02d}: " + " | ".join(cells))

print()
print("=" * 60)
print("[GMS by seller] 頭尾 5 列與欄位")
print("=" * 60)
df2 = pd.read_excel(F, sheet_name="GMS by seller", engine="openpyxl")
print(f"形狀: {df2.shape}, 欄位: {list(df2.columns)[:20]}")
print("\n前 5 列:")
print(df2.head().to_string(max_cols=10, max_colwidth=25))

# 看有無 '2025' 關鍵字欄位
gms_cols = [c for c in df2.columns if any(k in str(c) for k in ["2025", "GMS", "Total", "YTD"])]
print(f"\nGMS/2025 相關欄位: {gms_cols}")

# 若有單一 GMS Total 欄就總和
for c in gms_cols:
    try:
        s = pd.to_numeric(df2[c], errors="coerce").sum()
        print(f"  sum({c}) = {s:,.2f}")
    except Exception:
        pass

print()
print("=" * 60)
print("[2025 Dec. DSR] 頭尾")
print("=" * 60)
df3 = pd.read_excel(F, sheet_name="2025 Dec. DSR", engine="openpyxl")
print(f"形狀: {df3.shape}, 欄位: {list(df3.columns)[:20]}")
print("\n前 5 列:")
print(df3.head().to_string(max_cols=10, max_colwidth=25))
