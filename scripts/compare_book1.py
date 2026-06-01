"""對比 Book1 的 330 筆資料與我抓到的 359 筆,找出差異。"""
import pandas as pd
from pathlib import Path

BOOK1 = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\Book1.xlsx")
P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR\3. Mar\P0_20260331.xlsx")
ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

# 先看 Book1 結構
xls = pd.ExcelFile(BOOK1, engine="openpyxl")
print(f"Book1 工作表: {xls.sheet_names}")
for sh in xls.sheet_names:
    df = pd.read_excel(BOOK1, sheet_name=sh, engine="openpyxl")
    print(f"\n-- {sh} ({df.shape}) --")
    print(f"欄位: {list(df.columns)}")
    print("\n前 5 列:")
    print(df.head().to_string(max_cols=15, max_colwidth=20))
    if "mtd_ord_gms" in df.columns:
        for m in [1, 2, 3]:
            if "calendar_month" in df.columns:
                s = df[df["calendar_month"] == m]["mtd_ord_gms"].sum()
                n = len(df[df["calendar_month"] == m])
                print(f"  month={m}: GMS={s:,.2f}, rows={n}")
