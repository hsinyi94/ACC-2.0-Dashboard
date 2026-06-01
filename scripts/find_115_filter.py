"""找出如何從『最終錄取115位 + 候補18位』取到真正的 115 位。"""
import pandas as pd
from pathlib import Path

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

df = pd.read_excel(ACC_FILE, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
print(f"總筆數: {len(df)}")
print(f"欄位: {list(df.columns)}")
print()

# 檢查各欄位的分布 (非數字/ID 欄)
for col in df.columns:
    vals = df[col]
    nunique = vals.nunique(dropna=False)
    if nunique <= 10:
        print(f"\n-- {col} (唯一值 {nunique}) --")
        print(vals.value_counts(dropna=False))

# 特別看 '保留' 欄
print("\n=== '保留' 欄位分布 ===")
print(df["保留"].value_counts(dropna=False))
print(f"保留='不保留' 的筆數: {(df['保留'] == '不保留').sum()}")
print(f"保留='保留' 的筆數: {(df['保留'] == '保留').sum()}")

# '是否須被替換？'
print("\n=== '是否須被替換？' ===")
print(df["是否須被替換？"].value_counts(dropna=False))

# 也看一下 No. 欄 (錄取 115 位通常是 No.1-115 + 候補是之後)
print("\n=== No. 欄分布 (前 5 + 中間 + 尾端 5) ===")
print(df["No."].head())
print("...")
print(df["No."].tail(25))
