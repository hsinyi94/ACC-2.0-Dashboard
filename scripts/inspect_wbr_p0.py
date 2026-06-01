"""探索 WBR 最新週 P0 檔的 reporting_year / reporting_week_of_year 分布,
確認能否同時取 2025 和 2026 的 YTD 資料。"""
import pandas as pd
from pathlib import Path

P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wk17\P0_20260425.xlsx")
print(f"檔案: {P0.name}, 大小: {P0.stat().st_size / 1024 / 1024:.1f} MB")

# 只讀必要欄位
cols = ["reporting_year", "reporting_week_of_year", "launch_channel",
        "merchant_customer_id", "ytd_ord_gms"]
print(f"只讀欄位: {cols}")
df = pd.read_excel(P0, sheet_name="raw", engine="openpyxl", usecols=cols)
print(f"原始筆數: {len(df):,}")
print(f"\nreporting_year 分布:")
print(df["reporting_year"].value_counts(dropna=False).sort_index())
print(f"\nreporting_week_of_year 分布 (全部):")
print(df["reporting_week_of_year"].value_counts(dropna=False).sort_index())
print(f"\nlaunch_channel 分布:")
print(df["launch_channel"].value_counts(dropna=False))

# 針對 launch_channel=DSR + each year,看 week max
print("\n=== 每個 year + DSR 的 week 最大值與筆數 ===")
for yr in df["reporting_year"].dropna().unique():
    sub = df[(df["reporting_year"] == yr) & (df["launch_channel"] == "DSR")]
    if len(sub) == 0:
        print(f"year={int(yr)}: 無 DSR 資料")
        continue
    max_wk = sub["reporting_week_of_year"].max()
    latest = sub[sub["reporting_week_of_year"] == max_wk]
    print(f"year={int(yr)}: week max={int(max_wk)}, 筆數={len(latest):,}, 該週 ytd_ord_gms sum={latest['ytd_ord_gms'].sum():,.2f}")
