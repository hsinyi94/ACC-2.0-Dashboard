"""Debug:為何 2.0 月度數字比期望少。"""
import pandas as pd
from pathlib import Path

P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR\3. Mar\P0_20260331.xlsx")
ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

# 拿 115 MCID
acc_df = pd.read_excel(ACC, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
acc_df = acc_df[acc_df["No."].notna()]
mcids = set(acc_df["MCID"].astype(str).str.strip())
print(f"2.0 MCID 數: {len(mcids)}")

# 讀 P0 Sheet1,多拉欄位看分布
cols = ["cal_type", "calendar_year", "calendar_month", "launch_channel",
        "marketplace_id", "merchant_customer_id", "mtd_ord_gms"]
p0 = pd.read_excel(P0, sheet_name="Sheet1", engine="openpyxl", usecols=cols)
print(f"P0 原始筆數: {len(p0):,}")

# 字串化 MCID
p0 = p0.dropna(subset=["merchant_customer_id"]).copy()
p0["merchant_customer_id"] = p0["merchant_customer_id"].astype("int64").astype(str).str.strip()
p0["mtd_ord_gms"] = pd.to_numeric(p0["mtd_ord_gms"], errors="coerce").fillna(0.0)

# 只篩 2026
p0 = p0[p0["calendar_year"] == 2026]
print(f"年=2026 篩後: {len(p0):,}")

# 篩 MCID 在 115
p0_115 = p0[p0["merchant_customer_id"].isin(mcids)]
print(f"\n2026 + MCID 在 115 的筆數: {len(p0_115):,}")
print(f"這些 MCID 有幾個唯一值: {p0_115['merchant_customer_id'].nunique()}")

print("\n-- cal_type 分布 --")
print(p0_115["cal_type"].value_counts(dropna=False))
print("\n-- launch_channel 分布 --")
print(p0_115["launch_channel"].value_counts(dropna=False))
print("\n-- calendar_month 分布 --")
print(p0_115["calendar_month"].value_counts(dropna=False).sort_index())
print("\n-- marketplace_id 分布 --")
print(p0_115["marketplace_id"].value_counts(dropna=False))

# 試幾種組合比較
print("\n=== 不同過濾條件下 Jan/Feb/Mar GMS 加總 ===")
for label, flt in [
    ("僅 MCID∈115", p0_115),
    ("MCID + DSR", p0_115[p0_115["launch_channel"] == "DSR"]),
    ("MCID + cal_type=...", p0_115),  # 先看 cal_type 有哪些再決定
]:
    if flt is None:
        continue
    for m in [1, 2, 3]:
        s = flt[flt["calendar_month"] == m]["mtd_ord_gms"].sum()
        print(f"  {label:40s} | month={m}: {s:,.2f} (rows={len(flt[flt['calendar_month']==m])})")
