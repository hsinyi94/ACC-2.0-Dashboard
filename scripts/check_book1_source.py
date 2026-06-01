"""確認 Book1 的來源與 MCID 範圍。"""
import pandas as pd
from pathlib import Path

BOOK1 = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\Book1.xlsx")
ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
MBR_P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR\3. Mar\P0_20260331.xlsx")
WBR_P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wk17\P0_20260425.xlsx")

# Book1 的 MCID
b = pd.read_excel(BOOK1, sheet_name="Sheet1", engine="openpyxl",
                  usecols=["calendar_month", "calendar_year",
                           "launch_channel", "merchant_customer_id", "mtd_ord_gms"])
b["merchant_customer_id"] = b["merchant_customer_id"].astype("int64").astype(str)
b_mcids = set(b["merchant_customer_id"])
print(f"Book1 筆數: {len(b)}")
print(f"Book1 唯一 MCID: {len(b_mcids)}")
print(f"Book1 launch_channel 分布: {b['launch_channel'].value_counts().to_dict()}")

# 2.0 115 MCID
acc_df = pd.read_excel(ACC, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
acc_df = acc_df[acc_df["No."].notna()]
mcids_115 = set(acc_df["MCID"].astype(str))
print(f"\n115 MCID 數: {len(mcids_115)}")
print(f"Book1 ∩ 115: {len(b_mcids & mcids_115)}")
print(f"Book1 - 115 (Book1 有 115 沒): {len(b_mcids - mcids_115)}")
print(f"115 - Book1 (115 有 Book1 沒): {len(mcids_115 - b_mcids)}")

# 看 Book1 裡 115 MCID 的筆數
b_115 = b[b["merchant_customer_id"].isin(mcids_115)]
print(f"\nBook1 中屬於 115 的筆數: {len(b_115)}")
for m in [1, 2, 3]:
    s = b_115[b_115["calendar_month"] == m]["mtd_ord_gms"].sum()
    n = len(b_115[b_115["calendar_month"] == m])
    print(f"  Book1 115 month={m}: GMS={s:,.2f}, rows={n}")

# 也看 MBR P0 裡只取 DSR 的 115 (我原本的做法)
cols = ["calendar_month", "calendar_year", "launch_channel",
        "merchant_customer_id", "mtd_ord_gms"]
p = pd.read_excel(MBR_P0, sheet_name="Sheet1", engine="openpyxl", usecols=cols)
p = p.dropna(subset=["merchant_customer_id"]).copy()
p["merchant_customer_id"] = p["merchant_customer_id"].astype("int64").astype(str)
p = p[(p["calendar_year"] == 2026) & (p["merchant_customer_id"].isin(mcids_115))]
print(f"\nMBR P0 中 115 MCID 筆數 (不過濾 launch_channel): {len(p)}")
for m in [1, 2, 3]:
    s = p[p["calendar_month"] == m]["mtd_ord_gms"].sum()
    n = len(p[p["calendar_month"] == m])
    print(f"  MBR 115 全 channel month={m}: GMS={s:,.2f}, rows={n}")
