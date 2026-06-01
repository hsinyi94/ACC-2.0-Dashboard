"""找出 Book1 (330) 和我算的 (359) 差異的 29 筆。"""
import pandas as pd
from pathlib import Path

BOOK1 = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\Book1.xlsx")
P0 = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR\3. Mar\P0_20260331.xlsx")
ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

# 讀 Book1
b = pd.read_excel(BOOK1, sheet_name="Sheet1", engine="openpyxl",
                  usecols=["cal_type", "calendar_month", "calendar_year",
                           "launch_channel", "marketplace_id",
                           "merchant_customer_id", "mtd_ord_gms"])
b["merchant_customer_id"] = b["merchant_customer_id"].astype("int64").astype(str)
print(f"Book1 筆數: {len(b)}")

# 讀 2.0 MCID
acc_df = pd.read_excel(ACC, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
acc_df = acc_df[acc_df["No."].notna()]
mcids = set(acc_df["MCID"].astype(str))

# 讀 P0 全部,同樣過濾:year=2026 + MCID∈115
cols = ["cal_type", "calendar_month", "calendar_year",
        "launch_channel", "marketplace_id",
        "merchant_customer_id", "mtd_ord_gms"]
p = pd.read_excel(P0, sheet_name="Sheet1", engine="openpyxl", usecols=cols)
p = p.dropna(subset=["merchant_customer_id"]).copy()
p["merchant_customer_id"] = p["merchant_customer_id"].astype("int64").astype(str)
p = p[p["calendar_year"] == 2026]
p115 = p[p["merchant_customer_id"].isin(mcids)].copy()
print(f"我的筆數 (2026+MCID): {len(p115)}")

# 用 (MCID, month, marketplace_id) 當 key 對比
def key(df):
    return df.apply(lambda r: (r["merchant_customer_id"], int(r["calendar_month"]), r["marketplace_id"]), axis=1)

b_keys = set(key(b))
p_keys = set(key(p115))
only_mine = p_keys - b_keys
only_book = b_keys - p_keys
print(f"只在我這的 (我有、Book1 沒): {len(only_mine)}")
print(f"只在 Book1 的 (我沒、Book1 有): {len(only_book)}")

# 列出差異
if only_mine:
    print("\n--- 只在我這 (可能是我多抓) ---")
    mask = p115.apply(lambda r: (r["merchant_customer_id"], int(r["calendar_month"]), r["marketplace_id"]) in only_mine, axis=1)
    extra = p115[mask]
    print(extra[["calendar_month", "merchant_customer_id", "marketplace_id", "launch_channel", "mtd_ord_gms"]].to_string())
    print(f"這些多出的 GMS 加總: {extra['mtd_ord_gms'].sum():,.2f}")
    print("\n多抓的 marketplace_id 分布:")
    print(extra["marketplace_id"].value_counts())
    print("\n多抓的 launch_channel 分布:")
    print(extra["launch_channel"].value_counts(dropna=False))

if only_book:
    print("\n--- 只在 Book1 (可能我漏抓) ---")
    mask = b.apply(lambda r: (r["merchant_customer_id"], int(r["calendar_month"]), r["marketplace_id"]) in only_book, axis=1)
    miss = b[mask]
    print(miss[["calendar_month", "merchant_customer_id", "marketplace_id", "launch_channel", "mtd_ord_gms"]].to_string())
