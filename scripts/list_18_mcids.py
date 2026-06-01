"""列出 Book1 有但不在 115 錄取名單的 18 個 MCID 的詳情。"""
import pandas as pd
from pathlib import Path

BOOK1 = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\Book1.xlsx")
ACC = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")

# 讀 Book1
b = pd.read_excel(BOOK1, sheet_name="Sheet1", engine="openpyxl")
b["merchant_customer_id"] = b["merchant_customer_id"].astype("int64").astype(str)

# 讀 ACC 2.0 全名單 (含候補,看 18 個 MCID 是不是候補)
acc_all = pd.read_excel(ACC, sheet_name="最終錄取115位 + 候補18位", engine="openpyxl")
acc_all["MCID_str"] = acc_all["MCID"].astype(str).str.strip()
mcids_115 = set(acc_all[acc_all["No."].notna()]["MCID_str"])
mcids_all = set(acc_all["MCID_str"].dropna())

# Book1 中不在 115 的 MCID
b_mcids = set(b["merchant_customer_id"])
not_in_115 = b_mcids - mcids_115

print(f"Book1 MCID: {len(b_mcids)}")
print(f"不在 115 的 MCID 數: {len(not_in_115)}")
print(f"在 ACC 全名單 (含候補 133+雜項): {len(not_in_115 & mcids_all)}")
print(f"完全不在 ACC 任何名單: {len(not_in_115 - mcids_all)}")

# 每個 MCID 的詳情:公司名、launch_channel、三個月 GMS
print("\n=== 18 個不在 115 名單的 MCID 詳情 ===\n")
rows = []
for mcid in sorted(not_in_115):
    sub = b[b["merchant_customer_id"] == mcid]
    # 看 ACC 全名單有無此 MCID
    in_acc = acc_all[acc_all["MCID_str"] == mcid]
    if len(in_acc):
        company = in_acc["公司名稱 (請填寫完整公司設立登記名稱，未成立公司請填無)"].iloc[0]
        no = in_acc["No."].iloc[0]
        status_acc = f"在 ACC 名單 (No.={no})" if pd.notna(no) else "在 ACC 名單但 No. 空 (候補或其他)"
    else:
        # 從 P0 資料找 merchant_name
        mn = sub["merchant_name"].dropna()
        company = mn.iloc[0] if len(mn) else "(無法查到)"
        status_acc = "不在 ACC 名單"

    total_gms = sub["mtd_ord_gms"].sum()
    channels = sub["launch_channel"].unique().tolist()
    markets = sub["marketplace_id"].unique().tolist()
    months = sub["calendar_month"].unique().tolist()
    rows.append({
        "MCID": mcid,
        "公司名": company,
        "ACC 狀態": status_acc,
        "三個月 GMS 合計": f"{total_gms:,.2f}",
        "launch_channel": channels,
        "涵蓋月份": sorted(months),
        "marketplace 數": len(markets),
    })

df_out = pd.DataFrame(rows)
print(df_out.to_string(index=False, max_colwidth=40))

# 三個月 GMS 加總 (18 個 MCID)
print("\n=== 這 18 個 MCID 的每月 GMS 加總 ===")
extra_b = b[b["merchant_customer_id"].isin(not_in_115)]
for m in [1, 2, 3]:
    s = extra_b[extra_b["calendar_month"] == m]["mtd_ord_gms"].sum()
    print(f"  month={m}: GMS={s:,.2f}")
