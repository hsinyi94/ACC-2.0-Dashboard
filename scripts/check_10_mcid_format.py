"""看 ACC 1.0 MCID 欄位的各種格式,找出多站點的並嘗試拆開。"""
import pandas as pd
import re
from pathlib import Path

ACC_10_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")

df = pd.read_excel(ACC_10_FILE, sheet_name="GMS by seller", engine="openpyxl")
print(f"總筆數: {len(df)}")

# 看 MCID 欄位的原始值
multi = []
single = []
for v in df["MCID"].dropna():
    s = str(v).strip()
    if "\n" in s or ":" in s or "," in s or "/" in s:
        multi.append(s)
    else:
        single.append(s)

print(f"單一 MCID: {len(single)}")
print(f"多站點 MCID: {len(multi)}")
print("\n=== 多站點 MCID 原始值 ===")
for m in multi:
    print(f"  {m!r}")

# 嘗試拆開:取第一個數字序列
def extract_first_mcid(s: str) -> str:
    """從多站點格式中取第一個 MCID (純數字)。"""
    # 格式如 "EU: 319940101412\nSG: 874661700702" 或 "JP: 889262429902\nMENA: ..."
    # 取第一個出現的純數字序列 (長度 > 8)
    matches = re.findall(r'\d{9,}', s)
    if matches:
        return matches[0]
    return s

print("\n=== 拆開後取第一個 ===")
all_mcids = set()
for v in df["MCID"].dropna():
    s = str(v).strip()
    mcid = extract_first_mcid(s)
    all_mcids.add(mcid)
    if s in multi:
        print(f"  {s!r} -> {mcid}")

print(f"\n拆開後唯一 MCID 數: {len(all_mcids)}")

# 用拆開後的 MCID 去 match wk52 的 2025 Raw
NSR = Path(r"W:\Team Spaces\TWGS\00_Commonly Accessed\2025 Weekly Business Report\wk52\NSR Launch Tracker_20251220.xlsx")
nsr_df = pd.read_excel(NSR, sheet_name="2025 Raw", engine="openpyxl",
                       usecols=["merchant_customer_id"])
nsr_df = nsr_df.dropna(subset=["merchant_customer_id"])
nsr_df["merchant_customer_id"] = nsr_df["merchant_customer_id"].apply(
    lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x).strip()
)
nsr_mcids = set(nsr_df["merchant_customer_id"])
matched = all_mcids & nsr_mcids
print(f"wk52 matched (拆開後): {len(matched)} / {len(all_mcids)}")

# 看哪些還是 match 不到
unmatched = all_mcids - nsr_mcids
print(f"未 match: {len(unmatched)}")
for u in sorted(unmatched):
    print(f"  {u}")
