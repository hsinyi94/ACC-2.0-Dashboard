"""找出在最終名單中但不在 NSR Launch Tracker 2026 Raw 的賣家 (未開賣)。"""
import re
from pathlib import Path
import pandas as pd

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
WBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")

# 讀最終賣家名單
roster = pd.read_excel(ACC_FILE, sheet_name="最終賣家名單 (115)", engine="openpyxl")
roster["MCID_str"] = roster["MCID"].dropna().astype(str).str.strip()
all_mcids = set(roster["MCID_str"].dropna())
print(f"最終賣家名單 MCID 數: {len(all_mcids)}")

# 找最新 NSR Launch Tracker
pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
candidates = []
for child in WBR_BASE.iterdir():
    if not child.is_dir():
        continue
    m = pattern.match(child.name)
    if not m:
        continue
    tracker = [
        f for f in child.iterdir()
        if f.is_file() and f.suffix.lower() == ".xlsx"
        and f.name.lower().startswith("nsr launch tracker")
    ]
    if tracker:
        tracker.sort(key=lambda f: f.stat().st_mtime, reverse=True)
        candidates.append((int(m.group(1)), tracker[0]))

candidates.sort(key=lambda x: x[0])
tracker_path = candidates[-1][1]
print(f"NSR Launch Tracker: {tracker_path.parent.name}/{tracker_path.name}")

# 讀 Tracker
df = pd.read_excel(tracker_path, sheet_name="2026 Raw", engine="openpyxl",
                   usecols=["merchant_customer_id"])
df = df.dropna(subset=["merchant_customer_id"])
df["mcid"] = df["merchant_customer_id"].apply(
    lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x).strip()
)
tracker_mcids = set(df["mcid"])
print(f"Tracker 中唯一 MCID: {len(set(df['mcid']))}")

# 找差集
not_launched = all_mcids - tracker_mcids
print(f"\n未開賣賣家 ({len(not_launched)} 位):")
print("-" * 80)

# 找欄位名
name_col = next((c for c in roster.columns if "公司名稱" in str(c)), None)
ae_col = "AE" if "AE" in roster.columns else None

for mcid in sorted(not_launched):
    row = roster[roster["MCID_str"] == mcid]
    if not row.empty:
        name = row.iloc[0][name_col] if name_col else "N/A"
        ae = row.iloc[0][ae_col] if ae_col else "N/A"
        print(f"  MCID: {mcid} | 公司: {name} | AE: {ae}")
