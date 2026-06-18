"""區塊 6:Raw Data — 每位 2.0 賣家的完整數據表。

合併來源:
1. 最終賣家名單 (113) — 基本資訊 + AE
2. MBR P0 — 月度 mtd_ord_gms (Jan-Aug)
3. WBR P0 — ytd_ord_gms
4. NSR Launch Tracker — adoption 指標
"""
from __future__ import annotations

import re
from pathlib import Path

import pandas as pd

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
MBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR")
WBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")

ROSTER_COLS = [
    "AE", "MCID", "公司名稱 (請填寫完整公司設立登記名稱，未成立公司請填無)",
    "錄取國家", "Category", "公司類型", "公司所在區域",
    "品牌創立年限", "公司產品類型", "公司主要經營型態", "開賣狀態",
]

ADOPTION_COLS = [
    "is_pl", "is_fba_adopt_by_seller", "is_sp_adopt_by_seller",
    "is_deal_adopt_by_seller", "is_brand_rep_by_seller", "is_b2b_adopt_by_seller",
]

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"]


def _find_latest_month_folder(base: Path) -> Path:
    pattern = re.compile(r"^(\d+)\.\s*([A-Za-z]+)$")
    candidates = []
    for child in base.iterdir():
        if child.is_dir():
            m = pattern.match(child.name)
            if m:
                candidates.append((int(m.group(1)), child))
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1] if candidates else None


def _find_p0(folder: Path) -> Path | None:
    files = [f for f in folder.iterdir()
             if f.is_file() and f.suffix.lower() == ".xlsx"
             and f.name.lower().startswith("p0")]
    if not files:
        return None
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0]


def _find_latest_wbr_p0(base: Path) -> Path | None:
    pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
    candidates = []
    for child in base.iterdir():
        if child.is_dir():
            m = pattern.match(child.name)
            if m:
                p0 = _find_p0(child)
                if p0:
                    candidates.append((int(m.group(1)), p0))
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1] if candidates else None


def _find_latest_nsr_tracker(base: Path) -> Path | None:
    pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
    candidates = []
    for child in base.iterdir():
        if child.is_dir():
            m = pattern.match(child.name)
            if m:
                tracker = [f for f in child.iterdir()
                           if f.is_file() and f.suffix.lower() == ".xlsx"
                           and f.name.lower().startswith("nsr launch tracker")]
                if tracker:
                    tracker.sort(key=lambda f: f.stat().st_mtime, reverse=True)
                    candidates.append((int(m.group(1)), tracker[0]))
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1] if candidates else None


def build_raw_data() -> pd.DataFrame:
    """組裝每位賣家的完整 raw data DataFrame。"""
    # 1. 基本資訊
    print("  [Raw] 讀取賣家名單...")
    roster = pd.read_excel(ACC_FILE, sheet_name="最終賣家名單 (115)", engine="openpyxl")
    # 簡化公司名稱欄位名
    rename_map = {"公司名稱 (請填寫完整公司設立登記名稱，未成立公司請填無)": "公司名稱"}
    roster = roster.rename(columns=rename_map)
    roster["MCID_str"] = roster["MCID"].astype(str).str.strip()
    base_cols = ["AE", "MCID_str", "公司名稱", "錄取國家", "Category",
                 "公司類型", "公司所在區域", "品牌創立年限", "公司產品類型",
                 "公司主要經營型態", "開賣狀態"]
    df = roster[base_cols].copy()
    df = df.rename(columns={"MCID_str": "MCID"})

    # 2. MBR P0 月度 GMS
    print("  [Raw] 讀取 MBR P0 月度...")
    mbr_folder = _find_latest_month_folder(MBR_BASE)
    mbr_p0 = _find_p0(mbr_folder) if mbr_folder else None
    if mbr_p0:
        cols_needed = ["calendar_year", "calendar_month", "launch_channel",
                       "merchant_customer_id", "mtd_ord_gms"]
        p0 = pd.read_excel(mbr_p0, sheet_name="Sheet1", engine="openpyxl", usecols=cols_needed)
        p0 = p0.dropna(subset=["merchant_customer_id"]).copy()
        p0["merchant_customer_id"] = p0["merchant_customer_id"].astype("int64").astype(str).str.strip()
        p0 = p0[(p0["calendar_year"] == 2026) & (p0["launch_channel"] == "DSR")]
        p0["mtd_ord_gms"] = pd.to_numeric(p0["mtd_ord_gms"], errors="coerce").fillna(0.0)
        # pivot: MCID x month
        monthly = p0.groupby(["merchant_customer_id", "calendar_month"])["mtd_ord_gms"].sum().reset_index()
        monthly_pivot = monthly.pivot(index="merchant_customer_id", columns="calendar_month", values="mtd_ord_gms")
        monthly_pivot.columns = [MONTH_ABBR[int(c)-1] + " GMS" for c in monthly_pivot.columns]
        monthly_pivot = monthly_pivot.reset_index().rename(columns={"merchant_customer_id": "MCID"})
        df = df.merge(monthly_pivot, on="MCID", how="left")

    # 3. WBR P0 YTD
    print("  [Raw] 讀取 WBR P0 YTD...")
    wbr_p0 = _find_latest_wbr_p0(WBR_BASE)
    if wbr_p0:
        cols_needed = ["reporting_year", "reporting_week_of_year", "launch_channel",
                       "merchant_customer_id", "ytd_ord_gms"]
        wp = pd.read_excel(wbr_p0, sheet_name="raw", engine="openpyxl", usecols=cols_needed)
        wp = wp.dropna(subset=["merchant_customer_id"]).copy()
        wp["merchant_customer_id"] = wp["merchant_customer_id"].astype("int64").astype(str).str.strip()
        wp = wp[(wp["reporting_year"] == 2026) & (wp["launch_channel"] == "DSR")]
        max_wk = wp["reporting_week_of_year"].max()
        wp = wp[wp["reporting_week_of_year"] == max_wk]
        wp["ytd_ord_gms"] = pd.to_numeric(wp["ytd_ord_gms"], errors="coerce").fillna(0.0)
        ytd = wp.groupby("merchant_customer_id")["ytd_ord_gms"].sum().reset_index()
        ytd = ytd.rename(columns={"merchant_customer_id": "MCID", "ytd_ord_gms": f"YTD GMS (W{int(max_wk)})"})
        df = df.merge(ytd, on="MCID", how="left")

    # 4. NSR Launch Tracker Adoption
    print("  [Raw] 讀取 NSR Launch Tracker...")
    tracker = _find_latest_nsr_tracker(WBR_BASE)
    if tracker:
        needed = ["merchant_customer_id"] + ADOPTION_COLS
        tr = pd.read_excel(tracker, sheet_name="2026 Raw", engine="openpyxl", usecols=needed)
        tr = tr.dropna(subset=["merchant_customer_id"]).copy()
        tr["merchant_customer_id"] = tr["merchant_customer_id"].apply(
            lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x).strip()
        )
        seller = tr.groupby("merchant_customer_id")[ADOPTION_COLS].max().reset_index()
        seller = seller.rename(columns={"merchant_customer_id": "MCID"})
        df = df.merge(seller, on="MCID", how="left")

    return df


if __name__ == "__main__":
    df = build_raw_data()
    print(f"\n最終 Raw Data: {df.shape}")
    print(df.head().to_string(max_cols=10))
