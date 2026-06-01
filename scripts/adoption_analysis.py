"""區塊 5:Seller Adoption 指標計數。

資料來源:
- NSR Launch Tracker (WBR 最新有此檔的 wk 資料夾),工作表「2026 Raw」
- 用 MCID mapping 分三組:ACC 2.0 / NSR All / ACC 1.0
- 每個 MCID 取 max(因為同 MCID 可能有多列),再 sum 計數
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
ACC_10_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")
WBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")
NSR_TRACKER_10 = Path(r"W:\Team Spaces\TWGS\00_Commonly Accessed\2025 Weekly Business Report\wk52\NSR Launch Tracker_20251220.xlsx")
NSR_TRACKER_10_SHEET = "2025 Raw"



ADOPTION_COLS = [
    "is_pl",
    "is_fba_adopt_by_seller",
    "is_sp_adopt_by_seller",
    "is_deal_adopt_by_seller",
    "is_brand_rep_by_seller",
    "is_b2b_adopt_by_seller",
]

ADOPTION_LABELS = {
    "is_pl": "PL (Perfect Launch)",
    "is_fba_adopt_by_seller": "FBA Adopt",
    "is_sp_adopt_by_seller": "SP (Sponsored Products) Adopt",
    "is_deal_adopt_by_seller": "Deal Adopt",
    "is_brand_rep_by_seller": "Brand Registry",
    "is_b2b_adopt_by_seller": "B2B Adopt",
}


# ============================================================
# 定位 NSR Launch Tracker
# ============================================================

def find_latest_nsr_tracker(base: Path) -> Path:
    """找最新有 NSR Launch Tracker 的 wk 資料夾。"""
    pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
    candidates = []
    for child in base.iterdir():
        if not child.is_dir():
            continue
        m = pattern.match(child.name)
        if not m:
            continue
        # 找 NSR Launch Tracker 開頭的 xlsx
        tracker = [
            f for f in child.iterdir()
            if f.is_file() and f.suffix.lower() == ".xlsx"
            and f.name.lower().startswith("nsr launch tracker")
        ]
        if tracker:
            tracker.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            candidates.append((int(m.group(1)), tracker[0]))
    if not candidates:
        raise FileNotFoundError(f"{base} 底下找不到含 NSR Launch Tracker 的 wk 資料夾")
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]


# ============================================================
# 讀 MCID 集合
# ============================================================

def load_20_mcids() -> set[str]:
    df = pd.read_excel(ACC_FILE, sheet_name="最終賣家名單 (113)", engine="openpyxl")
    return set(df["MCID"].dropna().astype(str).str.strip())


def load_10_mcids() -> set[str]:
    """從 ACC 1.0分析.xlsx 的 GMS by seller 取 MCID。
    多站點格式 (如 'EU: 123\nSG: 456') 只取第一個 MCID。
    """
    import re
    df = pd.read_excel(ACC_10_FILE, sheet_name="GMS by seller", engine="openpyxl")
    mcids = set()
    for v in df["MCID"].dropna():
        s = str(v).strip()
        # 取第一個長度 >= 9 的純數字序列
        matches = re.findall(r'\d{9,}', s)
        if matches:
            mcids.add(matches[0])
        else:
            mcids.add(s)
    return mcids


# ============================================================
# 計算 adoption
# ============================================================

@dataclass
class AdoptionRow:
    label: str
    count: int
    total: int
    pct: float


@dataclass
class AdoptionResult:
    rows_20: list[AdoptionRow]
    rows_nsr: list[AdoptionRow]
    rows_10: list[AdoptionRow]
    n_20: int
    n_nsr: int
    n_10: int
    source: Path


def calc_adoption(tracker_path: Path, mcids_20: set[str], mcids_10: set[str]) -> AdoptionResult:
    print(f"  讀取 2026 NSR Launch Tracker: {tracker_path.name}")
    needed = ["merchant_customer_id"] + ADOPTION_COLS
    df = pd.read_excel(tracker_path, sheet_name="2026 Raw",
                       engine="openpyxl", usecols=needed)
    print(f"    原始筆數: {len(df):,}")

    # 正規化 MCID
    df = df.dropna(subset=["merchant_customer_id"]).copy()
    df["merchant_customer_id"] = df["merchant_customer_id"].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x).strip()
    )

    # 每個 MCID 取 max (seller 層級)
    seller = df.groupby("merchant_customer_id")[ADOPTION_COLS].max().reset_index()
    print(f"    唯一 MCID (seller): {len(seller):,}")

    def _calc(subset: pd.DataFrame, total: int) -> list[AdoptionRow]:
        rows = []
        for col in ADOPTION_COLS:
            cnt = int(subset[col].sum())
            rows.append(AdoptionRow(
                label=ADOPTION_LABELS[col],
                count=cnt,
                total=total,
                pct=cnt / total if total else 0,
            ))
        return rows

    # ACC 2.0
    s20 = seller[seller["merchant_customer_id"].isin(mcids_20)]
    n20 = len(s20)

    # NSR All (2026)
    n_nsr = len(seller)

    # ACC 1.0 — 從 2025 Raw 讀,n 固定用 70(沒 match 到的當 0)
    print(f"  讀取 2025 NSR Launch Tracker: {NSR_TRACKER_10.name}")
    df_10 = pd.read_excel(NSR_TRACKER_10, sheet_name=NSR_TRACKER_10_SHEET,
                          engine="openpyxl", usecols=needed)
    print(f"    2025 原始筆數: {len(df_10):,}")
    df_10 = df_10.dropna(subset=["merchant_customer_id"]).copy()
    df_10["merchant_customer_id"] = df_10["merchant_customer_id"].apply(
        lambda x: str(int(x)) if pd.notna(x) and isinstance(x, (int, float)) else str(x).strip()
    )
    seller_10 = df_10.groupby("merchant_customer_id")[ADOPTION_COLS].max().reset_index()
    s10 = seller_10[seller_10["merchant_customer_id"].isin(mcids_10)]
    n10 = 70  # 固定用 ACC 1.0 賽道賣家數,沒 match 到的當 0

    print(f"    ACC 2.0 matched: {n20}, NSR All: {n_nsr}, ACC 1.0 matched: {n10}")

    return AdoptionResult(
        rows_20=_calc(s20, n20),
        rows_nsr=_calc(seller, n_nsr),
        rows_10=_calc(s10, n10),
        n_20=n20,
        n_nsr=n_nsr,
        n_10=n10,
        source=tracker_path,
    )


def build_adoption_analysis() -> AdoptionResult:
    print("[1/3] 取 MCID 集合...")
    mcids_20 = load_20_mcids()
    mcids_10 = load_10_mcids()
    print(f"  2.0: {len(mcids_20)}, 1.0: {len(mcids_10)}")

    print("[2/3] 定位最新 NSR Launch Tracker...")
    tracker = find_latest_nsr_tracker(WBR_BASE)
    print(f"  {tracker.parent.name}/{tracker.name}")

    print("[3/3] 計算 Adoption...")
    return calc_adoption(tracker, mcids_20, mcids_10)


if __name__ == "__main__":
    result = build_adoption_analysis()
    print(f"\n=== Adoption 結果 (2.0 n={result.n_20}, NSR n={result.n_nsr}, 1.0 n={result.n_10}) ===")
    print(f"{'指標':<30} {'2.0':>12} {'NSR All':>12} {'1.0':>12}")
    for r20, rnsr, r10 in zip(result.rows_20, result.rows_nsr, result.rows_10):
        print(f"{r20.label:<30} {r20.count:>4} ({r20.pct:.0%})  {rnsr.count:>4} ({rnsr.pct:.0%})  {r10.count:>4} ({r10.pct:.0%})")
