"""區塊 4:YTD 達標率分析。

資料流:
- 從 ACC 2.0 最終錄取結果 -> 115 seller performance summary P2:Z5 讀:
  * ACC 2.0 的 GMS/Seller 2025 / Goal / YoY Goal
  * ACC 2.0 的 GMS 2025 / Goal / YoY Goal  (實際上是 GMS/Seller × 賣家數)
  * Program Enrolled Seller (2.0=115, NSR=553)

- 從 WBR 最新一週 P0 (TWGS - 2026 WBR/wk?? 的 P0*.xlsx, raw 工作表):
  篩選 launch_channel=DSR,以 reporting_year + reporting_week_of_year=max 取該週 YTD
  * 2.0 GMS 2026 Actual:MCID ∈ 115 位,sum(ytd_ord_gms)
  * NSR All 2026 Actual:不過濾 MCID,sum(ytd_ord_gms)
  * NSR All 2025:reporting_year=2025,同 week=max + DSR,sum(ytd_ord_gms)

衍生:
- GMS/Seller 2026 Actual = GMS Actual / Enrolled Seller
- YoY Actual = Actual / 2025 - 1
- 達標率 = Actual / Goal (NSR 無 Goal 時留白)
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
ACC_10_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")
WBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")

LAUNCH_CHANNEL = "DSR"


# ============================================================
# 最新週 P0 定位
# ============================================================

def find_latest_week_folder(base: Path) -> Path:
    """找最新的有 P0 檔的 wk 資料夾。"""
    pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
    candidates = []
    for child in base.iterdir():
        if child.is_dir():
            m = pattern.match(child.name)
            if m:
                # 確認裡面有 P0 開頭的 xlsx
                has_p0 = any(
                    f.is_file() and f.suffix.lower() == ".xlsx"
                    and f.name.lower().startswith("p0")
                    for f in child.iterdir()
                )
                if has_p0:
                    candidates.append((int(m.group(1)), child))
    if not candidates:
        raise FileNotFoundError(f"{base} 底下找不到含 P0 的 wk 資料夾")
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]


def find_latest_p0(folder: Path) -> Path:
    files = [
        f for f in folder.iterdir()
        if f.is_file() and f.suffix.lower() == ".xlsx"
        and f.name.lower().startswith("p0")
    ]
    if not files:
        raise FileNotFoundError(f"{folder} 無 P0 開頭檔案")
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0]


# ============================================================
# 讀 P2:Z5 靜態資料
# ============================================================

@dataclass
class PlanData:
    """來自 P2:Z5 的規劃數字 (ACC 2.0)。"""
    gms_per_seller_goal: float  # 62743 (2026 Goal)
    gms_per_seller_yoy_goal: float  # 0.220
    gms_goal: float            # 7215489 (2026 GMS Goal)
    gms_yoy_goal: float        # 0.108
    enrolled_20: int           # 115
    enrolled_nsr: int          # 553
    # 2025 base 改從 P24:Y29 的 ACC 1.0 Jan-Aug 加總取得
    gms_10_2025: float         # ACC 1.0 2025 Jan-Aug 合計
    enrolled_10: int           # ACC 1.0 賽道賽家數,固定 70


def load_plan_data() -> PlanData:
    df = pd.read_excel(ACC_FILE, sheet_name="115 seller performance summary",
                       header=None, engine="openpyxl")

    def _num(r, c):
        v = df.iat[r, c]
        return float(v) if pd.notna(v) else None

    # row 3 = ACC 2.0 (P2:Z5)
    # row 4 = 2026 NSR All
    # row 24 col 16-23 = ACC 1.0 Jan-Aug (P24:Y29)
    gms_10_months = []
    for c in range(16, 24):
        v = df.iat[24, c]
        if pd.notna(v) and isinstance(v, (int, float)):
            gms_10_months.append(float(v))
    gms_10_total = sum(gms_10_months)

    return PlanData(
        gms_per_seller_goal=_num(3, 18),
        gms_per_seller_yoy_goal=_num(3, 20),
        gms_goal=_num(3, 22),
        gms_yoy_goal=_num(3, 24),
        enrolled_20=113,  # 最終賣家名單 (113)
        enrolled_nsr=int(_num(4, 25)) if _num(4, 25) else 0,
        gms_10_2025=gms_10_total,
        enrolled_10=70,
    )


# ============================================================
# 讀 ACC 2.0 MCID
# ============================================================

def load_20_mcids() -> set[str]:
    """取 ACC 2.0 最終賣家名單 (113) 的 MCID 集合。"""
    df = pd.read_excel(ACC_FILE, sheet_name="最終賣家名單 (113)", engine="openpyxl")
    if len(df) != 110:
        print(f"  [注意] 最終賣家名單筆數={len(df)},非預期的 110,以實際筆數為準")
    return set(df["MCID"].dropna().astype(str).str.strip())


# ============================================================
# WBR P0 抓 YTD
# ============================================================

P0_NEEDED_COLS = ["reporting_year", "reporting_week_of_year",
                  "launch_channel", "merchant_customer_id", "ytd_ord_gms"]


@dataclass
class YTDResult:
    """YTD 結果集。"""
    week: int                    # 最新週 (取 2026 的 max)
    gms_20_ytd: float            # ACC 2.0 YTD sum
    gms_10_ytd: float            # ACC 1.0 YTD sum (2025 同週)
    gms_nsr_2026_ytd: float      # NSR All 2026 YTD sum
    gms_nsr_2025_ytd: float      # NSR All 2025 YTD sum
    p0_source: Path


def calc_ytd(p0_path: Path, mcids_20: set[str], mcids_10: set[str]) -> YTDResult:
    print(f"  讀取 WBR P0: {p0_path.name}")
    df = pd.read_excel(p0_path, sheet_name="raw",
                       engine="openpyxl", usecols=P0_NEEDED_COLS)
    print(f"    原始筆數: {len(df):,}")
    df = df[df["launch_channel"] == LAUNCH_CHANNEL].copy()
    # P0 的 merchant_customer_id 是 float (帶 .0),要先轉 int 再字串化,才能對齊 ACC 的 int MCID
    df = df.dropna(subset=["merchant_customer_id"]).copy()
    df["merchant_customer_id"] = df["merchant_customer_id"].astype("int64").astype(str).str.strip()
    df["ytd_ord_gms"] = pd.to_numeric(df["ytd_ord_gms"], errors="coerce").fillna(0.0)

    # 2026 取 max week
    df_2026 = df[df["reporting_year"] == 2026]
    max_wk_2026 = int(df_2026["reporting_week_of_year"].max())
    latest_2026 = df_2026[df_2026["reporting_week_of_year"] == max_wk_2026]

    # 2025 也取 max week (同一份檔)
    df_2025 = df[df["reporting_year"] == 2025]
    max_wk_2025 = int(df_2025["reporting_week_of_year"].max())
    latest_2025 = df_2025[df_2025["reporting_week_of_year"] == max_wk_2025]

    # ACC 2.0 2026 YTD
    gms_20_ytd = latest_2026[
        latest_2026["merchant_customer_id"].isin(mcids_20)
    ]["ytd_ord_gms"].sum()

    gms_nsr_2026_ytd = latest_2026["ytd_ord_gms"].sum()
    gms_nsr_2025_ytd = latest_2025["ytd_ord_gms"].sum()

    # ACC 1.0 YTD (2025 同週, MCID ∈ 1.0)
    gms_10_ytd = latest_2025[
        latest_2025["merchant_customer_id"].isin(mcids_10)
    ]["ytd_ord_gms"].sum()

    print(f"    2026 max week={max_wk_2026}, DSR 筆數={len(latest_2026):,}")
    print(f"    2025 max week={max_wk_2025}, DSR 筆數={len(latest_2025):,}")

    return YTDResult(
        week=max_wk_2026,
        gms_20_ytd=float(gms_20_ytd),
        gms_10_ytd=float(gms_10_ytd),
        gms_nsr_2026_ytd=float(gms_nsr_2026_ytd),
        gms_nsr_2025_ytd=float(gms_nsr_2025_ytd),
        p0_source=p0_path,
    )


# ============================================================
# 組裝
# ============================================================

@dataclass
class YTDAnalysis:
    plan: PlanData
    ytd: YTDResult

    # 衍生算 2.0
    def row_20(self) -> dict:
        n20 = self.plan.enrolled_20 or 1
        n10 = self.plan.enrolled_10 or 1
        actual_gms = self.ytd.gms_20_ytd
        actual_per = actual_gms / n20
        # 2025 用 ACC 1.0 Jan-Aug 合計 (來自 P24:Y29),per seller 用 1.0 賽道賽家數 70
        gms_2025 = self.plan.gms_10_2025
        ps_2025 = gms_2025 / n10 if gms_2025 else None
        v_goal = self.plan.gms_per_seller_goal
        gms_goal = self.plan.gms_goal
        return {
            # GMS/Seller
            "ps_2025": ps_2025,
            "ps_goal": v_goal,
            "ps_actual": actual_per,
            "ps_yoy_actual": (actual_per / ps_2025 - 1) if ps_2025 else None,
            "ps_yoy_goal": self.plan.gms_per_seller_yoy_goal,
            "ps_attain": (actual_per / v_goal) if v_goal else None,
            # GMS
            "g_2025": gms_2025,
            "g_goal": gms_goal,
            "g_actual": actual_gms,
            "g_yoy_actual": (actual_gms / gms_2025 - 1) if gms_2025 else None,
            "g_yoy_goal": self.plan.gms_yoy_goal,
            "g_attain": (actual_gms / gms_goal) if gms_goal else None,
            "n": n20,
        }

    def row_nsr(self) -> dict:
        n = self.plan.enrolled_nsr or 1
        actual_gms = self.ytd.gms_nsr_2026_ytd
        actual_per = actual_gms / n
        gms_2025 = self.ytd.gms_nsr_2025_ytd
        # NSR 的 per seller 2025 = 2025 YTD / enrolled (seller 數用同一個)
        ps_2025 = gms_2025 / n if gms_2025 else None
        return {
            # GMS/Seller
            "ps_2025": ps_2025,
            "ps_goal": None,
            "ps_actual": actual_per,
            "ps_yoy_actual": (actual_per / ps_2025 - 1) if ps_2025 else None,
            "ps_yoy_goal": None,
            "ps_attain": None,
            # GMS
            "g_2025": gms_2025,
            "g_goal": None,
            "g_actual": actual_gms,
            "g_yoy_actual": (actual_gms / gms_2025 - 1) if gms_2025 else None,
            "g_yoy_goal": None,
            "g_attain": None,
            "n": n,
        }


def build_ytd_analysis() -> YTDAnalysis:
    print("[1/4] 讀 P2:Z5 + P24:Y29 計畫數字...")
    plan = load_plan_data()
    print(f"  ACC 2.0 Goal: per-seller={plan.gms_per_seller_goal}, GMS={plan.gms_goal}, Enrolled={plan.enrolled_20}")
    print(f"  ACC 1.0 2025 GMS (Jan-Aug 合計) = {plan.gms_10_2025:,.0f}, per seller = {plan.gms_10_2025 / plan.enrolled_10:,.0f}")
    print(f"  NSR All: Enrolled={plan.enrolled_nsr}")

    print("[2/4] 取 2.0 MCID...")
    mcids = load_20_mcids()
    print(f"  MCID 數量: {len(mcids)}")

    # 也載入 1.0 MCID (用於計算 1.0 YTD)
    import re as _re
    df_10_mcid = pd.read_excel(ACC_10_FILE, sheet_name="GMS by seller", engine="openpyxl")
    mcids_10 = set()
    for v in df_10_mcid["MCID"].dropna():
        s = str(v).strip()
        matches = _re.findall(r'\d{9,}', s)
        if matches:
            mcids_10.add(matches[0])
        else:
            mcids_10.add(s)
    print(f"  1.0 MCID 數量: {len(mcids_10)}")

    print("[3/4] 定位最新週 P0...")
    folder = find_latest_week_folder(WBR_BASE)
    p0 = find_latest_p0(folder)
    print(f"  最新週資料夾: {folder.name}, P0: {p0.name}")

    print("[4/4] 計算 YTD...")
    ytd = calc_ytd(p0, mcids, mcids_10)
    print(f"  week {ytd.week} | 2.0 YTD={ytd.gms_20_ytd:,.2f}")
    print(f"                 | 1.0 YTD={ytd.gms_10_ytd:,.2f}")
    print(f"                 | NSR 2026 YTD={ytd.gms_nsr_2026_ytd:,.2f}")
    print(f"                 | NSR 2025 YTD={ytd.gms_nsr_2025_ytd:,.2f}")

    return YTDAnalysis(plan=plan, ytd=ytd)


if __name__ == "__main__":
    result = build_ytd_analysis()
    print("\n=== ACC 2.0 列 ===")
    r20 = result.row_20()
    for k, v in r20.items():
        print(f"  {k}: {v}")
    print("\n=== NSR All 列 ===")
    rnsr = result.row_nsr()
    for k, v in rnsr.items():
        print(f"  {k}: {v}")
