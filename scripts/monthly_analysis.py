"""區塊 3:月度銷售額變化 (ACC 1.0 / 2.0 / NSR All)。

資料來源:
- ACC 1.0 GMS:直接從「115 seller performance summary」P24:Y29 讀(Jan-Aug)。
- ACC 2.0 GMS:以 ACC 2.0 最終錄取 115 位的 MCID,去最新 P0 檔 Sheet1 過濾
  calendar_year=2026 + launch_channel=DSR,依 calendar_month 分組加總 mtd_ord_gms。
- NSR All:同樣 P0 過濾條件,但不做 MCID 篩,依 calendar_month 加總。

最新 P0 = TWGS - 2026 MBR 底下最新月份資料夾 (命名如「3. Mar」) 的 P0*.xlsx。
若有多個 P0 取最新修改時間者。
"""
from __future__ import annotations

import re
from pathlib import Path
from dataclasses import dataclass

import pandas as pd

# ============================================================
# 設定
# ============================================================

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
MBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR")

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug"]
MONTH_NUM = {m: i + 1 for i, m in enumerate(MONTH_ABBR)}

CALENDAR_YEAR = 2026
LAUNCH_CHANNEL = "DSR"


# ============================================================
# 最新 P0 定位
# ============================================================

def find_latest_month_folder(base: Path) -> Path:
    """MBR 底下命名如『1. Jan / 2. Feb ...』,取編號最大的。"""
    pattern = re.compile(r"^(\d+)\.\s*([A-Za-z]+)$")
    candidates = []
    for child in base.iterdir():
        if not child.is_dir():
            continue
        m = pattern.match(child.name)
        if m:
            candidates.append((int(m.group(1)), child))
    if not candidates:
        raise FileNotFoundError(f"{base} 底下找不到月份資料夾")
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]


def find_latest_p0_in_folder(folder: Path) -> Path:
    """挑 P0 開頭 xlsx,取最新修改時間者 (不分大小寫)。"""
    files = [
        f for f in folder.iterdir()
        if f.is_file()
        and f.suffix.lower() == ".xlsx"
        and f.name.lower().startswith("p0")
    ]
    if not files:
        raise FileNotFoundError(f"{folder} 沒有 P0 開頭檔案")
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0]


# ============================================================
# 讀取 1.0 數字 (P24:Y29)
# ============================================================

@dataclass
class MonthlyRow:
    """一列月度數字。months 對應 MONTH_ABBR 順序,None 代表無資料。"""
    label: str
    values: dict[str, float | None]


def load_10_row() -> MonthlyRow:
    """從 115 seller performance summary P24:Y29 讀 ACC 1.0 那列。

    P24:Y29 實際對應 0-index row 23..28, col 15..24:
      row 23: 標題列 (GMS / Jan. / Feb. / ...)
      row 24: ACC 1.0 數字
      row 25: ACC 2.0 數字
      ...
    Jan..Aug 月份在 col 16..23。
    """
    df = pd.read_excel(ACC_FILE, sheet_name="115 seller performance summary",
                       header=None, engine="openpyxl")
    # 先確認 row 24 第 15 欄標籤是「ACC 1.0」
    label = df.iat[24, 15]
    assert str(label).strip().lower().startswith("acc 1.0"), \
        f"預期 row 24 col 15 = 'ACC 1.0',實際為 {label!r}"
    row = df.iloc[24, 16:24]
    values: dict[str, float | None] = {}
    for m, v in zip(MONTH_ABBR, row.values):
        values[m] = float(v) if pd.notna(v) else None
    return MonthlyRow(label="ACC 1.0", values=values)


# ============================================================
# 讀取 2.0 MCID
# ============================================================

def load_20_mcids() -> set[str]:
    """取 ACC 2.0 最終賣家名單 (113) 的 MCID 集合。"""
    df = pd.read_excel(ACC_FILE, sheet_name="最終賣家名單 (113)", engine="openpyxl")
    if len(df) != 110:
        print(f"  [注意] 最終賣家名單筆數={len(df)},非預期的 110,以實際筆數為準")
    return set(df["MCID"].dropna().astype(str).str.strip())


# ============================================================
# 讀取 P0 並計算 2.0 / NSR All
# ============================================================

P0_NEEDED_COLS = [
    "calendar_year", "calendar_month", "launch_channel",
    "merchant_customer_id", "mtd_ord_gms",
]


def load_p0_filtered(p0_path: Path) -> pd.DataFrame:
    """讀 P0 Sheet1,只保留需要欄位,並過濾 calendar_year=2026 + launch_channel=DSR。"""
    print(f"  讀取 P0 (只取需要欄位): {p0_path.name}")
    df = pd.read_excel(
        p0_path,
        sheet_name="Sheet1",
        engine="openpyxl",
        usecols=P0_NEEDED_COLS,
    )
    print(f"    原始筆數: {len(df):,}")
    df = df[df["calendar_year"] == CALENDAR_YEAR].copy()
    df = df[df["launch_channel"] == LAUNCH_CHANNEL].copy()
    print(f"    過濾 year={CALENDAR_YEAR} + launch_channel={LAUNCH_CHANNEL} 後: {len(df):,}")
    # 型別正規化:MBR 的 merchant_customer_id 是 float (帶 .0),轉 int 再字串化以對齊 ACC int MCID
    df = df.dropna(subset=["merchant_customer_id"]).copy()
    df["merchant_customer_id"] = df["merchant_customer_id"].astype("int64").astype(str).str.strip()
    df["calendar_month"] = df["calendar_month"].astype(int)
    df["mtd_ord_gms"] = pd.to_numeric(df["mtd_ord_gms"], errors="coerce").fillna(0.0)
    return df


def calc_20_row(p0: pd.DataFrame, mcids: set[str]) -> MonthlyRow:
    sub = p0[p0["merchant_customer_id"].isin(mcids)]
    monthly = sub.groupby("calendar_month")["mtd_ord_gms"].sum()
    values: dict[str, float | None] = {}
    for m in MONTH_ABBR:
        num = MONTH_NUM[m]
        values[m] = float(monthly.get(num)) if num in monthly.index else None
    return MonthlyRow(label="ACC 2.0", values=values)


def calc_nsr_all_row(p0: pd.DataFrame) -> MonthlyRow:
    monthly = p0.groupby("calendar_month")["mtd_ord_gms"].sum()
    values: dict[str, float | None] = {}
    for m in MONTH_ABBR:
        num = MONTH_NUM[m]
        values[m] = float(monthly.get(num)) if num in monthly.index else None
    return MonthlyRow(label="2026 NSR All", values=values)


# ============================================================
# 衍生指標
# ============================================================

def calc_vs_10(row_20: MonthlyRow, row_10: MonthlyRow) -> MonthlyRow:
    """2.0 vs 1.0 的 YoY:(2.0 / 1.0) - 1。僅在兩邊都有值時計算。"""
    values: dict[str, float | None] = {}
    for m in MONTH_ABBR:
        a = row_20.values[m]
        b = row_10.values[m]
        if a is None or b is None or b == 0:
            values[m] = None
        else:
            values[m] = a / b - 1
    return MonthlyRow(label="2.0 vs 1.0", values=values)


def calc_mom(row_20: MonthlyRow) -> MonthlyRow:
    """2.0 的 MoM,以 2.0 前一個月為基期。第一個有值月份為 None。"""
    values: dict[str, float | None] = {}
    prev = None
    for m in MONTH_ABBR:
        cur = row_20.values[m]
        if cur is None or prev is None or prev == 0:
            values[m] = None
        else:
            values[m] = (cur - prev) / prev
        prev = cur if cur is not None else prev
    return MonthlyRow(label="MoM", values=values)


def calc_pct_of_nsr(row_20: MonthlyRow, row_nsr: MonthlyRow) -> MonthlyRow:
    values: dict[str, float | None] = {}
    for m in MONTH_ABBR:
        a = row_20.values[m]
        b = row_nsr.values[m]
        if a is None or b is None or b == 0:
            values[m] = None
        else:
            values[m] = a / b
    return MonthlyRow(label="% of NSR all", values=values)


# ============================================================
# 組裝一次性的結果物件
# ============================================================

@dataclass
class MonthlyResult:
    row_10: MonthlyRow
    row_20: MonthlyRow
    row_mom: MonthlyRow
    row_vs_10: MonthlyRow
    row_nsr: MonthlyRow
    row_pct: MonthlyRow
    p0_source: Path


def build_monthly_result() -> MonthlyResult:
    print("[1/4] 讀 1.0 P24:Y29...")
    row_10 = load_10_row()
    print(f"  1.0: {row_10.values}")

    print("[2/4] 取 2.0 MCID (110 位)...")
    mcids = load_20_mcids()
    print(f"  MCID 數量: {len(mcids)}")

    print("[3/4] 定位最新 P0...")
    latest_folder = find_latest_month_folder(MBR_BASE)
    p0 = find_latest_p0_in_folder(latest_folder)
    print(f"  最新月份資料夾: {latest_folder.name}")
    print(f"  P0 檔: {p0.name}")

    p0_df = load_p0_filtered(p0)

    print("[4/4] 計算 2.0 / NSR All / MoM / % of NSR...")
    row_20 = calc_20_row(p0_df, mcids)
    row_nsr = calc_nsr_all_row(p0_df)
    row_mom = calc_mom(row_20)
    row_vs_10 = calc_vs_10(row_20, row_10)
    row_pct = calc_pct_of_nsr(row_20, row_nsr)

    print(f"  2.0:     {row_20.values}")
    print(f"  NSR:     {row_nsr.values}")
    print(f"  MoM:     {row_mom.values}")
    print(f"  vs 1.0:  {row_vs_10.values}")
    print(f"  %NSR:    {row_pct.values}")

    return MonthlyResult(
        row_10=row_10,
        row_20=row_20,
        row_mom=row_mom,
        row_vs_10=row_vs_10,
        row_nsr=row_nsr,
        row_pct=row_pct,
        p0_source=p0,
    )


if __name__ == "__main__":
    result = build_monthly_result()
    print("\n=== 最終結果 ===")
    for r in [result.row_10, result.row_20, result.row_mom, result.row_nsr, result.row_pct]:
        print(r)
