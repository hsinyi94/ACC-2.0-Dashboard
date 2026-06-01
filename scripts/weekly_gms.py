"""Weekly GMS tab:ACC 2.0 vs ACC 1.0 每週 wtd_ord_gms 變化。

一次性跑所有有 P0 的 wk 資料夾,結果 cache 到 output/weekly_gms_cache.json。
之後只需跑最新一週更新 cache。
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from dataclasses import dataclass

import pandas as pd

ACC_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx")
ACC_10_FILE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx")
WBR_BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")
CACHE_FILE = Path(__file__).resolve().parent.parent / "output" / "weekly_gms_cache.json"

LAUNCH_CHANNEL = "DSR"


def load_20_mcids() -> set[str]:
    df = pd.read_excel(ACC_FILE, sheet_name="最終賣家名單 (113)", engine="openpyxl")
    return set(df["MCID"].dropna().astype(str).str.strip())


def load_10_mcids() -> set[str]:
    df = pd.read_excel(ACC_10_FILE, sheet_name="GMS by seller", engine="openpyxl")
    mcids = set()
    for v in df["MCID"].dropna():
        s = str(v).strip()
        matches = re.findall(r'\d{9,}', s)
        if matches:
            mcids.add(matches[0])
        else:
            mcids.add(s)
    return mcids


def find_all_wk_p0s(base: Path) -> list[tuple[int, Path]]:
    """找所有有 P0 的 wk 資料夾,回傳 [(week_num, p0_path), ...] 排序。"""
    pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
    results = []
    for child in base.iterdir():
        if not child.is_dir():
            continue
        m = pattern.match(child.name)
        if not m:
            continue
        wk = int(m.group(1))
        p0_files = [
            f for f in child.iterdir()
            if f.is_file() and f.suffix.lower() == ".xlsx"
            and f.name.lower().startswith("p0")
        ]
        if p0_files:
            p0_files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
            results.append((wk, p0_files[0]))
    results.sort(key=lambda x: x[0])
    return results


def calc_week(p0_path: Path, week: int, mcids_20: set[str], mcids_10: set[str]) -> dict:
    """從一個 P0 檔計算該週的 2.0 和 1.0 wtd_ord_gms。"""
    cols = ["reporting_year", "reporting_week_of_year", "launch_channel",
            "merchant_customer_id", "wtd_ord_gms"]

    # 偵測工作表名 (有些是 "raw",有些可能是其他名稱)
    xls = pd.ExcelFile(p0_path, engine="openpyxl")
    sheet = None
    for candidate in ["raw", "Raw", "Sheet1"]:
        if candidate in xls.sheet_names:
            sheet = candidate
            break
    if sheet is None:
        # 找第一個有 reporting_year 欄位的工作表
        for sh in xls.sheet_names:
            test = pd.read_excel(p0_path, sheet_name=sh, engine="openpyxl", nrows=1)
            if "reporting_year" in test.columns:
                sheet = sh
                break
    if sheet is None:
        print(f"      [跳過] {p0_path.name} 找不到含 reporting_year 的工作表")
        return {"week": week, "gms_20": 0.0, "gms_10": 0.0}

    df = pd.read_excel(p0_path, sheet_name=sheet, engine="openpyxl", usecols=cols)
    df = df[df["launch_channel"] == LAUNCH_CHANNEL].copy()
    df = df.dropna(subset=["merchant_customer_id"]).copy()
    df["merchant_customer_id"] = df["merchant_customer_id"].astype("int64").astype(str).str.strip()
    df["wtd_ord_gms"] = pd.to_numeric(df["wtd_ord_gms"], errors="coerce").fillna(0.0)

    # ACC 2.0: year=2026, week=該週
    df_20 = df[(df["reporting_year"] == 2026) & (df["reporting_week_of_year"] == week)]
    gms_20 = float(df_20[df_20["merchant_customer_id"].isin(mcids_20)]["wtd_ord_gms"].sum())

    # ACC 1.0: year=2025, week=該週
    df_10 = df[(df["reporting_year"] == 2025) & (df["reporting_week_of_year"] == week)]
    gms_10 = float(df_10[df_10["merchant_customer_id"].isin(mcids_10)]["wtd_ord_gms"].sum())

    return {"week": week, "gms_20": gms_20, "gms_10": gms_10}


def load_cache() -> dict[int, dict]:
    """載入 cache,回傳 {week: {week, gms_20, gms_10}}。"""
    if CACHE_FILE.exists():
        data = json.loads(CACHE_FILE.read_text(encoding="utf-8"))
        return {d["week"]: d for d in data}
    return {}


def save_cache(cache: dict[int, dict]) -> None:
    CACHE_FILE.parent.mkdir(exist_ok=True)
    data = sorted(cache.values(), key=lambda d: d["week"])
    CACHE_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


@dataclass
class WeeklyGMSResult:
    weeks: list[int]
    gms_20: list[float]
    gms_10: list[float]


def build_weekly_gms(force_all: bool = False) -> WeeklyGMSResult:
    """建構每週 GMS 資料。如果有 cache 只跑最新一週,否則跑全部。"""
    print("[Weekly GMS] 取 MCID...")
    mcids_20 = load_20_mcids()
    mcids_10 = load_10_mcids()
    print(f"  2.0: {len(mcids_20)}, 1.0: {len(mcids_10)}")

    all_wks = find_all_wk_p0s(WBR_BASE)
    print(f"  找到 {len(all_wks)} 個有 P0 的週資料夾: wk{all_wks[0][0]}~wk{all_wks[-1][0]}")

    cache = load_cache() if not force_all else {}
    cached_weeks = set(cache.keys())
    all_week_nums = {wk for wk, _ in all_wks}

    # 決定要跑哪些週
    if not cached_weeks or force_all:
        to_run = all_wks
    else:
        # 只跑 cache 裡沒有的(通常是最新一週)
        to_run = [(wk, p) for wk, p in all_wks if wk not in cached_weeks]

    if to_run:
        print(f"  需要計算 {len(to_run)} 週: {[w for w, _ in to_run]}")
        for i, (wk, p0) in enumerate(to_run):
            print(f"    [{i+1}/{len(to_run)}] wk{wk}: {p0.name}...")
            try:
                result = calc_week(p0, wk, mcids_20, mcids_10)
                cache[wk] = result
                # 每週跑完就存一次 cache,避免中斷遺失
                save_cache(cache)
            except Exception as e:
                print(f"      [錯誤] wk{wk}: {e}")
                cache[wk] = {"week": wk, "gms_20": 0.0, "gms_10": 0.0}
        print(f"  Cache 已更新: {CACHE_FILE}")
    else:
        print("  所有週數已在 cache,無需重算")

    # 組裝結果
    weeks = sorted(cache.keys())
    gms_20 = [cache[w]["gms_20"] for w in weeks]
    gms_10 = [cache[w]["gms_10"] for w in weeks]
    return WeeklyGMSResult(weeks=weeks, gms_20=gms_20, gms_10=gms_10)


if __name__ == "__main__":
    result = build_weekly_gms(force_all=True)
    print(f"\n=== Weekly GMS ({len(result.weeks)} 週) ===")
    print(f"{'Week':>6} {'ACC 2.0':>12} {'ACC 1.0':>12}")
    for w, g20, g10 in zip(result.weeks, result.gms_20, result.gms_10):
        print(f"  wk{w:>2}  {g20:>12,.0f}  {g10:>12,.0f}")
