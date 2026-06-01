"""找出 TWGS - 2026 WBR 底下最新一週的 P0 報表。

規則:
- 掃資料夾裡所有 wk<數字> 的子資料夾 (不分大小寫)
- 取週數最大的那個
- 從裡面挑檔名以 P0 開頭的 .xlsx (不分大小寫)
- 若有多個 P0 檔,回傳最新修改的那一個
"""
import re
from pathlib import Path

BASE = Path(r"C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR")

def find_latest_week_folder(base: Path) -> Path | None:
    pattern = re.compile(r"^wk(\d+)$", re.IGNORECASE)
    candidates = []
    for child in base.iterdir():
        if not child.is_dir():
            continue
        m = pattern.match(child.name)
        if m:
            candidates.append((int(m.group(1)), child))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return candidates[-1][1]

def find_p0_files(folder: Path) -> list[Path]:
    files = [
        f for f in folder.iterdir()
        if f.is_file()
        and f.suffix.lower() == ".xlsx"
        and f.name.lower().startswith("p0")
    ]
    # 最新修改排最前面
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files


def find_latest_p0_file(folder: Path) -> Path | None:
    files = find_p0_files(folder)
    return files[0] if files else None

def main() -> None:
    print(f"掃描基底資料夾: {BASE}")
    if not BASE.exists():
        print("  基底資料夾不存在")
        return

    latest = find_latest_week_folder(BASE)
    if latest is None:
        print("  找不到任何 wk<數字> 資料夾")
        return
    print(f"最新週次資料夾: {latest.name}")

    p0_files = find_p0_files(latest)
    if not p0_files:
        print("  該資料夾裡沒有 P0 開頭的 xlsx")
        return

    from datetime import datetime
    newest = p0_files[0]
    size_mb = newest.stat().st_size / (1024 * 1024)
    mstr = datetime.fromtimestamp(newest.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
    print(f"最新 P0 檔案: {newest.name}  ({size_mb:.1f} MB, 修改: {mstr})")
    if len(p0_files) > 1:
        print(f"(同資料夾共 {len(p0_files)} 個 P0 檔,挑最新的)")

if __name__ == "__main__":
    main()
