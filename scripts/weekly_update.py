"""每週更新腳本。

更新項目:
- 區塊 4: YTD 達標率 (最新 WBR P0)
- 區塊 5: Seller Adoption (最新 NSR Launch Tracker)
- Weekly GMS (最新一週的 wtd_ord_gms)

執行方式:
  python scripts/weekly_update.py

會自動:
1. 重新產出 HTML (只有上述三區塊的資料會更新,其他用 cache/既有資料)
2. 複製到 docs/index.html
3. git commit + push

前提:
- ACC 2.0 Excel 檔案要關閉
- WBR 最新週 P0 要關閉
- NSR Launch Tracker 要關閉
"""
import subprocess
import shutil
import sys
from pathlib import Path

SCRIPTS_DIR = Path(__file__).resolve().parent
ROOT_DIR = SCRIPTS_DIR.parent
OUTPUT_DIR = ROOT_DIR / "output"
DOCS_DIR = ROOT_DIR / "docs"

PYTHON = sys.executable


def run(cmd, **kwargs):
    print(f"  > {cmd}")
    import os
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    result = subprocess.run(cmd, shell=True, capture_output=True, text=True,
                           cwd=str(ROOT_DIR), encoding="utf-8", errors="replace",
                           env=env, **kwargs)
    if result.stdout:
        # 只印最後幾行
        lines = result.stdout.strip().split("\n")
        for line in lines[-10:]:
            print(f"    {line}")
    if result.returncode != 0:
        print(f"  [ERROR] {result.stderr[-500:]}")
        return False
    return True


def main():
    print("=" * 60)
    print("ACC 2.0 Dashboard - 每週更新")
    print("=" * 60)

    # Step 1: 產出 HTML
    print("\n[Step 1] 產出 Dashboard HTML...")
    ok = run(f'"{PYTHON}" scripts/build_html.py')
    if not ok:
        print("\n產出失敗!請確認所有 Excel 檔案已關閉。")
        return

    # Step 2: 找最新的 HTML 輸出
    html_files = sorted(OUTPUT_DIR.glob("ACC_2.0_分析_*.html"))
    if not html_files:
        print("找不到輸出 HTML!")
        return
    latest = html_files[-1]
    print(f"\n[Step 2] 最新輸出: {latest.name}")

    # Step 3: 複製到 docs/index.html
    DOCS_DIR.mkdir(exist_ok=True)
    dest = DOCS_DIR / "index.html"
    shutil.copy2(latest, dest)
    print(f"  複製到 {dest}")

    # Step 4: Git commit + push
    print("\n[Step 3] Git commit & push...")
    import datetime
    today = datetime.date.today().strftime("%Y-%m-%d")
    run("git add -A")
    run(f'git commit -m "Weekly update {today}"')
    run("git push")

    print("\n" + "=" * 60)
    print("更新完成!")
    print("GitHub Pages 會在 1-2 分鐘後自動部署。")
    print("連結: https://hsinyi94.github.io/ACC-2.0-Dashboard/")
    print("=" * 60)


if __name__ == "__main__":
    main()
