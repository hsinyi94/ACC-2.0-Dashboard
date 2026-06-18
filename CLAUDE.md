# ACC 2.0 Dashboard — Claude Code 指引

## 專案概述

這是 ACC 2.0 (Amazon Cross-border Commerce 2.0) 的資料分析 Dashboard。
每週產出 HTML 報表,部署到 GitHub Pages: https://hsinyi94.github.io/ACC-2.0-Dashboard/

## 快速指令

### 每週更新 (最常用)

```powershell
& "C:\Users\hsinyih\AppData\Local\Python\bin\python.exe" scripts/weekly_update.py
```

這會自動:
1. 執行 `build_html.py` 產出最新 HTML
2. 複製到 `docs/index.html`
3. `git add -A` → `git commit` → `git push`

### 單獨產出 HTML (不 push)

```powershell
& "C:\Users\hsinyih\AppData\Local\Python\bin\python.exe" scripts/build_html.py
```

### 強制重算所有 Weekly GMS Cache

```powershell
& "C:\Users\hsinyih\AppData\Local\Python\bin\python.exe" scripts/weekly_gms.py
```

## 環境

- Python: `C:\Users\hsinyih\AppData\Local\Python\bin\python.exe` (3.14.3)
- 套件: pandas 3.0.1, openpyxl 3.1.5
- 工作目錄: `C:\Users\hsinyih\AI related\ACC 2.0資料分析`
- Shell: PowerShell (Windows)

## 資料來源

所有 Excel 來源路徑都是硬寫在腳本中:

| 資料 | 路徑 |
|------|------|
| ACC 2.0 最終錄取結果 | `C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 2.0 最終錄取結果 (20260105).xlsx` |
| ACC 1.0 分析 | `C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0\ACC 1.0分析.xlsx` |
| WBR 週報 P0 | `C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wkXX\P0*.xlsx` |
| MBR 月報 P0 | `C:\Users\hsinyih\amazon.com\TWGS - 2026 MBR\X. Mon\P0*.xlsx` |
| NSR Launch Tracker (2026) | `C:\Users\hsinyih\amazon.com\TWGS - 2026 WBR\wkXX\NSR Launch Tracker*.xlsx` |
| NSR Launch Tracker (2025/1.0) | `W:\Team Spaces\TWGS\00_Commonly Accessed\2025 Weekly Business Report\wk52\NSR Launch Tracker_20251220.xlsx` |

**重要前提**: 執行前所有 Excel 檔案必須關閉 (否則 openpyxl 會讀取失敗)。

## 專案結構

```
scripts/
├── weekly_update.py       # 入口: 每週更新 (build → copy → git push)
├── build_html.py          # 主 HTML 產生器,整合所有分析模組
├── build_analysis.py      # 區塊1+2: 2.0 賣家分析 + 1.0 vs 2.0 對比
├── monthly_analysis.py    # 區塊3: 月度 GMS (ACC 1.0/2.0/NSR All)
├── ytd_analysis.py        # 區塊4: YTD 達標率
├── adoption_analysis.py   # 區塊5: Seller Adoption 指標
├── raw_data.py            # 區塊6: 每位賣家完整數據表
└── weekly_gms.py          # 區塊7: 每週 GMS 折線圖 (有 cache 機制)

output/
├── ACC_2.0_分析_YYYYMMDD.html   # 產出的 HTML 報表
├── weekly_gms_cache.json         # Weekly GMS cache (避免重複讀 P0)
└── *.xlsx                        # 歷史 Excel 輸出

docs/
└── index.html             # GitHub Pages 部署檔 (= 最新 HTML 的副本)
```

## Dashboard 區塊說明

| 區塊 | 內容 | 資料來源 |
|------|------|----------|
| 1 | ACC 2.0 錄取賣家分析 (n=115) | ACC 2.0 最終錄取結果 |
| 2 | ACC 1.0 vs 2.0 對比 | ACC 2.0 + ACC 1.0分析 |
| 3 | 月度 GMS 變化 | MBR P0 + ACC 2.0 (P24:Y29) |
| 4 | YTD 達標率 | WBR 最新週 P0 |
| 5 | Seller Adoption | NSR Launch Tracker |
| 6 | Raw Data (每位賣家) | 合併所有來源 |
| 7 | Weekly GMS 趨勢 | WBR 每週 P0 (cached) |

## 常見操作

### 修改 HTML 樣式
編輯 `scripts/build_html.py` 中的 `CSS` 變數和各 `render_*()` 函式。

### 新增分析欄位
1. 在對應的分析模組 (如 `build_analysis.py`) 加邏輯
2. 在 `build_html.py` 加渲染函式
3. 在 `build_html.py` 的 `main()` 中呼叫

### Weekly GMS Cache
- Cache 存在 `output/weekly_gms_cache.json`
- 正常每週更新只會計算新增的週數
- 若需要全部重算: `python scripts/weekly_gms.py` (直接執行 = force_all=True)

### P0 檔案定位邏輯
- WBR: `TWGS - 2026 WBR/wkXX/` 取最大週數、最新修改的 P0*.xlsx
- MBR: `TWGS - 2026 MBR/X. Mon/` 取最大月份編號的資料夾

## 注意事項

1. **Excel 必須關閉**: OneDrive 同步的 Excel 檔案如果正在開啟,openpyxl 會讀取到鎖定檔導致錯誤
2. **MCID 格式**: P0 中 `merchant_customer_id` 是 float (帶 .0),腳本會轉 int64 → str 對齊
3. **NSR 賣家數**: 2.0 固定用 115 (實際資料列=110,有 5 列空值),1.0 固定用 70
4. **國家合併**: AE+SA→MENA, UK+DE+FR→EU
5. **Git Push**: 預設 push 到 main branch,GitHub Pages 從 docs/ 部署
6. **W: drive**: 1.0 的 NSR Launch Tracker 在網路磁碟 W:,需連上公司網路才能存取

## 故障排除

| 問題 | 原因 | 解法 |
|------|------|------|
| `PermissionError` | Excel 檔案被開啟 | 關閉所有相關 Excel |
| `FileNotFoundError: P0` | 新週的 WBR 資料夾還沒放 P0 | 確認 wkXX 資料夾有 P0 |
| `KeyError: 欄位名` | P0 格式變更 | 檢查 P0 工作表名/欄位名是否改了 |
| Weekly GMS 數字異常 | Cache 過期 | 刪除 `output/weekly_gms_cache.json` 重跑 |
| W: drive 無法存取 | 未連公司網路 | 連上 VPN 或公司內網 |
