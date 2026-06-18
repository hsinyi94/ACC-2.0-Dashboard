"""產出 ACC 2.0 分析表的 HTML 版本。

沿用 build_analysis.py 的資料擷取函式,重新組成有設計感的 HTML。
設計特色:
- 頂部 sticky 導覽列,點擊跳到對應項目
- KPI 卡片總覽 (n=115 / 1.0 n=70 / US vs Expansion 比例)
- 每個項目一張卡,表格帶 CSS 長條圖
- Δ 百分點紅綠標示、向上/向下箭頭
- 中文字型最佳化 (Noto Sans TC)
"""
from __future__ import annotations

import datetime as dt
import html
from pathlib import Path

import pandas as pd

from build_analysis import (
    COMPARISON_MAP,
    FILE_10,
    FILE_20,
    count_10_brand_age,
    count_10_country,
    count_10_us_vs_expansion,
    count_column,
    get_20_blocks,
    get_recruit_country,
    load_10_roster,
)
from monthly_analysis import MONTH_ABBR, build_monthly_result
from ytd_analysis import build_ytd_analysis
from adoption_analysis import build_adoption_analysis, ADOPTION_COLS, ADOPTION_LABELS
from raw_data import build_raw_data
from weekly_gms import build_weekly_gms

OUT_DIR = Path(__file__).resolve().parent.parent / "output"
OUT_DIR.mkdir(exist_ok=True)
OUT_FILE = OUT_DIR / f"ACC_2.0_分析_{dt.date.today().strftime('%Y%m%d')}.html"


CSS = """
:root {
  --c-primary: #1f4e79;
  --c-primary-soft: #2e75b6;
  --c-accent: #e07b00;
  --c-bg: #f4f6fa;
  --c-card: #ffffff;
  --c-text: #1f2937;
  --c-muted: #6b7280;
  --c-border: #e5e7eb;
  --c-bar-20: #2e75b6;
  --c-bar-10: #9dc3e6;
  --c-up: #2f9e44;
  --c-down: #e03131;
  --c-us: #2e75b6;
  --c-expansion: #e07b00;
  --shadow-sm: 0 1px 2px rgba(0,0,0,0.05);
  --shadow-md: 0 4px 14px rgba(0,0,0,0.08);
}

* { box-sizing: border-box; }

body {
  margin: 0;
  font-family: "Noto Sans TC", "Microsoft JhengHei", "PingFang TC", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  background: var(--c-bg);
  color: var(--c-text);
  line-height: 1.6;
  -webkit-font-smoothing: antialiased;
}

.hero {
  background: linear-gradient(135deg, #1f4e79 0%, #2e75b6 100%);
  color: white;
  padding: 48px 32px 36px;
  box-shadow: var(--shadow-md);
}

.hero h1 {
  margin: 0 0 8px 0;
  font-size: 28px;
  font-weight: 700;
  letter-spacing: 0.5px;
}

.hero .meta {
  font-size: 14px;
  opacity: 0.88;
}

.kpi-row {
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 14px;
  margin-top: 22px;
  max-width: 1200px;
}

.hero-content {
  display: flex;
  align-items: flex-start;
  gap: 24px;
  margin-top: 22px;
}

.hero-content .kpi-row {
  margin-top: 0;
  flex: 1;
}

.snapshot-box {
  background: rgba(255,255,255,0.12);
  backdrop-filter: blur(8px);
  border: 1px solid rgba(255,255,255,0.2);
  border-radius: 10px;
  padding: 16px 20px;
  min-width: 380px;
  max-width: 420px;
  flex-shrink: 0;
}

.snapshot-title {
  font-size: 13px;
  font-weight: 700;
  margin-bottom: 10px;
  color: rgba(255,255,255,0.95);
  line-height: 1.4;
}

.snapshot-list {
  margin: 0;
  padding-left: 18px;
  font-size: 12.5px;
  color: rgba(255,255,255,0.88);
  line-height: 1.8;
}

.snapshot-list li {
  margin-bottom: 2px;
}

.snapshot-list strong {
  color: rgba(255,255,255,0.95);
}

.kpi {
  background: rgba(255,255,255,0.14);
  backdrop-filter: blur(8px);
  border-radius: 10px;
  padding: 14px 18px;
  border: 1px solid rgba(255,255,255,0.2);
}

.kpi .label {
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 1px;
  opacity: 0.82;
  margin-bottom: 4px;
}

.kpi .value {
  font-size: 24px;
  font-weight: 700;
}

.kpi .sub {
  font-size: 12px;
  opacity: 0.78;
  margin-top: 2px;
}

nav.toc {
  position: sticky;
  top: 0;
  z-index: 10;
  background: rgba(255,255,255,0.97);
  backdrop-filter: blur(10px);
  border-bottom: 1px solid var(--c-border);
  padding: 0 32px;
  box-shadow: var(--shadow-sm);
  display: flex;
  align-items: flex-end;
  gap: 4px;
  overflow-x: auto;
}

nav.toc button.tab {
  background: none;
  border: none;
  padding: 14px 22px 12px;
  font-size: 14px;
  font-weight: 600;
  color: var(--c-muted);
  cursor: pointer;
  border-bottom: 3px solid transparent;
  transition: all 0.15s;
  font-family: inherit;
  white-space: nowrap;
}

nav.toc button.tab:hover {
  color: var(--c-primary);
  background: #f8fafc;
}

nav.toc button.tab.active {
  color: var(--c-primary);
  border-bottom-color: var(--c-accent);
  background: #fff;
}

.tab-panel {
  display: none;
  animation: fadeIn 0.25s ease;
}

.tab-panel.active { display: block; }

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(4px); }
  to { opacity: 1; transform: translateY(0); }
}

main {
  max-width: 1200px;
  margin: 0 auto;
  padding: 28px 32px 60px;
}

.section-title {
  font-size: 22px;
  font-weight: 700;
  color: var(--c-primary);
  margin: 36px 0 18px;
  padding-left: 14px;
  border-left: 5px solid var(--c-accent);
}

.section-title:first-child {
  margin-top: 0;
}

.card {
  background: var(--c-card);
  border-radius: 12px;
  box-shadow: var(--shadow-sm);
  margin-bottom: 18px;
  overflow: hidden;
  border: 1px solid var(--c-border);
  scroll-margin-top: 70px;
}

.card-header {
  padding: 14px 22px;
  background: linear-gradient(90deg, #eef3fb 0%, #f8fafc 100%);
  border-bottom: 1px solid var(--c-border);
  font-size: 16px;
  font-weight: 600;
  color: var(--c-primary);
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.card-header .total-pill {
  font-size: 12px;
  font-weight: 500;
  background: var(--c-primary);
  color: white;
  padding: 3px 10px;
  border-radius: 999px;
}

table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
}

thead th {
  background: #f8fafc;
  color: var(--c-muted);
  font-weight: 600;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.4px;
  padding: 10px 16px;
  text-align: left;
  border-bottom: 2px solid var(--c-border);
}

thead th.num { text-align: right; }
thead th.bar-col { width: 32%; }

tbody td {
  padding: 10px 16px;
  border-bottom: 1px solid #f1f3f5;
  vertical-align: middle;
}

tbody tr:last-child td {
  border-bottom: none;
}

tbody tr:hover {
  background: #fafbfd;
}

td.num {
  text-align: right;
  font-variant-numeric: tabular-nums;
  font-weight: 500;
}

td.label {
  font-weight: 500;
}

.bar-wrap {
  position: relative;
  height: 8px;
  background: #f1f3f5;
  border-radius: 999px;
  overflow: hidden;
}

.bar {
  position: absolute;
  top: 0; left: 0; bottom: 0;
  background: var(--c-bar-20);
  border-radius: 999px;
  transition: width 0.4s ease;
}

.bar.bar-10 { background: var(--c-bar-10); }
.bar.bar-us { background: var(--c-us); }
.bar.bar-expansion { background: var(--c-expansion); }

/* 並排雙 bar (1.0 vs 2.0) */
.dual-bar {
  display: flex;
  flex-direction: column;
  gap: 3px;
}
.dual-bar .bar-row {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 11px;
  color: var(--c-muted);
}
.dual-bar .bar-row .tag {
  width: 24px;
  flex-shrink: 0;
  font-weight: 600;
}
.dual-bar .bar-row .bar-wrap {
  flex: 1;
  height: 6px;
}

/* AE 細分表樣式 */
.country-row {
  background: #fafbfd;
}
.country-row td.label {
  font-weight: 700;
  color: var(--c-primary);
}
.subtotal-row td {
  background: #fff4e6;
  font-weight: 700;
  border-top: 1px solid #ffd8a8;
}

/* US vs Expansion 附加列 */
.expansion-section {
  background: linear-gradient(180deg, #f0f9ff 0%, #e0f2fe 100%);
  padding: 16px 22px;
  border-top: 2px solid #bae6fd;
  margin-top: 4px;
}
.expansion-section h4 {
  margin: 0 0 10px 0;
  font-size: 13px;
  color: var(--c-primary);
  text-transform: uppercase;
  letter-spacing: 1px;
}
.exp-row {
  display: flex;
  align-items: center;
  gap: 14px;
  padding: 8px 0;
}
.exp-row .exp-label {
  width: 180px;
  font-weight: 600;
}
.exp-row .exp-num {
  width: 80px;
  font-size: 18px;
  font-weight: 700;
  font-variant-numeric: tabular-nums;
}
.exp-row .exp-pct {
  width: 60px;
  color: var(--c-muted);
  font-variant-numeric: tabular-nums;
}
.exp-row .exp-bar-wrap {
  flex: 1;
  height: 12px;
  background: rgba(255,255,255,0.7);
  border-radius: 999px;
  overflow: hidden;
}
.exp-row .exp-bar {
  height: 100%;
  border-radius: 999px;
}

/* 對比表 Δ 欄 */
.delta {
  font-weight: 600;
  font-variant-numeric: tabular-nums;
}
.delta.up { color: var(--c-up); }
.delta.down { color: var(--c-down); }
.delta.zero { color: var(--c-muted); }

.delta::before {
  display: inline-block;
  width: 12px;
  text-align: center;
  margin-right: 2px;
}
.delta.up::before { content: "▲"; font-size: 10px; }
.delta.down::before { content: "▼"; font-size: 10px; }

/* 對比表欄位 */
.compare-10 { background: #f8fafc; }

.note-box {
  background: #fffbeb;
  border: 1px solid #fcd34d;
  border-radius: 10px;
  padding: 16px 22px;
  margin-top: 24px;
  font-size: 13px;
  color: #78350f;
}
.note-box h4 {
  margin: 0 0 8px 0;
  color: #b45309;
}
.note-box ul {
  margin: 0;
  padding-left: 20px;
}
.note-box li {
  margin-bottom: 4px;
}

footer {
  text-align: center;
  color: var(--c-muted);
  font-size: 12px;
  padding: 20px 0 30px;
}

@media (max-width: 720px) {
  .hero { padding: 28px 18px 22px; }
  .hero-content { flex-direction: column; }
  .snapshot-box { min-width: unset; max-width: 100%; }
  .kpi-row { grid-template-columns: repeat(auto-fit, minmax(140px, 1fr)); }
  main { padding: 20px 16px 40px; }
  nav.toc { padding: 8px 16px; }
  thead th, tbody td { padding: 8px 10px; }
  thead th.bar-col { display: none; }
  tbody td.bar-col { display: none; }
}

/* === 區塊 3: 月度 GMS === */
.monthly-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
  font-variant-numeric: tabular-nums;
}
.monthly-table th, .monthly-table td {
  padding: 10px 12px;
  border-bottom: 1px solid #f1f3f5;
  text-align: right;
}
.monthly-table th {
  background: #f8fafc;
  color: var(--c-muted);
  font-weight: 600;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.4px;
  text-align: center;
}
.monthly-table th:first-child,
.monthly-table td:first-child {
  text-align: left;
  font-weight: 600;
  color: var(--c-text);
  background: #f8fafc;
  position: sticky;
  left: 0;
  z-index: 1;
}
.monthly-table tr.row-10 td:first-child { color: #1e3a5f; }
.monthly-table tr.row-20 td:first-child { color: #c76a00; }
.monthly-table tr.row-mom td { color: var(--c-muted); font-size: 12px; font-style: italic; }
.monthly-table tr.row-mom td:first-child { color: var(--c-muted); font-style: normal; }
.monthly-table tr.row-vs td { color: var(--c-muted); font-size: 12px; font-style: italic; }
.monthly-table tr.row-vs td:first-child { color: var(--c-muted); font-style: normal; }
.monthly-table tr.row-nsr td:first-child { color: #495057; }
.monthly-table tr.row-pct td { color: var(--c-muted); font-size: 12px; font-style: italic; }
.monthly-table tr.row-pct td:first-child { color: var(--c-muted); font-style: normal; }
.monthly-table td.no-data { color: #c0c4cc; }
.monthly-table td.value { font-weight: 500; }

/* 合計欄 (最後一欄) */
.monthly-table th:last-child,
.monthly-table td:last-child {
  background: #eef3fb;
  border-left: 2px solid var(--c-primary-soft);
  font-weight: 700;
}
.monthly-table tr.row-mom td:last-child,
.monthly-table tr.row-vs td:last-child,
.monthly-table tr.row-pct td:last-child {
  background: #f4f6fa;
}

.mom-up { color: var(--c-up); font-weight: 600; }
.mom-up::before { content: "▲ "; font-size: 9px; }
.mom-down { color: var(--c-down); font-weight: 600; }
.mom-down::before { content: "▼ "; font-size: 9px; }

.chart-wrap {
  padding: 22px 22px 10px;
  background: #fafbfd;
  border-top: 1px solid var(--c-border);
}
.chart-wrap h4 {
  margin: 0 0 10px 0;
  font-size: 13px;
  color: var(--c-muted);
  font-weight: 600;
  text-transform: uppercase;
  letter-spacing: 0.8px;
}
.chart-legend {
  display: flex;
  gap: 20px;
  font-size: 12px;
  color: var(--c-muted);
  margin-top: 6px;
}
.chart-legend .legend-item {
  display: flex;
  align-items: center;
  gap: 6px;
}
.chart-legend .legend-swatch {
  width: 14px;
  height: 3px;
  border-radius: 2px;
}
.chart-source {
  margin: 10px 22px 16px;
  font-size: 12px;
  color: var(--c-muted);
  background: #f8fafc;
  padding: 8px 12px;
  border-left: 3px solid var(--c-primary-soft);
  border-radius: 4px;
}

/* === 區塊 4: YTD 達標率 === */
.ytd-group {
  margin-bottom: 22px;
}
.ytd-group-title {
  font-size: 14px;
  font-weight: 700;
  color: var(--c-primary);
  background: #eef3fb;
  padding: 10px 22px;
  border-bottom: 1px solid var(--c-border);
  letter-spacing: 0.5px;
}
.ytd-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 14px;
  font-variant-numeric: tabular-nums;
}
.ytd-table th, .ytd-table td {
  padding: 12px 14px;
  border-bottom: 1px solid #f1f3f5;
  text-align: right;
}
.ytd-table th {
  background: #f8fafc;
  color: var(--c-muted);
  font-weight: 600;
  font-size: 12px;
  text-transform: uppercase;
  letter-spacing: 0.4px;
  text-align: center;
  white-space: nowrap;
}
.ytd-table th:first-child,
.ytd-table td:first-child {
  text-align: left;
  font-weight: 600;
}
.ytd-table tr.track-20 td:first-child { color: #c76a00; }
.ytd-table tr.track-nsr td:first-child { color: #495057; }
.ytd-table td.attain {
  font-weight: 700;
  background: #eef3fb;
  border-left: 2px solid var(--c-primary-soft);
}
.ytd-table td.attain.hit { color: var(--c-up); }
.ytd-table td.attain.miss { color: var(--c-down); }
.ytd-table td.no-data { color: #c0c4cc; }

/* KPI 卡 (YTD 區塊頂部) */
.ytd-kpi-row {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
  gap: 14px;
  padding: 22px;
  background: linear-gradient(135deg, #eef3fb 0%, #f8fafc 100%);
  border-bottom: 1px solid var(--c-border);
}
.ytd-kpi {
  background: white;
  border-radius: 10px;
  padding: 14px 18px;
  border: 1px solid var(--c-border);
  box-shadow: var(--shadow-sm);
}
.ytd-kpi .kpi-label {
  font-size: 11px;
  text-transform: uppercase;
  color: var(--c-muted);
  letter-spacing: 0.8px;
  margin-bottom: 4px;
}
.ytd-kpi .kpi-value {
  font-size: 22px;
  font-weight: 700;
  color: var(--c-primary);
}
.ytd-kpi .kpi-sub {
  font-size: 11px;
  color: var(--c-muted);
  margin-top: 2px;
}

/* 達標率進度條 */
.attain-bar {
  display: inline-block;
  width: 100%;
  max-width: 80px;
  height: 6px;
  background: #e5e7eb;
  border-radius: 999px;
  overflow: hidden;
  vertical-align: middle;
  margin-left: 6px;
}
.attain-bar-fill {
  height: 100%;
  border-radius: 999px;
}
"""


def esc(s) -> str:
    if s is None:
        return ""
    return html.escape(str(s))


def pct(p: float) -> str:
    return f"{p * 100:.1f}%"


def bar_html(pct_val: float, cls: str = "") -> str:
    w = max(0.0, min(1.0, pct_val)) * 100
    return f'<div class="bar-wrap"><div class="bar {cls}" style="width:{w:.1f}%"></div></div>'


def dual_bar_html(p10: float, p20: float) -> str:
    w10 = max(0.0, min(1.0, p10)) * 100
    w20 = max(0.0, min(1.0, p20)) * 100
    return (
        '<div class="dual-bar">'
        f'<div class="bar-row"><span class="tag">1.0</span>'
        f'<div class="bar-wrap"><div class="bar bar-10" style="width:{w10:.1f}%"></div></div></div>'
        f'<div class="bar-row"><span class="tag">2.0</span>'
        f'<div class="bar-wrap"><div class="bar" style="width:{w20:.1f}%"></div></div></div>'
        '</div>'
    )


def delta_html(diff_pp: float) -> str:
    """diff_pp 是百分點差,轉為 bps (×100) 顯示。"""
    diff_bps = diff_pp * 100
    if abs(diff_bps) < 5:
        return f'<span class="delta zero">{diff_bps:+,.0f} bps</span>'
    cls = "up" if diff_bps > 0 else "down"
    return f'<span class="delta {cls}">{abs(diff_bps):,.0f} bps</span>'


def anchor(item_name: str, prefix: str) -> str:
    safe = item_name.replace(" ", "-").replace("(", "").replace(")", "").replace("/", "-")
    return f"{prefix}-{safe}"


# ============================================================
# 區塊 1 渲染
# ============================================================

def render_section1(blocks: dict, ae_rows, us_exp_rows) -> str:
    parts = ['<h2 class="section-title" id="section1">ACC 2.0 錄取賣家分析</h2>']

    # --- 錄取國家卡片 ---
    parts.append(render_country_card(ae_rows, us_exp_rows))

    # --- 其餘項目 ---
    order = [
        "公司類型",
        "公司所在區域",
        "Category",
        "公司產品類型",
        "公司主要經營型態",
        "是否有國內電商銷售經驗",
        "品牌創立年限",
    ]
    for name in order:
        rows = blocks.get(name, [])
        parts.append(render_simple_card(name, rows, anchor_id=anchor(name, "s1")))
    return "\n".join(parts)


def render_country_card(ae_rows, us_exp_rows) -> str:
    total = sum(cnt for country, bd, cnt, pct_ in ae_rows if country.endswith("小計"))
    rows_html = []
    for country, bd, cnt, pct_ in ae_rows:
        is_subtotal = country.endswith("小計")
        tr_cls = "subtotal-row" if is_subtotal else ("country-row" if country and not is_subtotal else "")
        country_cell = esc(country)
        rows_html.append(
            f'<tr class="{tr_cls}">'
            f'<td class="label">{country_cell}</td>'
            f'<td>{esc(bd)}</td>'
            f'<td class="num">{cnt}</td>'
            f'<td class="num">{pct(pct_)}</td>'
            f'<td class="bar-col">{bar_html(pct_)}</td>'
            '</tr>'
        )
    table_html = (
        '<table>'
        '<thead><tr>'
        '<th>國家</th><th>AE</th>'
        '<th class="num">人數</th><th class="num">佔比</th>'
        '<th class="bar-col">分布</th>'
        '</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        '</table>'
    )

    # US vs Expansion 附加區
    exp_rows = []
    for label, cnt, pct_ in us_exp_rows:
        bar_cls = "bar-us" if label == "US" else "bar-expansion"
        color = "var(--c-us)" if label == "US" else "var(--c-expansion)"
        exp_rows.append(
            '<div class="exp-row">'
            f'<span class="exp-label">{esc(label)}</span>'
            f'<span class="exp-num">{cnt}</span>'
            f'<span class="exp-pct">{pct(pct_)}</span>'
            '<div class="exp-bar-wrap">'
            f'<div class="exp-bar" style="width:{pct_*100:.1f}%;background:{color}"></div>'
            '</div>'
            '</div>'
        )
    exp_section = (
        '<div class="expansion-section">'
        '<h4>附加 · US vs Expansion</h4>'
        + "".join(exp_rows)
        + '</div>'
    )

    return (
        f'<section class="card" id="{anchor("錄取國家", "s1")}">'
        '<div class="card-header">'
        '<span>錄取國家 (依 AE)</span>'
        f'<span class="total-pill">Total {total}</span>'
        '</div>'
        f'{table_html}'
        f'{exp_section}'
        '</section>'
    )


def render_simple_card(name: str, rows: list[tuple[str, int, float]], anchor_id: str) -> str:
    total = sum(c for _, c, _ in rows)
    rows_html = []
    for label, cnt, pct_ in rows:
        rows_html.append(
            '<tr>'
            f'<td class="label">{esc(label)}</td>'
            f'<td class="num">{cnt}</td>'
            f'<td class="num">{pct(pct_)}</td>'
            f'<td class="bar-col">{bar_html(pct_)}</td>'
            '</tr>'
        )
    return (
        f'<section class="card" id="{anchor_id}">'
        '<div class="card-header">'
        f'<span>{esc(name)}</span>'
        f'<span class="total-pill">Total {total}</span>'
        '</div>'
        '<table>'
        '<thead><tr><th>類別</th>'
        '<th class="num">人數</th><th class="num">佔比</th>'
        '<th class="bar-col">分布</th></tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        '</table>'
        '</section>'
    )


# ============================================================
# 區塊 2 渲染 (對比)
# ============================================================

def render_section2(blocks_20, df_10, n10: int) -> str:
    parts = [f'<h2 class="section-title" id="section2">ACC 1.0 vs ACC 2.0 對比 (1.0 n={n10} | 2.0 n=115)</h2>']

    # 國家對比
    parts.append(render_country_compare_card(df_10, blocks_20))

    # 其他項目
    for item_20, col_10, custom in COMPARISON_MAP:
        if col_10 not in df_10.columns:
            continue
        rows_20 = blocks_20.get(item_20, [])
        if custom == "brand_age":
            rows_10 = count_10_brand_age(df_10, col_10)
        else:
            rows_10 = count_column(df_10, col_10)
        parts.append(render_compare_card(item_20, rows_10, rows_20, anchor_id=anchor(item_20, "s2")))

    # 備註
    parts.append(render_notes())
    return "\n".join(parts)


def render_country_compare_card(df_10, blocks_20) -> str:
    ae_rows, us_exp_20 = get_recruit_country()
    country_20 = {country.replace(" 小計", ""): (cnt, pct_)
                  for country, bd, cnt, pct_ in ae_rows
                  if country.endswith("小計")}
    country_10 = {k: (c, p) for k, c, p in count_10_country(df_10)}
    us_exp_10 = count_10_us_vs_expansion(df_10)
    map10 = {k: (c, p) for k, c, p in us_exp_10}
    map20 = {k: (c, p) for k, c, p in us_exp_20}

    rows_html = []
    for country in ["US", "EU", "JP", "MENA", "SG"]:
        c10, p10 = country_10.get(country, (0, 0))
        c20, p20 = country_20.get(country, (0, 0))
        diff = (p20 - p10) * 100
        rows_html.append(_compare_row(country, c10, p10, c20, p20, diff))
    for extra in sorted(set(country_10) - {"US", "EU", "JP", "MENA", "AU", "SG"}):
        c10, p10 = country_10[extra]
        rows_html.append(_compare_row(extra, c10, p10, 0, 0, (0 - p10) * 100))

    # US vs Expansion
    exp_rows_html = []
    for label in ["US", "Expansion (non-US)"]:
        c10, p10 = map10.get(label, (0, 0))
        c20, p20 = map20.get(label, (0, 0))
        diff = (p20 - p10) * 100
        exp_rows_html.append(_compare_row(label, c10, p10, c20, p20, diff, highlight=True))

    table_html = (
        '<table>'
        '<thead><tr>'
        '<th>類別</th>'
        '<th class="num compare-10">1.0 人數</th><th class="num compare-10">1.0 %</th>'
        '<th class="num">2.0 人數</th><th class="num">2.0 %</th>'
        '<th class="num">Δ</th>'
        '<th class="bar-col">分布對比</th>'
        '</tr></thead>'
        f'<tbody>{"".join(rows_html)}'
        '<tr><td colspan="7" style="padding:10px 16px;background:#fff4e6;font-size:12px;color:#78350f;font-weight:600;letter-spacing:0.5px;">US vs Expansion (互斥分類)</td></tr>'
        f'{"".join(exp_rows_html)}'
        '</tbody></table>'
    )

    return (
        f'<section class="card" id="{anchor("錄取國家", "s2")}">'
        '<div class="card-header"><span>錄取國家 (國家總計) + US vs Expansion</span></div>'
        f'{table_html}'
        '</section>'
    )


def render_compare_card(name: str, rows_10, rows_20, anchor_id: str) -> str:
    map10 = {k: (c, p) for k, c, p in rows_10}
    map20 = {k: (c, p) for k, c, p in rows_20}
    # 以 2.0 為主序,再補上 1.0 專有的
    all_labels = list(map20.keys())
    for k in map10:
        if k not in all_labels:
            all_labels.append(k)

    rows_html = []
    for label in all_labels:
        c10, p10 = map10.get(label, (0, 0))
        c20, p20 = map20.get(label, (0, 0))
        diff = (p20 - p10) * 100
        rows_html.append(_compare_row(label, c10, p10, c20, p20, diff))

    table_html = (
        '<table>'
        '<thead><tr>'
        '<th>類別</th>'
        '<th class="num compare-10">1.0 人數</th><th class="num compare-10">1.0 %</th>'
        '<th class="num">2.0 人數</th><th class="num">2.0 %</th>'
        '<th class="num">Δ</th>'
        '<th class="bar-col">分布對比</th>'
        '</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        '</table>'
    )
    return (
        f'<section class="card" id="{anchor_id}">'
        f'<div class="card-header"><span>{esc(name)}</span></div>'
        f'{table_html}'
        '</section>'
    )


def _compare_row(label, c10, p10, c20, p20, diff, highlight=False):
    tr_cls = "subtotal-row" if highlight else ""
    return (
        f'<tr class="{tr_cls}">'
        f'<td class="label">{esc(label)}</td>'
        f'<td class="num compare-10">{c10}</td>'
        f'<td class="num compare-10">{pct(p10)}</td>'
        f'<td class="num">{c20}</td>'
        f'<td class="num">{pct(p20)}</td>'
        f'<td class="num">{delta_html(diff)}</td>'
        f'<td class="bar-col">{dual_bar_html(p10, p20)}</td>'
        '</tr>'
    )


# ============================================================
# 區塊 3 渲染 (月度 GMS)
# ============================================================

def _fmt_num(v) -> str:
    if v is None:
        return '<span class="no-data">—</span>'
    return f"{v:,.0f}"


def _fmt_pct(v, colorize: bool = False, signed: bool = True) -> str:
    """格式化百分比。
    colorize=True: MoM 樣式 (紅/綠箭頭 + 絕對值)
    signed=True:  加 +/- 號 (適合變化量)
    signed=False: 不加 + (適合占比)
    """
    if v is None:
        return '<span class="no-data">—</span>'
    if colorize:
        if abs(v) < 0.0005:
            return f"{v * 100:.1f}%"
        cls = "mom-up" if v > 0 else "mom-down"
        return f'<span class="{cls}">{abs(v) * 100:.1f}%</span>'
    if signed:
        sign = "+" if v > 0 else ""
        return f"{sign}{v * 100:.1f}%"
    return f"{v * 100:.1f}%"


def render_monthly_table(result) -> str:
    """月度 GMS 表格,最後一欄為「有資料月份合計」。"""
    months = MONTH_ABBR
    header = (
        '<tr><th>GMS (USD)</th>'
        + "".join(f'<th>{m}.</th>' for m in months)
        + '<th>合計</th>'
        + '</tr>'
    )

    def row_html(label, values, cls, fmt_fn, include_total=True):
        cells = [f'<td>{esc(label)}</td>']
        for m in months:
            v = values.get(m)
            css = "value" if v is not None else "no-data"
            cells.append(f'<td class="{css}">{fmt_fn(v)}</td>')
        # 合計欄
        if include_total:
            nums = [v for v in values.values() if v is not None]
            total = sum(nums) if nums else None
            css = "value" if total is not None else "no-data"
            cells.append(f'<td class="{css}" style="font-weight:700;">{fmt_fn(total)}</td>')
        else:
            cells.append('<td class="no-data">—</td>')
        return f'<tr class="{cls}">' + "".join(cells) + '</tr>'

    body = (
        row_html("ACC 1.0", result.row_10.values, "row-10", _fmt_num)
        + row_html("ACC 2.0", result.row_20.values, "row-20", _fmt_num)
        + row_html("2.0 vs 1.0", result.row_vs_10.values, "row-vs",
                   lambda v: _fmt_pct(v, colorize=True), include_total=False)
        + row_html("MoM (2.0)", result.row_mom.values, "row-mom",
                   lambda v: _fmt_pct(v, colorize=True), include_total=False)
        + row_html("2026 NSR All", result.row_nsr.values, "row-nsr", _fmt_num)
        + row_html("% of NSR all", result.row_pct.values, "row-pct",
                   lambda v: _fmt_pct(v, colorize=False, signed=False), include_total=False)
    )
    return (
        '<div style="overflow-x:auto;">'
        '<table class="monthly-table">'
        f'<thead>{header}</thead>'
        f'<tbody>{body}</tbody>'
        '</table></div>'
    )


def render_monthly_chart(result) -> str:
    """三條線折線圖 (SVG),自動縮放 Y 軸。"""
    months = MONTH_ABBR
    series = [
        ("ACC 1.0", result.row_10.values, "#1f4e79"),
        ("ACC 2.0", result.row_20.values, "#e07b00"),
        ("2026 NSR All", result.row_nsr.values, "#6c757d"),
    ]
    # 找 max 值做 Y 軸縮放
    all_vals = [v for _, vals, _ in series for v in vals.values() if v is not None]
    if not all_vals:
        return '<div class="chart-wrap"><h4>月度趨勢</h4><p>無資料</p></div>'
    y_max = max(all_vals) * 1.1

    # SVG 參數
    width, height = 860, 240
    pad_l, pad_r, pad_t, pad_b = 50, 20, 20, 30
    plot_w = width - pad_l - pad_r
    plot_h = height - pad_t - pad_b
    n = len(months)
    x_step = plot_w / (n - 1)

    def pt(i, v):
        x = pad_l + i * x_step
        y = pad_t + plot_h * (1 - v / y_max)
        return x, y

    # Y 軸格線 (5 格)
    grid_lines = []
    for i in range(5):
        ratio = i / 4
        y = pad_t + plot_h * (1 - ratio)
        label = y_max * ratio
        label_txt = f"{label / 1000:.0f}K" if label < 1_000_000 else f"{label / 1_000_000:.1f}M"
        grid_lines.append(
            f'<line x1="{pad_l}" y1="{y:.1f}" x2="{width - pad_r}" y2="{y:.1f}" '
            'stroke="#e5e7eb" stroke-width="1" stroke-dasharray="2,3"/>'
        )
        grid_lines.append(
            f'<text x="{pad_l - 6}" y="{y + 4:.1f}" text-anchor="end" '
            'font-size="10" fill="#9ca3af">' + label_txt + '</text>'
        )

    # X 軸標籤
    x_labels = []
    for i, m in enumerate(months):
        x = pad_l + i * x_step
        x_labels.append(
            f'<text x="{x:.1f}" y="{height - pad_b + 18}" text-anchor="middle" '
            'font-size="11" fill="#6b7280">' + m + '</text>'
        )

    # 三條線 + 圓點
    lines_svg = []
    for label, values, color in series:
        pts = []
        for i, m in enumerate(months):
            v = values.get(m)
            if v is None:
                continue
            x, y = pt(i, v)
            pts.append((x, y))
        if len(pts) < 2:
            # 只有一個點就畫圓
            for x, y in pts:
                lines_svg.append(
                    f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="{color}"/>'
                )
            continue
        d = "M " + " L ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
        lines_svg.append(
            f'<path d="{d}" fill="none" stroke="{color}" stroke-width="2.2" '
            'stroke-linecap="round" stroke-linejoin="round"/>'
        )
        for x, y in pts:
            lines_svg.append(
                f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="{color}" stroke="#fff" stroke-width="1.5"/>'
            )

    legend_html = "".join(
        f'<div class="legend-item"><div class="legend-swatch" style="background:{c}"></div>{esc(l)}</div>'
        for l, _, c in series
    )

    svg = (
        f'<svg viewBox="0 0 {width} {height}" width="100%" style="max-width:100%;height:auto;">'
        + "".join(grid_lines)
        + "".join(x_labels)
        + "".join(lines_svg)
        + '</svg>'
    )

    return (
        '<div class="chart-wrap">'
        '<h4>月度趨勢 (USD)</h4>'
        + svg
        + f'<div class="chart-legend">{legend_html}</div>'
        + '</div>'
    )


def render_section3(result) -> str:
    source_note = (
        f'<div class="chart-source">'
        f'<strong>ACC 1.0</strong>:直接取自 115 seller performance summary P24:Y29。 '
        f'<strong>ACC 2.0</strong>:以 115 位錄取 MCID 過濾最新 P0 (<code>{esc(result.p0_source.name)}</code>) '
        f'的 Sheet1,條件 calendar_year=2026 + launch_channel=DSR,依 calendar_month 加總 mtd_ord_gms。 '
        f'<strong>2026 NSR All</strong>:同條件不過濾 MCID。'
        f'</div>'
    )
    return (
        '<h2 class="section-title" id="section3">月度銷售額變化</h2>'
        '<section class="card" id="s3-monthly">'
        '<div class="card-header">'
        '<span>月度 GMS (Jan-Aug)</span>'
        f'<span class="total-pill">Source: {esc(result.p0_source.name)}</span>'
        '</div>'
        + render_monthly_table(result)
        + source_note
        + render_monthly_chart(result)
        + '</section>'
    )


# ============================================================
# 區塊 4 渲染 (YTD 達標率)
# ============================================================

def _ytd_num(v, decimals: int = 0) -> str:
    if v is None:
        return '<span class="no-data">—</span>'
    fmt = f"{{:,.{decimals}f}}"
    return fmt.format(v)


def _ytd_yoy(v) -> str:
    if v is None:
        return '<span class="no-data">—</span>'
    sign = "+" if v > 0 else ""
    cls = "mom-up" if v > 0 else ("mom-down" if v < 0 else "")
    if cls:
        return f'<span class="{cls}">{abs(v) * 100:.1f}%</span>'
    return f"{sign}{v * 100:.1f}%"


def _ytd_attain(v) -> str:
    if v is None:
        return '<td class="num no-data">—</td>'
    pct_txt = f"{v * 100:.1f}%"
    cls = "hit" if v >= 1 else "miss"
    return f'<td class="attain {cls}">{pct_txt}</td>'


def render_section4(ytd, n_launched: int = 112) -> str:
    r20 = ytd.row_20()
    rnsr = ytd.row_nsr()
    week = ytd.ytd.week

    # KPI 區
    n20_sellers = n_launched
    gms_per_seller = r20["g_actual"] / n20_sellers if n20_sellers else 0
    kpi = (
        '<div class="ytd-kpi-row">'
        f'<div class="ytd-kpi"><div class="kpi-label">2.0 YTD GMS</div>'
        f'<div class="kpi-value">${_ytd_num(r20["g_actual"])}</div>'
        f'<div class="kpi-sub">截至 W{week}</div></div>'
        f'<div class="ytd-kpi"><div class="kpi-label">2.0 YTD GMS / Seller</div>'
        f'<div class="kpi-value">${_ytd_num(gms_per_seller)}</div>'
        f'<div class="kpi-sub">{_ytd_num(r20["g_actual"])} / {n20_sellers} launched</div></div>'
        f'<div class="ytd-kpi"><div class="kpi-label">2.0 YTD GMS 達標率</div>'
        f'<div class="kpi-value" style="color:#e07b00">{_ytd_num(r20["g_attain"] * 100, 1) if r20["g_attain"] else "—"}%</div>'
        f'<div class="kpi-sub">{_ytd_num(r20["g_actual"])} / {_ytd_num(r20["g_goal"])}</div></div>'
        f'<div class="ytd-kpi"><div class="kpi-label">2.0 YTD YoY vs 2025</div>'
        f'<div class="kpi-value" style="color:{"#2f9e44" if r20["g_yoy_actual"] and r20["g_yoy_actual"] > 0 else "#e03131"}">'
        f'{_ytd_yoy(r20["g_yoy_actual"])}</div><div class="kpi-sub">vs ACC 1.0 2025</div></div>'
        '</div>'
    )

    # GMS/Seller 表
    ps_table = _render_ytd_table(
        "GMS / Seller (USD)",
        [("ACC 2.0", r20, "track-20"), ("2026 NSR All", rnsr, "track-nsr")],
        prefix="ps",
        decimals=0,
    )

    # GMS 表
    g_table = _render_ytd_table(
        "GMS (USD)",
        [("ACC 2.0", r20, "track-20"), ("2026 NSR All", rnsr, "track-nsr")],
        prefix="g",
        decimals=0,
    )

    source_note = (
        f'<div class="chart-source" style="margin:10px 22px 22px;">'
        f'<strong>Goal / YoY Goal</strong>:取自 115 seller performance summary P2:Z5。 '
        f'<strong>ACC 2.0 列 2025</strong>:取自 P24:Y29 的 ACC 1.0 Jan-Aug 合計,per seller 分母為 ACC 1.0 賽道 70 位賣家。 '
        f'<strong>NSR All 列 Goal</strong>:無對應資料,留白。 '
        f'<strong>2026 Actual</strong>:取自 WBR 最新週 P0 (<code>{esc(ytd.ytd.p0_source.name)}</code>) '
        f'raw 工作表,篩選 reporting_year=2026 + reporting_week_of_year={week} + launch_channel=DSR,'
        f'ACC 2.0 再以 115 位 MCID 過濾後加總 ytd_ord_gms。 '
        f'<strong>NSR All 2025</strong>:同份 P0 改 reporting_year=2025 + week max,不過濾 MCID。 '
        f'<strong>YoY Actual</strong>=(2026/2025)-1;<strong>達標率</strong>=2026 Actual / Goal。'
        f'</div>'
    )

    return (
        '<h2 class="section-title" id="section4">YTD 達標率</h2>'
        '<section class="card" id="s4-ytd">'
        '<div class="card-header">'
        f'<span>YTD 達標率 · 截至 2026 W{week}</span>'
        f'<span class="total-pill">Source: {esc(ytd.ytd.p0_source.name)}</span>'
        '</div>'
        + kpi
        + ps_table
        + g_table
        + source_note
        + '</section>'
    )


def _render_ytd_table(group_title: str, tracks: list, prefix: str, decimals: int = 0) -> str:
    """渲染一組表格 (GMS/Seller 或 GMS)。
    tracks: [(賽道名, dict, tr_cls), ...]
    prefix: "ps" 或 "g"
    """
    header = (
        '<tr>'
        '<th>賽道</th>'
        '<th>2025</th>'
        '<th>Goal</th>'
        '<th>2026 Actual</th>'
        '<th>YoY Actual</th>'
        '<th>達標率</th>'
        '</tr>'
    )
    rows_html = []
    for name, data, cls in tracks:
        v_2025 = data.get(f"{prefix}_2025")
        v_goal = data.get(f"{prefix}_goal")
        v_actual = data.get(f"{prefix}_actual")
        v_yoy_a = data.get(f"{prefix}_yoy_actual")
        v_attain = data.get(f"{prefix}_attain")
        rows_html.append(
            f'<tr class="{cls}">'
            f'<td>{esc(name)}</td>'
            f'<td>{_ytd_num(v_2025, decimals)}</td>'
            f'<td>{_ytd_num(v_goal, decimals)}</td>'
            f'<td>{_ytd_num(v_actual, decimals)}</td>'
            f'<td>{_ytd_yoy(v_yoy_a)}</td>'
            + _ytd_attain(v_attain)
            + '</tr>'
        )
    return (
        '<div class="ytd-group">'
        f'<div class="ytd-group-title">{esc(group_title)}</div>'
        '<table class="ytd-table">'
        f'<thead>{header}</thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        '</table>'
        '</div>'
    )


# ============================================================
# 區塊 5 渲染 (Adoption)
# ============================================================

def render_section5(adoption) -> str:
    """Seller Adoption 指標對比表。"""
    header = (
        '<tr>'
        '<th>指標</th>'
        '<th class="num">ACC 2.0<br><small>n=' + str(adoption.n_20) + '</small></th>'
        '<th class="num">Δ vs NSR<br><small>(bps)</small></th>'
        '<th class="num">NSR All<br><small>n=' + str(adoption.n_nsr) + '</small></th>'
        '<th class="num">Δ vs 1.0<br><small>(bps)</small></th>'
        '<th class="num">ACC 1.0<br><small>n=' + str(adoption.n_10) + '</small></th>'
        '</tr>'
    )
    rows_html = []
    for r20, rnsr, r10 in zip(adoption.rows_20, adoption.rows_nsr, adoption.rows_10):
        bps_vs_nsr = (r20.pct - rnsr.pct) * 10000
        bps_vs_10 = (r20.pct - r10.pct) * 10000
        bps_nsr_cls = "mom-up" if bps_vs_nsr > 0 else ("mom-down" if bps_vs_nsr < 0 else "")
        bps_10_cls = "mom-up" if bps_vs_10 > 0 else ("mom-down" if bps_vs_10 < 0 else "")
        bps_nsr_html = f'<span class="{bps_nsr_cls}">{abs(bps_vs_nsr):,.0f}bps</span>' if bps_nsr_cls else f'{bps_vs_nsr:,.0f}bps'
        bps_10_html = f'<span class="{bps_10_cls}">{abs(bps_vs_10):,.0f}bps</span>' if bps_10_cls else f'{bps_vs_10:,.0f}bps'
        rows_html.append(
            '<tr>'
            f'<td class="label">{esc(r20.label)}</td>'
            f'<td class="num">{r20.count} <small style="color:var(--c-muted)">({r20.pct*100:.0f}%)</small></td>'
            f'<td class="num">{bps_nsr_html}</td>'
            f'<td class="num">{rnsr.count} <small style="color:var(--c-muted)">({rnsr.pct*100:.0f}%)</small></td>'
            f'<td class="num">{bps_10_html}</td>'
            f'<td class="num">{r10.count} <small style="color:var(--c-muted)">({r10.pct*100:.0f}%)</small></td>'
            '</tr>'
        )

    # 視覺化:每個指標的三組佔比 bar
    bars_html = []
    for r20, rnsr, r10 in zip(adoption.rows_20, adoption.rows_nsr, adoption.rows_10):
        bars_html.append(
            '<div style="margin-bottom:14px;">'
            f'<div style="font-size:13px;font-weight:600;margin-bottom:4px;">{esc(r20.label)}</div>'
            + _adoption_bar("2.0", r20.pct, "#e07b00")
            + _adoption_bar("NSR", rnsr.pct, "#6c757d")
            + _adoption_bar("1.0", r10.pct, "#1f4e79")
            + '</div>'
        )

    source_note = (
        f'<div class="chart-source" style="margin:10px 22px 22px;">'
        f'<strong>資料來源</strong>:NSR Launch Tracker (<code>{esc(adoption.source.name)}</code>) '
        f'工作表「2026 Raw」。每位賣家(MCID)取各指標 max 值(0/1),再計數。 '
        f'<strong>ACC 2.0</strong>:以最終賣家名單 MCID 過濾(matched {adoption.n_20} 位)。 '
        f'<strong>NSR All</strong>:不過濾 MCID(全部 {adoption.n_nsr} 位)。 '
        f'<strong>ACC 1.0</strong>:以 ACC 1.0 GMS by seller 的 MCID 過濾(matched {adoption.n_10} 位)。'
        f'</div>'
    )

    return (
        '<h2 class="section-title" id="section5">Seller Adoption 指標</h2>'
        '<section class="card" id="s5-adoption">'
        '<div class="card-header">'
        f'<span>Adoption 對比 (ACC 2.0 vs NSR All vs ACC 1.0)</span>'
        f'<span class="total-pill">Source: {esc(adoption.source.name)}</span>'
        '</div>'
        '<table class="ytd-table">'
        f'<thead>{header}</thead>'
        f'<tbody>{"".join(rows_html)}</tbody>'
        '</table>'
        '<div style="padding:22px;">'
        + "".join(bars_html)
        + '</div>'
        + source_note
        + '</section>'
    )


def _adoption_bar(label: str, pct: float, color: str) -> str:
    w = max(0, min(pct * 100, 100))
    return (
        '<div style="display:flex;align-items:center;gap:8px;margin-bottom:3px;">'
        f'<span style="width:30px;font-size:11px;color:var(--c-muted);font-weight:600;">{label}</span>'
        f'<div style="flex:1;height:8px;background:#e5e7eb;border-radius:999px;overflow:hidden;">'
        f'<div style="width:{w:.1f}%;height:100%;background:{color};border-radius:999px;"></div>'
        '</div>'
        f'<span style="width:40px;font-size:11px;color:var(--c-muted);text-align:right;">{pct*100:.0f}%</span>'
        '</div>'
    )


# ============================================================
# 區塊 6 渲染 (Raw Data)
# ============================================================

def render_section6(raw_df: "pd.DataFrame") -> str:
    """Raw Data 表格:全寬、每欄篩選、底部加總。"""

    cols = list(raw_df.columns)
    ae_list = sorted(raw_df["AE"].dropna().unique().tolist())

    # 判斷哪些欄位是數字欄(用來做加總)
    num_cols = set()
    for c in cols:
        if raw_df[c].dtype in ("float64", "int64", "float32", "int32"):
            num_cols.add(c)

    # 表格 header (每欄加 select 篩選,MCID/公司名稱 用搜尋框,數字欄加排序)
    header_cells = []
    for i, c in enumerate(cols):
        if c == "MCID" or c == "公司名稱":
            # 搜尋輸入框
            placeholder = "搜尋 MCID..." if c == "MCID" else "搜尋公司名稱..."
            filter_html = (
                f'<input type="text" class="col-search" data-col="{i}" '
                f'placeholder="{placeholder}" '
                f'style="display:block;width:100%;margin-top:4px;font-size:11px;'
                f'padding:3px 6px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;">'
            )
        elif c not in num_cols and len(raw_df[c].dropna().astype(str).unique()) <= 50:
            unique_vals = sorted(raw_df[c].dropna().astype(str).unique().tolist())
            options = '<option value="">全部</option>' + "".join(
                f'<option value="{esc(v)}">{esc(v)}</option>' for v in unique_vals
            )
            filter_html = (
                f'<select class="col-filter" data-col="{i}" '
                f'style="display:block;width:100%;margin-top:4px;font-size:11px;'
                f'padding:2px 4px;border:1px solid #ddd;border-radius:4px;">'
                f'{options}</select>'
            )
        else:
            filter_html = ''

        # 數字欄加排序按鈕
        if c in num_cols:
            sort_html = (
                f'<span class="sort-btn" data-col="{i}" data-dir="none" '
                f'style="cursor:pointer;margin-left:4px;font-size:10px;color:#9ca3af;user-select:none;" '
                f'title="點擊排序">⇅</span>'
            )
        else:
            sort_html = ''

        header_cells.append(
            f'<th style="white-space:nowrap;min-width:80px;">'
            f'{esc(c)}{sort_html}{filter_html}</th>'
        )
    header = f'<tr>{"".join(header_cells)}</tr>'

    # 表格 body
    rows_html = []
    for _, row in raw_df.iterrows():
        cells = []
        cell_vals = []
        for c in cols:
            v = row[c]
            if pd.isna(v):
                cells.append('<td class="no-data">—</td>')
                cell_vals.append("")
            elif c in num_cols:
                cells.append(f'<td class="num">{v:,.0f}</td>')
                cell_vals.append(str(v))
            else:
                cells.append(f'<td>{esc(str(v))}</td>')
                cell_vals.append(str(v))
        # data attributes for filtering
        data_attrs = " ".join(f'data-c{i}="{esc(cell_vals[i])}"' for i in range(len(cols)))
        rows_html.append(f'<tr {data_attrs}>{"".join(cells)}</tr>')

    # 底部加總列
    total_cells = []
    for c in cols:
        if c in num_cols:
            s = raw_df[c].sum()
            total_cells.append(f'<td class="num" style="font-weight:700;background:#eef3fb;">{s:,.0f}</td>')
        elif c == cols[0]:
            total_cells.append(f'<td style="font-weight:700;background:#eef3fb;">合計 (<span id="raw-count">{len(raw_df)}</span> 位)</td>')
        else:
            total_cells.append('<td style="background:#eef3fb;"></td>')
    total_row = f'<tr class="total-row">{"".join(total_cells)}</tr>'

    # JS: 多欄篩選 (select + MCID 搜尋框)
    filter_js = """
<script>
(function() {
  const filters = document.querySelectorAll('#raw-table .col-filter');
  const searches = document.querySelectorAll('#raw-table .col-search');
  const tbody = document.querySelector('#raw-table tbody');
  const rows = tbody.querySelectorAll('tr:not(.total-row)');
  const countEl = document.getElementById('raw-count');
  const totalRow = tbody.querySelector('tr.total-row');

  function applyFilters() {
    const active = {};
    filters.forEach(sel => {
      if (sel.value) active[sel.dataset.col] = { type: 'exact', val: sel.value };
    });
    searches.forEach(inp => {
      if (inp.value.trim()) active[inp.dataset.col] = { type: 'search', val: inp.value.trim() };
    });
    let visible = 0;
    const sums = {};
    rows.forEach(tr => {
      let show = true;
      for (const [col, filter] of Object.entries(active)) {
        const cellVal = tr.getAttribute('data-c' + col) || '';
        if (filter.type === 'exact') {
          if (cellVal !== filter.val) { show = false; break; }
        } else if (filter.type === 'search') {
          if (!cellVal.includes(filter.val)) { show = false; break; }
        }
      }
      tr.style.display = show ? '' : 'none';
      if (show) {
        visible++;
        tr.querySelectorAll('td.num').forEach(td => {
          const idx = [...td.parentNode.children].indexOf(td);
          const v = parseFloat(td.textContent.replace(/,/g, ''));
          if (!isNaN(v)) sums[idx] = (sums[idx] || 0) + v;
        });
      }
    });
    countEl.textContent = visible;
    if (totalRow) {
      totalRow.querySelectorAll('td.num').forEach(td => {
        const idx = [...td.parentNode.children].indexOf(td);
        td.textContent = (sums[idx] || 0).toLocaleString('en-US', {maximumFractionDigits:0});
      });
    }
  }
  filters.forEach(sel => sel.addEventListener('change', applyFilters));
  searches.forEach(inp => inp.addEventListener('input', applyFilters));

  // 排序功能
  const sortBtns = document.querySelectorAll('#raw-table .sort-btn');
  const rowsArr = [...rows];
  rowsArr.forEach((tr, i) => tr.dataset.origIdx = i);

  sortBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const col = parseInt(btn.dataset.col);
      let dir = btn.dataset.dir;
      if (dir === 'none') dir = 'desc';
      else if (dir === 'desc') dir = 'asc';
      else dir = 'none';
      btn.dataset.dir = dir;
      btn.textContent = dir === 'desc' ? '↓' : dir === 'asc' ? '↑' : '⇅';
      sortBtns.forEach(b => { if (b !== btn) { b.dataset.dir = 'none'; b.textContent = '⇅'; } });

      if (dir === 'none') {
        rowsArr.sort((a, b) => parseInt(a.dataset.origIdx) - parseInt(b.dataset.origIdx));
      } else {
        rowsArr.sort((a, b) => {
          const va = parseFloat((a.getAttribute('data-c' + col) || '0').replace(/,/g, '')) || 0;
          const vb = parseFloat((b.getAttribute('data-c' + col) || '0').replace(/,/g, '')) || 0;
          return dir === 'desc' ? vb - va : va - vb;
        });
      }
      rowsArr.forEach(tr => tbody.insertBefore(tr, totalRow));
      applyFilters();
    });
  });
})();
</script>
"""

    return (
        '<h2 class="section-title" id="section6">Raw Data</h2>'
        '<section class="card" id="s6-raw" style="max-width:100vw;">'
        '<div class="card-header">'
        '<span>賣家明細</span>'
        '</div>'
        '<div style="overflow-x:auto;overflow-y:auto;max-height:80vh;">'
        '<table class="ytd-table" id="raw-table" style="font-size:12px;white-space:nowrap;">'
        f'<thead style="position:sticky;top:0;z-index:1;background:#f8fafc;">{header}</thead>'
        f'<tbody>{"".join(rows_html)}{total_row}</tbody>'
        '</table></div>'
        '</section>'
        + filter_js
    )


# ============================================================
# 區塊 7 渲染 (Weekly GMS)
# ============================================================

def render_section7(weekly, ytd_20: float = 0, ytd_10: float = 0, n_launched: int = 112) -> str:
    """Weekly Performance: GMS + GMS/Seller 折線圖與表格。"""
    weeks = weekly.weeks
    gms_20 = weekly.gms_20
    gms_10 = weekly.gms_10

    # 賣家數 (GMS/Seller 用已開賣數)
    n_20 = n_launched
    n_10 = 70

    # GMS/Seller (用每週 YTD GMS ÷ 賣家數)
    gps_20 = [v / n_20 if v > 0 else 0 for v in weekly.ytd_20]
    gps_10 = [v / n_10 if v > 0 else 0 for v in weekly.ytd_10]

    # SVG 折線圖
    all_vals = [v for v in gms_20 + gms_10 if v > 0]
    y_max = max(all_vals) * 1.15 if all_vals else 1

    width, height = 900, 280
    pad_l, pad_r, pad_t, pad_b = 60, 20, 20, 40
    plot_w = width - pad_l - pad_r
    plot_h = height - pad_t - pad_b
    n = len(weeks)
    x_step = plot_w / max(n - 1, 1)

    def pt(i, v):
        x = pad_l + i * x_step
        y = pad_t + plot_h * (1 - v / y_max) if y_max else pad_t
        return x, y

    # Grid
    grid = []
    for i in range(5):
        ratio = i / 4
        y = pad_t + plot_h * (1 - ratio)
        label = y_max * ratio
        label_txt = f"${label/1000:.0f}K"
        grid.append(f'<line x1="{pad_l}" y1="{y:.0f}" x2="{width-pad_r}" y2="{y:.0f}" stroke="#e5e7eb" stroke-width="1" stroke-dasharray="2,3"/>')
        grid.append(f'<text x="{pad_l-6}" y="{y+4:.0f}" text-anchor="end" font-size="10" fill="#9ca3af">{label_txt}</text>')

    # X labels
    x_labels = []
    for i, w in enumerate(weeks):
        x = pad_l + i * x_step
        x_labels.append(f'<text x="{x:.0f}" y="{height-pad_b+18}" text-anchor="middle" font-size="10" fill="#6b7280">W{w}</text>')

    # Lines
    series = [("ACC 2.0 (2026)", gms_20, "#e07b00"), ("ACC 1.0 (2025)", gms_10, "#1f4e79")]
    lines_svg = []
    for label, vals, color in series:
        pts = [(pt(i, v)) for i, v in enumerate(vals) if v > 0]
        if len(pts) >= 2:
            d = "M " + " L ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
            lines_svg.append(f'<path d="{d}" fill="none" stroke="{color}" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>')
        for x, y in pts:
            lines_svg.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="{color}" stroke="#fff" stroke-width="1.5"/>')

    svg = (
        f'<svg viewBox="0 0 {width} {height}" width="100%" style="max-width:100%;height:auto;">'
        + "".join(grid) + "".join(x_labels) + "".join(lines_svg)
        + '</svg>'
    )

    legend = (
        '<div class="chart-legend">'
        '<div class="legend-item"><div class="legend-swatch" style="background:#e07b00"></div>ACC 2.0 (2026)</div>'
        '<div class="legend-item"><div class="legend-swatch" style="background:#1f4e79"></div>ACC 1.0 (2025)</div>'
        '</div>'
    )

    # 表格
    # 合計欄用最新 P0 的 ytd_ord_gms (和 YTD 達標率一致)
    # 從 weekly_gms cache 裡取不到 ytd,需要從外部傳入
    # 暫時用 sum,之後在 build() 傳入 ytd 值
    header = '<tr><th>Week</th>' + "".join(f'<th class="num">W{w}</th>' for w in weeks) + '<th class="num" style="background:#eef3fb;border-left:2px solid var(--c-primary-soft);">YTD 合計</th></tr>'
    
    # ACC 2.0 row + YTD 合計
    row_20 = '<tr class="row-20"><td style="font-weight:600;color:#e07b00;">ACC 2.0</td>' + "".join(
        f'<td class="num">{v:,.0f}</td>' if v > 0 else '<td class="no-data">—</td>' for v in gms_20
    ) + f'<td class="num" style="font-weight:700;background:#eef3fb;border-left:2px solid var(--c-primary-soft);">{ytd_20:,.0f}</td></tr>'
    
    # ACC 1.0 row + YTD 合計
    row_10 = '<tr class="row-10"><td style="font-weight:600;color:#1f4e79;">ACC 1.0</td>' + "".join(
        f'<td class="num">{v:,.0f}</td>' if v > 0 else '<td class="no-data">—</td>' for v in gms_10
    ) + f'<td class="num" style="font-weight:700;background:#eef3fb;border-left:2px solid var(--c-primary-soft);">{ytd_10:,.0f}</td></tr>'
    
    # YoY row (合計放 YTD 的 YoY%)
    yoy_vals = []
    for g20, g10 in zip(gms_20, gms_10):
        if g20 > 0 and g10 > 0:
            yoy_vals.append((g20 / g10 - 1) * 100)
        else:
            yoy_vals.append(None)
    ytd_yoy = (ytd_20 / ytd_10 - 1) * 100 if ytd_10 > 0 and ytd_20 > 0 else None
    ytd_yoy_html = ''
    if ytd_yoy is not None:
        cls = "mom-up" if ytd_yoy > 0 else "mom-down"
        ytd_yoy_html = f'<span class="{cls}">{abs(ytd_yoy):.0f}%</span>'
    else:
        ytd_yoy_html = '—'
    row_yoy = '<tr class="row-mom"><td style="font-weight:600;color:var(--c-muted);">YoY</td>' + "".join(
        f'<td class="num"><span class="{"mom-up" if v > 0 else "mom-down"}">{abs(v):.0f}%</span></td>' if v is not None
        else '<td class="no-data">—</td>' for v in yoy_vals
    ) + f'<td class="num" style="font-weight:700;background:#eef3fb;border-left:2px solid var(--c-primary-soft);">{ytd_yoy_html}</td></tr>'
    
    # WoW row (2.0 的每週變化,合計留空)
    wow_vals = []
    prev = None
    for v in gms_20:
        if v > 0 and prev is not None and prev > 0:
            wow_vals.append((v / prev - 1) * 100)
        else:
            wow_vals.append(None)
        if v > 0:
            prev = v
    row_wow = '<tr class="row-vs"><td style="font-weight:600;color:var(--c-muted);">WoW (2.0)</td>' + "".join(
        f'<td class="num"><span class="{"mom-up" if v > 0 else "mom-down"}">{abs(v):.0f}%</span></td>' if v is not None
        else '<td class="no-data">—</td>' for v in wow_vals
    ) + '<td style="background:#eef3fb;border-left:2px solid var(--c-primary-soft);"></td></tr>'

    table = (
        '<div style="overflow-x:auto;margin-top:16px;">'
        '<h4 style="padding:0 16px;font-size:13px;color:var(--c-primary);margin-bottom:4px;">Weekly GMS (USD)</h4>'
        '<table class="monthly-table">'
        f'<thead>{header}</thead>'
        f'<tbody>{row_20}{row_10}{row_yoy}{row_wow}</tbody>'
        '</table></div>'
    )

    # === GMS / Seller 折線圖 ===
    all_ps_vals = [v for v in gps_20 + gps_10 if v > 0]
    y_max_ps = max(all_ps_vals) * 1.15 if all_ps_vals else 1

    lines_svg_ps = []
    series_ps = [("ACC 2.0 (2026)", gps_20, "#e07b00"), ("ACC 1.0 (2025)", gps_10, "#1f4e79")]
    for label_s, vals, color in series_ps:
        pts = []
        for i, v in enumerate(vals):
            if v > 0:
                x = pad_l + i * x_step
                y = pad_t + plot_h * (1 - v / y_max_ps)
                pts.append((x, y))
        if len(pts) >= 2:
            d = "M " + " L ".join(f"{x:.1f},{y:.1f}" for x, y in pts)
            lines_svg_ps.append(f'<path d="{d}" fill="none" stroke="{color}" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/>')
        for x, y in pts:
            lines_svg_ps.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="{color}" stroke="#fff" stroke-width="1.5"/>')

    grid_ps = []
    for i in range(5):
        ratio = i / 4
        y = pad_t + plot_h * (1 - ratio)
        label_v = y_max_ps * ratio
        label_txt = f"${label_v:,.0f}"
        grid_ps.append(f'<line x1="{pad_l}" y1="{y:.0f}" x2="{width-pad_r}" y2="{y:.0f}" stroke="#e5e7eb" stroke-width="1" stroke-dasharray="2,3"/>')
        grid_ps.append(f'<text x="{pad_l-6}" y="{y+4:.0f}" text-anchor="end" font-size="10" fill="#9ca3af">{label_txt}</text>')

    svg_ps = (
        f'<svg viewBox="0 0 {width} {height}" width="100%" style="max-width:100%;height:auto;">'
        + "".join(grid_ps) + "".join(x_labels) + "".join(lines_svg_ps)
        + '</svg>'
    )

    # === GMS / Seller 表格 ===
    header_ps = '<tr><th>Week</th>' + "".join(f'<th class="num">W{w}</th>' for w in weeks) + '<th class="num" style="background:#eef3fb;border-left:2px solid var(--c-primary-soft);">YTD Avg</th></tr>'

    ytd_ps_20 = ytd_20 / n_20 if n_20 else 0
    ytd_ps_10 = ytd_10 / n_10 if n_10 else 0

    row_ps_20 = '<tr class="row-20"><td style="font-weight:600;color:#e07b00;">ACC 2.0</td>' + "".join(
        f'<td class="num">{v:,.0f}</td>' if v > 0 else '<td class="no-data">—</td>' for v in gps_20
    ) + f'<td class="num" style="font-weight:700;background:#eef3fb;border-left:2px solid var(--c-primary-soft);">{ytd_ps_20:,.0f}</td></tr>'

    row_ps_10 = '<tr class="row-10"><td style="font-weight:600;color:#1f4e79;">ACC 1.0</td>' + "".join(
        f'<td class="num">{v:,.0f}</td>' if v > 0 else '<td class="no-data">—</td>' for v in gps_10
    ) + f'<td class="num" style="font-weight:700;background:#eef3fb;border-left:2px solid var(--c-primary-soft);">{ytd_ps_10:,.0f}</td></tr>'

    # YoY per seller
    yoy_ps_vals = []
    for g20, g10 in zip(gps_20, gps_10):
        if g20 > 0 and g10 > 0:
            yoy_ps_vals.append((g20 / g10 - 1) * 100)
        else:
            yoy_ps_vals.append(None)
    ytd_ps_yoy = (ytd_ps_20 / ytd_ps_10 - 1) * 100 if ytd_ps_10 > 0 and ytd_ps_20 > 0 else None
    ytd_ps_yoy_html = ''
    if ytd_ps_yoy is not None:
        cls = "mom-up" if ytd_ps_yoy > 0 else "mom-down"
        ytd_ps_yoy_html = f'<span class="{cls}">{abs(ytd_ps_yoy):.0f}%</span>'
    else:
        ytd_ps_yoy_html = '—'
    row_ps_yoy = '<tr class="row-mom"><td style="font-weight:600;color:var(--c-muted);">YoY</td>' + "".join(
        f'<td class="num"><span class="{"mom-up" if v > 0 else "mom-down"}">{abs(v):.0f}%</span></td>' if v is not None
        else '<td class="no-data">—</td>' for v in yoy_ps_vals
    ) + f'<td class="num" style="font-weight:700;background:#eef3fb;border-left:2px solid var(--c-primary-soft);">{ytd_ps_yoy_html}</td></tr>'

    table_ps = (
        '<div class="chart-wrap" style="margin-top:24px;">'
        '<h4>Weekly GMS / Seller (USD)</h4>'
        + svg_ps + legend
        + '</div>'
        '<div style="overflow-x:auto;margin-top:16px;">'
        '<table class="monthly-table">'
        f'<thead>{header_ps}</thead>'
        f'<tbody>{row_ps_20}{row_ps_10}{row_ps_yoy}</tbody>'
        '</table></div>'
    )

    source_note = (
        '<div class="chart-source" style="margin:10px 22px 22px;">'
        '<strong>資料來源</strong>:TWGS - 2026 WBR 各週資料夾的 P0 開頭檔案,工作表 raw。 '
        '<strong>篩選條件</strong>:launch_channel=DSR,reporting_week_of_year=該週週數。 '
        '<strong>ACC 2.0</strong>:reporting_year=2026 + MCID ∈ 最終賣家名單 (115),加總 wtd_ord_gms。 '
        '<strong>ACC 1.0</strong>:reporting_year=2025 + MCID ∈ ACC 1.0 GMS by seller,加總 wtd_ord_gms。'
        '</div>'
    )

    return (
        '<h2 class="section-title" id="section7">Weekly Performance</h2>'
        '<section class="card" id="s7-weekly">'
        '<div class="card-header"><span>ACC 2.0 vs ACC 1.0 · 每週 Performance</span></div>'
        '<div class="chart-wrap">'
        '<h4>Weekly WTD GMS (USD)</h4>'
        + svg + legend
        + '</div>'
        + table
        + table_ps
        + source_note
        + '</section>'
    )


def render_notes() -> str:
    return (
        '<div class="note-box">'
        '<h4>資料說明</h4>'
        '<ul>'
        '<li>1.0 站點欄以「區域」為單位 (如「北美 (美國/加拿大/墨西哥)」),括號內明細不拆。括號外多區域以逗號拆分,每區域各計 +1,分母為 1.0 賣家數。</li>'
        '<li>US vs Expansion 為互斥分類:賣家站點含 US 即計 US,否則計 Expansion,加總 = 100%。</li>'
        '<li>1.0 無對應欄位項目 (跳過對比):公司所在區域、Category、註冊商標。</li>'
        '<li>1.0 「預計投入亞馬遜 FBA 起始庫存」與 2.0 選項邏輯完全不同 (1.0 以產品價值區分、2.0 以美金區間區分),已從對比移除。</li>'
        '<li>「品牌創立年限」:1.0 的「六到十年」+「十年以上」合併為「六年以上 (1.0 合併)」,可對比 2.0 的「六到八年 + 八年以上」合計。</li>'
        '<li>「預計投入廣告預算」:1.0 當年問卷無 $1,200-1,500/月 級距,2.0 顯示 1.0 該類 0 人。</li>'
        '<li>「是否有國內電商銷售經驗」2.0 來源樞紐表將複選組合拆為多列,類別加總 = 115 但語意可能重疊,照原樣呈現。</li>'
        '<li>區塊 3 月度 GMS 使用最新 P0 檔重算,舊月份數字可能與固定報表 (P24:Y29) 有差異,屬正常刷新。</li>'
        '</ul>'
        '</div>'
    )


# ============================================================
# TOC 與頁面組裝
# ============================================================

def render_tabs_nav() -> str:
    return (
        '<nav class="toc">'
        '<button class="tab active" data-target="panel1">ACC 2.0 Profile</button>'
        '<button class="tab" data-target="panel2">1.0 vs 2.0 Profile 比較</button>'
        '<button class="tab" data-target="panel3">月度 GMS</button>'
        '<button class="tab" data-target="panel4">YTD 達標率</button>'
        '<button class="tab" data-target="panel7">Weekly Performance</button>'
        '<button class="tab" data-target="panel5">Adoption</button>'
        '<button class="tab" data-target="panel6">Raw Data</button>'
        '</nav>'
    )


TABS_JS = """
<script>
(function() {
  const tabs = document.querySelectorAll('nav.toc button.tab');
  const panels = document.querySelectorAll('.tab-panel');
  tabs.forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.target;
      tabs.forEach(t => t.classList.toggle('active', t === btn));
      panels.forEach(p => p.classList.toggle('active', p.id === target));
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  });
})();
</script>
"""


def render_hero(n20: int, n10: int, us_exp_rows, n_target: int = 115, n_launched: int = 0,
               ytd_gms: float = 0, ytd_yoy: float = 0, gms_per_seller: float = 0,
               snapshot: dict | None = None) -> str:
    us_pct = next((p for l, c, p in us_exp_rows if l == "US"), 0)
    exp_pct = next((p for l, c, p in us_exp_rows if l.startswith("Expansion")), 0)
    updated = dt.datetime.now().strftime("%Y-%m-%d")
    attain_pct = n20 / n_target * 100 if n_target else 0
    yoy_pct = (n20 / n10 - 1) * 100 if n10 else 0
    launched_pct = n_launched / n20 * 100 if n20 else 0

    # Snapshot summary text
    snapshot_html = ""
    if snapshot:
        week = snapshot["week"]
        p0_name = snapshot["p0_name"]
        s_ytd_gms = snapshot["ytd_gms"]
        s_ytd_gms_10 = snapshot["ytd_gms_10"]
        s_gms_goal = snapshot["gms_goal"]
        s_ps_goal = snapshot["ps_goal"]
        s_per_seller = snapshot["per_seller"]
        s_per_seller_10 = snapshot["per_seller_10"]
        s_launched = snapshot["launched"]
        s_launched_10 = snapshot["launched_10"]
        s_pl_count = snapshot["pl_count"]
        s_pl_rate = snapshot["pl_rate"]
        s_pl_bps_nsr = snapshot["pl_bps_nsr"]
        s_pl_bps_10 = snapshot["pl_bps_10"]
        # Goal curve = week / 52
        goal_curve = week / 52
        gms_curve_target = s_gms_goal * goal_curve
        ps_curve_target = s_ps_goal * goal_curve
        gms_vs_curve = (s_ytd_gms / gms_curve_target - 1) * 100 if gms_curve_target else 0
        ps_vs_curve = (s_per_seller / ps_curve_target - 1) * 100 if ps_curve_target else 0
        launched_diff = s_launched - n_target
        # vs ACC 1.0
        gms_vs_10 = (s_ytd_gms / s_ytd_gms_10 - 1) * 100 if s_ytd_gms_10 else 0
        launched_vs_10 = s_launched - s_launched_10
        ps_vs_10 = (s_per_seller / s_per_seller_10 - 1) * 100 if s_per_seller_10 else 0
        snap_date = dt.datetime.now().strftime("%#m/%#d")

        # 未開賣賣家列表
        unlaunched = snapshot.get("unlaunched", [])
        unlaunched_html = ""
        if unlaunched:
            unlaunched_html = (
                '<div style="font-size:11px;color:rgba(255,255,255,0.7);margin-top:10px;'
                'border-top:1px solid rgba(255,255,255,0.2);padding-top:8px;">'
                f'<strong>未開賣 ({len(unlaunched)} 位):</strong>'
            )
            for s in unlaunched:
                unlaunched_html += f'<br>MCID: {s["mcid"]} · {s["name"]} · AE: {s["ae"]}'
            unlaunched_html += '</div>'

        snapshot_html = (
            '<div class="snapshot-box">'
            f'<div class="snapshot-title">ACC 2.0\'s performance (Snapshot on {snap_date}, with {p0_name} W{week} data)</div>'
            '<ul class="snapshot-list">'
            f'<li><strong>YTD GMS:</strong> ${s_ytd_gms:,.0f} ({gms_vs_curve:+.1f}% vs goal curve, {gms_vs_10:+.1f}% vs ACC 1.0)</li>'
            f'<li><strong>Seller launched:</strong> {s_launched} ({launched_diff:+d} vs goal {n_target}, {launched_vs_10:+d} vs ACC 1.0)</li>'
            f'<li><strong>Per Seller GMS:</strong> ${s_per_seller:,.0f} ({ps_vs_curve:+.1f}% vs goal curve, {ps_vs_10:+.1f}% vs ACC 1.0)</li>'
            f'<li><strong>Perfect Launch:</strong> {s_pl_count} Sellers with {s_pl_rate:.0f}% rate '
            f'({s_pl_bps_nsr:+,.0f}bps vs NSR, {s_pl_bps_10:+,.0f}bps vs ACC 1.0)</li>'
            '</ul>'
            + unlaunched_html
            + '</div>'
        )

    return (
        '<header class="hero">'
        '<h1>ACC 2.0 錄取賣家分析</h1>'
        f'<div class="meta">資料來源:ACC 2.0 最終錄取結果 (20260105).xlsx · 1.0 對照:ACC 1.0分析.xlsx · 產出日期:{updated}</div>'
        '<div class="hero-content">'
        '<div class="kpi-row">'
        f'<div class="kpi"><div class="label">ACC 2.0 錄取賣家</div>'
        f'<div class="value">{n20} <span style="font-size:14px;color:rgba(255,255,255,0.7)">/ {n_target}</span></div>'
        f'<div class="sub">達成率 {attain_pct:.0f}%</div></div>'
        f'<div class="kpi"><div class="label">錄取賣家數 YoY (vs ACC 1.0)</div>'
        f'<div class="value">{yoy_pct:+.0f}%</div>'
        f'<div class="sub">ACC 1.0: {n10} 位 → ACC 2.0: {n20} 位</div></div>'
        f'<div class="kpi"><div class="label">Launched Seller</div>'
        f'<div class="value">{n_launched} <span style="font-size:14px;color:rgba(255,255,255,0.7)">/ {n20}</span></div>'
        f'<div class="sub">US {pct(us_pct)} · Expansion {pct(exp_pct)}</div></div>'
        f'<div class="kpi"><div class="label">YTD GMS</div><div class="value">${ytd_gms:,.0f}</div>'
        f'<div class="sub">ACC 2.0 · YoY {ytd_yoy:+.0f}% vs ACC 1.0 2025</div></div>'
        f'<div class="kpi"><div class="label">YTD GMS / Seller</div><div class="value">${gms_per_seller:,.0f}</div>'
        f'<div class="sub">${ytd_gms:,.0f} / {n_launched} launched</div></div>'
        '</div>'
        + snapshot_html
        + '</div>'
        '</header>'
    )


def build() -> None:
    print("[1/7] 讀取 2.0 分析...")
    blocks_20 = get_20_blocks()
    ae_rows, us_exp = get_recruit_country()

    print("[2/7] 讀取 1.0 最終錄取名單...")
    df_10 = load_10_roster()
    df_10 = df_10[df_10["公司名稱"].notna()].copy()
    n10 = len(df_10)
    print(f"  1.0 賣家數: {n10}")

    print("[3/7] 計算月度 GMS...")
    monthly_result = build_monthly_result()

    print("[4/8] 計算 YTD 達標率...")
    ytd_result = build_ytd_analysis()

    print("[5/8] 計算 Adoption...")
    adoption_result = build_adoption_analysis()

    print("[6/8] 組裝 Raw Data...")
    raw_df = build_raw_data()

    print("[7/8] 計算 Weekly GMS...")
    weekly_result = build_weekly_gms()

    print("[8/8] 產出 HTML...")
    compare_keys = [item for item, col, _ in COMPARISON_MAP if col in df_10.columns]

    # Build snapshot data
    week = ytd_result.ytd.week
    pl_row = next((r for r in adoption_result.rows_20 if r.label == "PL (Perfect Launch)"), None)
    pl_nsr_row = next((r for r in adoption_result.rows_nsr if r.label == "PL (Perfect Launch)"), None)
    pl_10_row = next((r for r in adoption_result.rows_10 if r.label == "PL (Perfect Launch)"), None)

    # 找未開賣賣家 (在名單但不在 Tracker)
    from adoption_analysis import load_20_mcids as _load_20_mcids, find_latest_nsr_tracker, WBR_BASE
    _all_mcids = _load_20_mcids()
    _tracker_path = find_latest_nsr_tracker(WBR_BASE)
    _tracker_df = pd.read_excel(_tracker_path, sheet_name="2026 Raw", engine="openpyxl",
                                usecols=["merchant_customer_id"])
    _tracker_df = _tracker_df.dropna(subset=["merchant_customer_id"])
    _tracker_df["mcid"] = _tracker_df["merchant_customer_id"].apply(
        lambda x: str(int(x)) if isinstance(x, (int, float)) else str(x).strip()
    )
    _tracker_mcids = set(_tracker_df["mcid"])
    _not_launched_mcids = _all_mcids - _tracker_mcids
    # 從 roster 取公司名稱和 AE
    _roster = pd.read_excel(FILE_20, sheet_name="最終賣家名單 (115)", engine="openpyxl")
    _roster["MCID_str"] = _roster["MCID"].dropna().astype(str).str.strip()
    _name_col = next((c for c in _roster.columns if "公司名稱" in str(c)), None)
    _unlaunched_list = []
    for _mcid in sorted(_not_launched_mcids):
        _row = _roster[_roster["MCID_str"] == _mcid]
        if not _row.empty:
            _unlaunched_list.append({
                "mcid": _mcid,
                "name": str(_row.iloc[0][_name_col]) if _name_col else "N/A",
                "ae": str(_row.iloc[0]["AE"]) if "AE" in _roster.columns else "N/A",
            })

    snapshot_data = {
        "week": week,
        "p0_name": ytd_result.ytd.p0_source.stem,
        "ytd_gms": ytd_result.ytd.gms_20_ytd,
        "ytd_gms_10": ytd_result.ytd.gms_10_ytd,
        "gms_goal": ytd_result.plan.gms_goal,
        "ps_goal": ytd_result.plan.gms_per_seller_goal,
        "per_seller": ytd_result.ytd.gms_20_ytd / adoption_result.n_20 if adoption_result.n_20 else 0,
        "per_seller_10": ytd_result.ytd.gms_10_ytd / 70 if ytd_result.ytd.gms_10_ytd else 0,
        "launched": adoption_result.n_20,
        "launched_10": 70,
        "unlaunched": _unlaunched_list,
        "pl_count": pl_row.count if pl_row else 0,
        "pl_rate": pl_row.pct * 100 if pl_row else 0,
        "pl_bps_nsr": (pl_row.pct - pl_nsr_row.pct) * 10000 if pl_row and pl_nsr_row else 0,
        "pl_bps_10": (pl_row.pct - pl_10_row.pct) * 10000 if pl_row and pl_10_row else 0,
    }

    body = (
        render_hero(115, n10, us_exp, n_target=115, n_launched=adoption_result.n_20,
                    ytd_gms=ytd_result.ytd.gms_20_ytd,
                    ytd_yoy=(ytd_result.ytd.gms_20_ytd / ytd_result.ytd.gms_10_ytd - 1) * 100 if ytd_result.ytd.gms_10_ytd > 0 else 0,
                    gms_per_seller=ytd_result.ytd.gms_20_ytd / adoption_result.n_20 if adoption_result.n_20 else 0,
                    snapshot=snapshot_data)
        + render_tabs_nav()
        + '<main>'
        + '<div class="tab-panel active" id="panel1">'
        + render_section1(blocks_20, ae_rows, us_exp)
        + '</div>'
        + '<div class="tab-panel" id="panel2">'
        + render_section2(blocks_20, df_10, n10)
        + '</div>'
        + '<div class="tab-panel" id="panel3">'
        + render_section3(monthly_result)
        + '</div>'
        + '<div class="tab-panel" id="panel4">'
        + render_section4(ytd_result, n_launched=adoption_result.n_20)
        + '</div>'
        + '<div class="tab-panel" id="panel5">'
        + render_section5(adoption_result)
        + '</div>'
        + '<div class="tab-panel" id="panel6">'
        + render_section6(raw_df)
        + '</div>'
        + '<div class="tab-panel" id="panel7">'
        + render_section7(weekly_result, ytd_20=ytd_result.ytd.gms_20_ytd, ytd_10=ytd_result.ytd.gms_10_ytd, n_launched=adoption_result.n_20)
        + '</div>'
        + '</main>'
        + '<footer>ACC 2.0 分析 · '
        + dt.datetime.now().strftime("%Y-%m-%d %H:%M")
        + '</footer>'
        + TABS_JS
    )

    html_doc = (
        '<!DOCTYPE html>'
        '<html lang="zh-Hant">'
        '<head>'
        '<meta charset="UTF-8">'
        '<meta name="viewport" content="width=device-width,initial-scale=1">'
        '<title>ACC 2.0 錄取賣家分析</title>'
        '<link rel="preconnect" href="https://fonts.googleapis.com">'
        '<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>'
        '<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap" rel="stylesheet">'
        f'<style>{CSS}</style>'
        '</head>'
        '<body>'
        '<div id="login-overlay" style="position:fixed;top:0;left:0;right:0;bottom:0;z-index:9999;'
        'background:linear-gradient(135deg,#1f4e79 0%,#2e75b6 100%);display:flex;align-items:center;'
        'justify-content:center;font-family:inherit;">'
        '<div style="background:white;border-radius:16px;padding:40px 48px;box-shadow:0 20px 60px rgba(0,0,0,0.3);'
        'text-align:center;max-width:360px;width:90%;">'
        '<h2 style="margin:0 0 8px;color:#1f4e79;font-size:22px;">ACC 2.0 Dashboard</h2>'
        '<p style="color:#6b7280;font-size:13px;margin:0 0 24px;">請輸入密碼以存取</p>'
        '<input id="pwd-input" type="password" placeholder="Password" '
        'style="width:100%;padding:12px 16px;border:2px solid #e5e7eb;border-radius:8px;font-size:15px;'
        'outline:none;transition:border-color 0.2s;box-sizing:border-box;" '
        'onkeydown="if(event.key===\'Enter\')document.getElementById(\'pwd-btn\').click()">'
        '<div id="pwd-error" style="color:#e03131;font-size:12px;margin-top:8px;min-height:18px;"></div>'
        '<button id="pwd-btn" style="margin-top:12px;width:100%;padding:12px;background:#1f4e79;color:white;'
        'border:none;border-radius:8px;font-size:15px;font-weight:600;cursor:pointer;transition:background 0.2s;" '
        'onmouseover="this.style.background=\'#2e75b6\'" onmouseout="this.style.background=\'#1f4e79\'">進入</button>'
        '</div></div>'
        '<script>'
        'document.getElementById("pwd-btn").addEventListener("click",function(){'
        'var p=document.getElementById("pwd-input").value;'
        'if(p==="ACC2.0"){'
        'document.getElementById("login-overlay").style.display="none";'
        '}else{'
        'document.getElementById("pwd-error").textContent="密碼錯誤，請重試";'
        'document.getElementById("pwd-input").style.borderColor="#e03131";'
        '}'
        '});'
        '</script>'
        f'{body}'
        '</body></html>'
    )

    OUT_FILE.write_text(html_doc, encoding="utf-8")
    print(f"完成: {OUT_FILE}")


if __name__ == "__main__":
    build()
