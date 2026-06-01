"""找出 ACC 2.0 問卷原始資料在哪個工作表 (含公司類型/區域/產品類型等欄位)。"""
import pandas as pd
from pathlib import Path

BASE = Path(r"C:\Users\hsinyih\OneDrive - amazon.com\ACC 2.0")
FPATH = BASE / "ACC 2.0 最終錄取結果 (20260105).xlsx"

KEYWORDS = ["公司類型", "公司所在區域", "產品類型", "經營型態", "亞馬遜賣家帳號",
            "品牌創立", "註冊商標", "廣告預算", "FBA 起始庫存", "專職人員", "專人營運",
            "國內電商", "海外銷售", "預計開始銷售", "Category"]

xls = pd.ExcelFile(FPATH, engine="openpyxl")
print(f"共 {len(xls.sheet_names)} 個工作表\n")

for sheet in xls.sheet_names:
    try:
        # 讀前 3 列找 header
        raw = pd.read_excel(FPATH, sheet_name=sheet, header=None, nrows=3, engine="openpyxl")
        df_head = pd.read_excel(FPATH, sheet_name=sheet, engine="openpyxl", nrows=1)
    except Exception as e:
        print(f"[{sheet}] 讀取失敗: {e}")
        continue

    cols = [str(c) for c in df_head.columns]
    # 同時檢查原始前幾列的字串,因為真正的 header 可能不在第一列
    all_text = " ".join(cols) + " " + " ".join(str(v) for v in raw.values.flatten() if pd.notna(v))
    hits = [k for k in KEYWORDS if k in all_text]
    shape_full = pd.read_excel(FPATH, sheet_name=sheet, engine="openpyxl").shape
    print(f"[{sheet}] 形狀 {shape_full}, 命中關鍵字: {hits}")
