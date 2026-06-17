"""
Excel 読み書きモジュール

卸問屋のJANコードExcelを読み込み、
Keepa結果をSheet3形式で出力する。
"""

from __future__ import annotations

import io
from datetime import date
import openpyxl
import pandas as pd


def read_wholesaler_excel(file_buffer) -> pd.DataFrame:
    """
    卸問屋Excelから JAN コード・品名・卸価格などを読み取る。
    シート「商品データ」(先頭シート) の D列=JAN, E列=品名, F列=品番, H列=定価, I列=卸, J列=箱入数。
    """
    wb = openpyxl.load_workbook(file_buffer, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True)):
        jan = str(row[3]).strip() if row[3] else ""
        if not jan or not jan.isdigit() or len(jan) < 8:
            continue
        rows.append({
            "jan_code": jan,
            "product_name_jp": str(row[4]).strip() if row[4] else "",
            "part_number": str(row[5]).strip() if row[5] else "",
            "retail_price": _to_float(row[7]),
            "wholesale_price": _to_float(row[8]),
            "box_qty": _to_int(row[9]),
        })

    wb.close()
    return pd.DataFrame(rows)


# Sheet3 のヘッダー定義（Row2）
_SHEET3_HEADERS = [
    "重複確認",          # A (1)
    "日付",              # B (2)
    "購入先問屋",        # C (3)
    "商品名",            # D (4)
    "タイトル",          # E (5)
    "SKU",               # F (6)
    "ASIN",              # G (7)
    "備考",              # H (8)
    "型番",              # I (9)
    "購入在庫",          # J (10)
    "売価最新更新日",    # K (11)
    "売価90日変化率",    # L (12)
    "BBセラー",          # M (13)
    "出品許可",          # N (14)
    "30キーパ",          # O (15)
    "セラー数（FBA）",   # P (16)
    "セラー数（無在庫）",# Q (17)
    "販売数（自社1M）",  # R (18)
    "初回仕入れ",        # S (19)
    "売価USD",           # T (20)
    "仕入＋送料（FBA）", # U (21)
    "Amazon手数料",      # V (22)
    "仕入値",            # W (23)
    "Weight（FBA）",     # X (24)
    "損益",              # Y (25)
    "利益額",            # Z (26)
    "利益率",            # AA (27)
]


def write_output_excel(results: list[dict], wholesaler_name: str = "") -> io.BytesIO:
    """
    Keepa 取得結果 + 利益計算済みデータを Sheet3 形式の Excel として出力する。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "リサーチ結果"

    # ヘッダー行
    ws.append(_SHEET3_HEADERS)

    today = date.today().isoformat()

    for r in results:
        ws.append([
            "",                                      # A: 重複確認
            today,                                   # B: 日付
            wholesaler_name,                         # C: 購入先問屋
            r.get("product_name_jp", ""),             # D: 商品名
            r.get("title", ""),                       # E: タイトル
            "",                                      # F: SKU
            r.get("asin", ""),                        # G: ASIN
            "" if r.get("found") else "US未出品",     # H: 備考
            r.get("part_number", ""),                 # I: 型番
            "",                                      # J: 購入在庫
            "",                                      # K: 売価最新更新日
            "",                                      # L: 売価90日変化率
            r.get("buy_box_seller", ""),              # M: BBセラー
            "",                                      # N: 出品許可
            r.get("sales_rank_drops_30"),             # O: 30キーパ
            r.get("fba_seller_count"),                # P: セラー数（FBA）
            r.get("fbm_seller_count"),                # Q: セラー数（無在庫）
            "",                                      # R: 販売数
            "",                                      # S: 初回仕入れ
            r.get("buy_box_price_usd"),               # T: 売価USD
            r.get("purchase_plus_shipping_usd"),      # U: 仕入＋送料（FBA）
            r.get("amazon_fees_usd"),                 # V: Amazon手数料
            r.get("wholesale_price"),                 # W: 仕入値
            r.get("package_weight_g"),                # X: Weight（FBA）
            r.get("profit_usd"),                      # Y: 損益
            r.get("profit_jpy"),                      # Z: 利益額
            r.get("profit_margin"),                   # AA: 利益率
        ])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _to_float(val) -> float:
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def _to_int(val) -> int:
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return 1
