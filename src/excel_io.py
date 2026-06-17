"""
Excel 読み書きモジュール

卸問屋のJANコードExcelを読み込み、
Keepa結果をSheet3形式で出力する。
"""

from __future__ import annotations

import io
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter
import pandas as pd


def read_wholesaler_excel(file_buffer) -> pd.DataFrame:
    """
    卸問屋Excelから JAN コード・品名・卸価格などを読み取る。
    シート「商品データ」(先頭シート) の D列=JAN, E列=品名, F列=品番, H列=定価, I列=卸, J列=箱入数。
    """
    wb = openpyxl.load_workbook(file_buffer, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
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


# ── ヘッダー配色（元Sheet3と同一） ──
_BLUE   = PatternFill("solid", fgColor="0000FF")
_GREEN  = PatternFill("solid", fgColor="B6D7A8")
_YELLOW = PatternFill("solid", fgColor="FFF2CC")
_RED    = PatternFill("solid", fgColor="FF0000")
_LIME   = PatternFill("solid", fgColor="00FF00")
_NONE   = PatternFill(fill_type=None)

_HEADER_FONT = Font(bold=True, size=10)
_NORMAL_FONT = Font(size=11)

# Row2 ヘッダー定義: (列番号, ヘッダー名, 背景色, 列幅)
_ROW2_HEADERS = [
    (1,  "重複確認",         _BLUE,   13.0),
    (2,  "日付",             _BLUE,   13.0),
    (3,  "購入先問屋",       _BLUE,   13.0),
    (4,  "商品名",           _GREEN,  13.0),
    (5,  "タイトル",         _GREEN,  13.0),
    (6,  "SKU",              _GREEN,  13.0),
    (7,  "ASIN",             _GREEN,  15.1),
    (8,  "備考",             _GREEN,  13.0),
    (9,  "販売可能品",       _GREEN,  11.6),
    (10, "購入在庫",         _GREEN,  15.4),
    (11, "売価最新更新日",   _GREEN,  13.0),
    (12, "売価90日変化率",   _GREEN,  13.0),
    (13, "BBセラー",         _GREEN,  13.0),
    (14, "出品許可",         _GREEN,   9.0),
    (15, "30キーパ",         _GREEN,  13.0),
    (16, "セラー数（FBA）",  _GREEN,  13.0),
    (17, "セラー数（無在庫）", _GREEN, 13.0),
    (18, "販売数（自社1M）", _YELLOW, 13.0),
    (19, "初回仕入れ",       _BLUE,   13.0),
    (20, "売価USD",          _GREEN,  13.0),
    (21, "仕入＋送料（FBA）", PatternFill("solid", fgColor="F9CB9C"), 13.0),
    (22, "Amazon手数料",     _GREEN,  13.0),
    (23, "仕入値",           _GREEN,   9.5),
    (24, "Weight（FBA）",    _GREEN,  13.0),
    (25, "損益",             _RED,    13.0),
    (26, "利益額",           _RED,    13.0),
    (27, "利益率",           _RED,    13.0),
    (28, "利益額2",          _LIME,   13.0),
    (29, "売上",             _LIME,   13.0),
    (30, "手数料",           _LIME,   13.0),
    (31, "原価",             _LIME,   10.2),
    (32, "原価送料",         _LIME,   11.2),
    (33, "利益/円",          _LIME,   11.0),
]

# Row1 サブヘッダー: (列番号, テキスト)
_ROW1_HEADERS = [
    (9,  "型番"),
    (11, "出荷元販売単位"),
    (12, "セット内容"),
    (13, "単位"),
    (14, "1本単価税込"),
    (15, "1セット料金"),
]


def write_output_excel(results: list[dict], wholesaler_name: str = "") -> io.BytesIO:
    """
    Keepa 取得結果 + 利益計算済みデータを Sheet3 形式の Excel として出力する。
    元ファイルと同一のデザイン（配色・列幅・フォント・数値書式）を再現。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "リサーチ結果"

    # ── 列幅設定 ──
    for col_num, _, _, width in _ROW2_HEADERS:
        ws.column_dimensions[get_column_letter(col_num)].width = width

    # ── Row1: サブヘッダー ──
    for col_num, text in _ROW1_HEADERS:
        cell = ws.cell(row=1, column=col_num, value=text)
        cell.font = _NORMAL_FONT

    # ── Row2: メインヘッダー（配色 + 太字） ──
    for col_num, text, fill, _ in _ROW2_HEADERS:
        cell = ws.cell(row=2, column=col_num, value=text)
        cell.font = _HEADER_FONT
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True)

    # ── データ行（Row3〜） ──
    today = date.today().isoformat()

    for row_idx, r in enumerate(results, start=3):
        exchange_rate = r.get("_exchange_rate", 150.0)
        sell_usd = r.get("buy_box_price_usd")
        profit_usd = r.get("profit_usd")
        wholesale = r.get("wholesale_price") or 0

        ws.cell(row=row_idx, column=1, value="")                          # A: 重複確認
        ws.cell(row=row_idx, column=2, value=today)                       # B: 日付
        ws.cell(row=row_idx, column=3, value=wholesaler_name)             # C: 購入先問屋
        ws.cell(row=row_idx, column=4, value=r.get("product_name_jp", ""))# D: 商品名
        ws.cell(row=row_idx, column=5, value=r.get("title", ""))          # E: タイトル
        ws.cell(row=row_idx, column=6, value="")                          # F: SKU
        ws.cell(row=row_idx, column=7, value=r.get("asin", ""))           # G: ASIN
        h_val = "" if r.get("found") else "US未出品"
        ws.cell(row=row_idx, column=8, value=h_val)                       # H: 備考
        ws.cell(row=row_idx, column=9, value=r.get("part_number", ""))    # I: 型番
        ws.cell(row=row_idx, column=10, value="")                         # J: 購入在庫
        ws.cell(row=row_idx, column=11, value="")                         # K: 売価最新更新日
        ws.cell(row=row_idx, column=12, value="")                         # L: 売価90日変化率
        ws.cell(row=row_idx, column=13, value=r.get("buy_box_seller", ""))# M: BBセラー
        ws.cell(row=row_idx, column=14, value="")                         # N: 出品許可
        ws.cell(row=row_idx, column=15, value=r.get("sales_rank_drops_30"))# O: 30キーパ
        ws.cell(row=row_idx, column=16, value=r.get("fba_seller_count"))  # P: セラー数(FBA)
        ws.cell(row=row_idx, column=17, value=r.get("fbm_seller_count"))  # Q: セラー数(無在庫)
        ws.cell(row=row_idx, column=18, value="")                         # R: 販売数
        ws.cell(row=row_idx, column=19, value="")                         # S: 初回仕入れ
        _set_num(ws, row_idx, 20, sell_usd)                               # T: 売価USD
        _set_num(ws, row_idx, 21, r.get("purchase_plus_shipping_usd"), "0.00")  # U: 仕入＋送料
        _set_num(ws, row_idx, 22, r.get("amazon_fees_usd"))               # V: Amazon手数料
        ws.cell(row=row_idx, column=23, value=wholesale)                  # W: 仕入値
        ws.cell(row=row_idx, column=24, value=r.get("package_weight_g"))  # X: Weight
        _set_num(ws, row_idx, 25, r.get("profit_usd"), "0.00")           # Y: 損益
        _set_num(ws, row_idx, 26, r.get("profit_jpy"), "0.00")           # Z: 利益額
        _set_pct(ws, row_idx, 27, r.get("profit_margin"))                # AA: 利益率

        # AB-AG: 利益額2・売上・手数料・原価・原価送料・利益/円（計算列）
        if sell_usd and profit_usd is not None:
            sell_jpy = sell_usd * exchange_rate
            fees_usd = r.get("amazon_fees_usd") or 0
            pps = r.get("purchase_plus_shipping_usd") or 0
            _set_num(ws, row_idx, 28, profit_usd * 5 if profit_usd else None, "0.00")  # 利益額2 (5個想定)
            _set_num(ws, row_idx, 29, sell_jpy, "0.00")                    # 売上(円)
            _set_num(ws, row_idx, 30, fees_usd * exchange_rate, "0.00")    # 手数料(円)
            _set_num(ws, row_idx, 31, wholesale, "0.00")                   # 原価
            shipping_jpy = ((r.get("package_weight_g") or 500) * 3.0)
            _set_num(ws, row_idx, 32, shipping_jpy, "0.00")               # 原価送料
            profit_jpy_val = r.get("profit_jpy")
            _set_num(ws, row_idx, 33, profit_jpy_val, "0.00")             # 利益/円

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _set_num(ws, row: int, col: int, val, fmt: str = "General"):
    """数値セルを書式付きで設定"""
    cell = ws.cell(row=row, column=col, value=val)
    if fmt != "General":
        cell.number_format = fmt


def _set_pct(ws, row: int, col: int, val):
    """パーセント値セルを設定 (0.15 → 15.0% 表示)"""
    cell = ws.cell(row=row, column=col, value=val)
    cell.number_format = "0.00%"


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
