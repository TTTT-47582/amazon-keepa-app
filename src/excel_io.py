"""
Excel 読み書きモジュール

卸問屋のJANコードExcelを読み込み、
Keepa結果をSheet3形式で出力する。
"""

from __future__ import annotations

import io
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import pandas as pd


def read_wholesaler_excel(file_buffer) -> pd.DataFrame:
    """
    卸問屋Excelから JAN コード・品名・卸価格などを読み取る。
    必要列のみ読み込むことで大容量ファイルでも高速。
    """
    raw = pd.read_excel(
        file_buffer,
        sheet_name=0,
        usecols=[3, 4, 5, 7, 8, 9],  # D, E, F, H, I, J 列のみ
        header=0,
        dtype=str,
        engine="openpyxl",
    )
    raw.columns = [
        "jan_code", "product_name_jp", "part_number",
        "retail_price", "wholesale_price", "box_qty",
    ]

    raw["jan_code"] = raw["jan_code"].fillna("").str.strip()
    raw = raw[raw["jan_code"].str.match(r"^\d{8,}$", na=False)].copy()
    raw["retail_price"] = pd.to_numeric(raw["retail_price"], errors="coerce").fillna(0)
    raw["wholesale_price"] = pd.to_numeric(raw["wholesale_price"], errors="coerce").fillna(0)
    raw["box_qty"] = pd.to_numeric(raw["box_qty"], errors="coerce").fillna(1).astype(int)
    raw["product_name_jp"] = raw["product_name_jp"].fillna("")
    raw["part_number"] = raw["part_number"].fillna("")

    return raw.reset_index(drop=True)


def read_keepa_export_sheet(file_buffer, sheet_name: str | int = 1) -> dict[str, dict]:
    """
    Excel内のKeepaエクスポートシート（シート2）を読み込み、
    EANコードをキーにした辞書を返す。APIを使わずに既存データを再利用できる。

    Returns:
        {ean_code: {title, asin, buy_box_price_usd, ...}} のマッピング
    """
    raw = pd.read_excel(
        file_buffer,
        sheet_name=sheet_name,
        header=0,
        dtype=str,
        engine="openpyxl",
    )

    ean_to_product: dict[str, dict] = {}

    for _, row in raw.iterrows():
        ean = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not ean or not ean.replace(" ", "").isdigit():
            continue

        ean = ean.replace(" ", "")

        buy_box_raw = _safe_float(row.iloc[14])  # Col15: Buy Box 現在価格
        weight_raw = _safe_float(row.iloc[29])     # Col30: パッケージ重さ(g)

        fba_count = _safe_int(row.iloc[24]) or 0   # Col25: FBA数（空=0）
        fbm_count = _safe_int(row.iloc[25]) or 0   # Col26: FBM数（空=0）
        drops_30 = _safe_int(row.iloc[5])            # Col6: 30日ランク下落

        # Col43(AQ): FBA手数料 = FBA Pick&Pack + 紹介料の合算値
        total_fee = _safe_float(row.iloc[42]) if len(row) > 42 else None
        if total_fee is None:
            fba_pp = _safe_float(row.iloc[20]) or 0
            referral = _safe_float(row.iloc[21]) or 0
            total_fee = (fba_pp + referral) if (fba_pp + referral) > 0 else None

        # Col39(AM): Keepaエクスポートの卸価格（元Sheet3のN列参照元）
        keepa_wholesale = _safe_float(row.iloc[38]) if len(row) > 38 else None

        buy_box_seller = str(row.iloc[16]).strip() if pd.notna(row.iloc[16]) else ""

        product = {
            "found": buy_box_raw is not None and buy_box_raw > 0,
            "asin": str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "",
            "title": str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else "",
            "buy_box_price_usd": buy_box_raw,
            "buy_box_seller": buy_box_seller if buy_box_seller != "-" else "",
            "fba_seller_count": fba_count,
            "fbm_seller_count": fbm_count,
            "sales_rank_drops_30": drops_30,
            "package_weight_g": int(weight_raw) if weight_raw else None,
            "total_amazon_fee_usd": total_fee,
            "keepa_wholesale": keepa_wholesale,
        }

        # 同じEANに複数ASINがある場合、Buy Box価格が最も高い商品を優先
        existing = ean_to_product.get(ean)
        if existing is None:
            ean_to_product[ean] = product
        else:
            old_price = existing.get("buy_box_price_usd") or 0
            new_price = buy_box_raw or 0
            if new_price > old_price:
                ean_to_product[ean] = product

    return ean_to_product


def _safe_float(val) -> float | None:
    try:
        v = float(val)
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None


def _safe_int(val) -> int | None:
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None


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
    Keepa 取得結果を Sheet3 形式の Excel として出力する。
    元Sheet3と同一の計算式を埋め込み、Excel上で値を変更すると自動再計算される。
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

    # ── Row2: メインヘッダー（配色 + 太字） + 定数セル ──
    for col_num, text, fill, _ in _ROW2_HEADERS:
        cell = ws.cell(row=2, column=col_num, value=text)
        cell.font = _HEADER_FONT
        cell.fill = fill
        cell.alignment = Alignment(wrap_text=True)

    # 定数（元Sheet3と同一: AH2=為替レート, AI2=送料単価/g）
    exchange_rate = 150.0
    if results:
        exchange_rate = results[0].get("_exchange_rate", 150.0)
    ws.cell(row=2, column=34, value=exchange_rate)   # AH2: 為替レート
    ws.cell(row=2, column=35, value=3)               # AI2: 送料単価(円/g)
    ws.cell(row=2, column=36, value="ポンド")        # AJ2
    ws.cell(row=2, column=37, value=453.59)          # AK2
    ws.cell(row=2, column=38, value=0.23)            # AL2
    ws.cell(row=2, column=39, value=106.14006)       # AM2

    # ── データ行（Row3〜）: 値 + Excel計算式 ──
    today = date.today().isoformat()

    for row_idx, r in enumerate(results, start=3):
        n = row_idx  # 行番号（数式用）

        ws.cell(row=n, column=1, value="")                              # A: 重複確認
        ws.cell(row=n, column=2, value=today)                           # B: 日付
        ws.cell(row=n, column=3, value=wholesaler_name)                 # C: 購入先問屋
        ws.cell(row=n, column=4, value=r.get("product_name_jp", ""))    # D: 商品名
        ws.cell(row=n, column=5, value=r.get("title", ""))              # E: タイトル
        ws.cell(row=n, column=6, value="")                              # F: SKU
        ws.cell(row=n, column=7, value=r.get("asin", ""))               # G: ASIN
        h_val = "" if r.get("found") else "US未出品"
        ws.cell(row=n, column=8, value=h_val)                           # H: 備考
        ws.cell(row=n, column=9, value=r.get("part_number", ""))        # I: 型番
        ws.cell(row=n, column=10, value="")                             # J: 購入在庫
        ws.cell(row=n, column=11, value="")                             # K: 売価最新更新日
        ws.cell(row=n, column=12, value=1)                              # L: セット内容（デフォルト1）
        ws.cell(row=n, column=13, value=r.get("buy_box_seller", ""))    # M: BBセラー
        # N: 1本単価税込 - Keepa卸と卸データを比較し、乖離が大きければ卸データを使用
        keepa_ws = r.get("keepa_wholesale") or 0
        input_ws = r.get("wholesale_price") or 0
        if keepa_ws > 0 and input_ws > 0:
            ratio = max(keepa_ws, input_ws) / min(keepa_ws, input_ws)
            n_val = keepa_ws if ratio < 3 else input_ws
        else:
            n_val = keepa_ws or input_ws
        ws.cell(row=n, column=14, value=n_val)
        ws.cell(row=n, column=15, value=r.get("sales_rank_drops_30"))   # O: 30キーパ
        ws.cell(row=n, column=16, value=r.get("fba_seller_count"))      # P: セラー数(FBA)
        ws.cell(row=n, column=17, value=r.get("fbm_seller_count"))      # Q: セラー数(無在庫)
        ws.cell(row=n, column=18, value="")                             # R: 販売数
        ws.cell(row=n, column=19, value=5)                              # S: 初回仕入れ（デフォルト5）
        _set_num(ws, n, 20, r.get("buy_box_price_usd"))                 # T: 売価USD
        # U: 仕入＋送料（FBA）= (仕入値 + Weight×送料単価) / 為替レート
        ws.cell(row=n, column=21, value=f"=(W{n}+(X{n}*$AI$2))/$AH$2")
        ws[f"U{n}"].number_format = "0.00"
        # V: Amazon手数料（Keepaエクスポートのcol43 FBA手数料 = FBA+紹介料合算）
        _set_num(ws, n, 22, r.get("total_amazon_fee_usd") or r.get("amazon_fees_usd"))
        # W: 仕入値 = セット内容 × 1本単価 × 1.1（税込）
        ws.cell(row=n, column=23, value=f"=L{n}*N{n}*1.1")
        ws[f"W{n}"].number_format = "0"
        ws.cell(row=n, column=24, value=r.get("package_weight_g"))      # X: Weight(FBA)
        # Y: 損益 = 売価 - 仕入送料 - Amazon手数料
        ws.cell(row=n, column=25, value=f"=(T{n}-U{n}-V{n})")
        ws[f"Y{n}"].number_format = "0.00"
        # Z: 利益額 = (30日drop / (FBAセラー+1)) × 損益
        ws.cell(row=n, column=26, value=f"=IFERROR((O{n}/(P{n}+1))*Y{n},0)")
        ws[f"Z{n}"].number_format = "0.00"
        # AA: 利益率 = 損益 / 売価
        ws.cell(row=n, column=27, value=f"=IFERROR(Y{n}/T{n},0)")
        ws[f"AA{n}"].number_format = "0.00%"
        # AB: 利益額2 = 損益 × 初回仕入れ
        ws.cell(row=n, column=28, value=f"=Y{n}*S{n}")
        ws[f"AB{n}"].number_format = "0.00"
        # AC: 売上 = 初回仕入れ × 売価 × 送料単価
        ws.cell(row=n, column=29, value=f"=S{n}*T{n}*$AI$2")
        ws[f"AC{n}"].number_format = "0.00"
        # AD: 手数料 = Amazon手数料 × 送料単価 × 初回仕入れ
        ws.cell(row=n, column=30, value=f"=V{n}*$AI$2*S{n}")
        ws[f"AD{n}"].number_format = "0.00"
        # AE: 原価 = 初回仕入れ × 仕入値
        ws.cell(row=n, column=31, value=f"=S{n}*W{n}")
        ws[f"AE{n}"].number_format = "0.00"
        # AF: 原価送料 = 初回仕入れ × Weight × 送料単価
        ws.cell(row=n, column=32, value=f"=S{n}*X{n}*$AI$2")
        ws[f"AF{n}"].number_format = "0.00"
        # AG: 利益/円 = 利益額2 × 為替レート
        ws.cell(row=n, column=33, value=f"=AB{n}*$AH$2")
        ws[f"AG{n}"].number_format = "0.00"

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
