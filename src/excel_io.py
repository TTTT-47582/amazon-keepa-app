"""
Excel 読み書きモジュール

元Excelをコピーし、Sheet1のJANコードとSheet2のKeepaデータを
マッチングしてSheet3に数式ごと自動生成する。
"""

from __future__ import annotations

import io
import shutil
from pathlib import Path
import openpyxl
import pandas as pd


def read_wholesaler_excel(file_buffer) -> pd.DataFrame:
    """
    卸問屋Excelから JAN コード・品名・卸価格などを読み取る。
    ヘッダー名で列を自動検出するため、列位置が異なるExcelにも対応。
    """
    raw = pd.read_excel(
        file_buffer,
        sheet_name=0,
        header=0,
        dtype=str,
        engine="openpyxl",
    )

    # ヘッダー名から列を自動検出
    col_map = {}
    for col in raw.columns:
        col_upper = str(col).strip().upper()
        if "JAN" in col_upper and "jan_code" not in col_map:
            col_map["jan_code"] = col
        elif col_upper in ("品名", "商品名") and "product_name_jp" not in col_map:
            col_map["product_name_jp"] = col
        elif col_upper in ("品番", "型番") and "part_number" not in col_map:
            col_map["part_number"] = col
        elif col_upper in ("定価", "上代") and "retail_price" not in col_map:
            col_map["retail_price"] = col
        elif col_upper in ("卸", "卸価格", "仕入値", "仕入れ値") and "wholesale_price" not in col_map:
            col_map["wholesale_price"] = col
        elif col_upper in ("箱入数", "出荷単位", "入数", "ロット") and "box_qty" not in col_map:
            col_map["box_qty"] = col

    # JAN列が見つからない場合、13桁数字が多い列を探す
    if "jan_code" not in col_map:
        for col in raw.columns:
            vals = raw[col].fillna("").str.strip()
            digit_match = vals.str.match(r"^\d{8,13}$")
            if digit_match.sum() > len(raw) * 0.3:
                col_map["jan_code"] = col
                break

    if "jan_code" not in col_map:
        return pd.DataFrame()

    result = pd.DataFrame()
    result["jan_code"] = raw[col_map["jan_code"]].fillna("").str.strip()
    result["product_name_jp"] = raw[col_map.get("product_name_jp", result.columns[0])].fillna("") if "product_name_jp" in col_map else ""
    result["part_number"] = raw[col_map.get("part_number", result.columns[0])].fillna("") if "part_number" in col_map else ""
    result["retail_price"] = pd.to_numeric(raw[col_map["retail_price"]], errors="coerce").fillna(0) if "retail_price" in col_map else 0
    result["wholesale_price"] = pd.to_numeric(raw[col_map["wholesale_price"]], errors="coerce").fillna(0) if "wholesale_price" in col_map else 0
    result["box_qty"] = pd.to_numeric(raw[col_map["box_qty"]], errors="coerce").fillna(1).astype(int) if "box_qty" in col_map else 1

    result = result[result["jan_code"].str.match(r"^\d{8,}$", na=False)].copy()
    return result.reset_index(drop=True)


def build_ean_to_sheet2_row(file_buffer, sheet_name: str | int = 1) -> dict[str, int]:
    """
    Sheet2（Keepaエクスポート）を読み、EAN → Sheet2の行番号（1-indexed）のマッピングを作る。
    同じEANに複数行がある場合、Buy Box価格が最も高い行を優先。
    """
    raw = pd.read_excel(
        file_buffer, sheet_name=sheet_name,
        header=0, dtype=str, engine="openpyxl",
    )

    ean_to_row: dict[str, int] = {}
    ean_to_price: dict[str, float] = {}

    for idx, row in raw.iterrows():
        ean_raw = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not ean_raw:
            continue
        excel_row = idx + 2  # pandas 0-indexed → Excel 1-indexed + header
        buy_box = _safe_float(row.iloc[14]) or 0

        # カンマ区切りの複数EANに対応
        for ean in ean_raw.split(","):
            ean = ean.strip().replace(" ", "")
            if not ean or not ean.isdigit() or len(ean) < 8:
                continue
            old_price = ean_to_price.get(ean, -1)
            if buy_box > old_price:
                ean_to_row[ean] = excel_row
                ean_to_price[ean] = buy_box

    return ean_to_row


def generate_sheet3(
    source_path: str | Path,
    output_path: str | Path,
    sheet2_name: str = "2026-01-13",
):
    """
    元Excelをコピーし、Sheet1のJANコードをSheet2とマッチングして
    Sheet3にSheet2への参照数式を自動生成する。

    Args:
        source_path: 元Excelファイルのパス
        output_path: 出力先パス（コピー）
        sheet2_name: Keepaエクスポートのシート名
    """
    source_path = Path(source_path)
    output_path = Path(output_path)

    # ① 元Excelをコピー
    shutil.copy2(source_path, output_path)

    # ② Sheet1からJANコード読み取り
    with open(source_path, "rb") as f:
        df = read_wholesaler_excel(f)

    # ③ EAN → Sheet2行番号のマッピング
    with open(source_path, "rb") as f:
        ean_to_row = build_ean_to_sheet2_row(f, sheet_name=1)

    # ④ Sheet2からタイトル（col1）を取得（数式参照できない固定値用）
    with open(source_path, "rb") as f:
        sheet2_raw = pd.read_excel(
            f, sheet_name=1, header=0, dtype=str, engine="openpyxl",
            usecols=[0, 1, 2],
        )

    row_to_title: dict[int, str] = {}
    for idx, row in sheet2_raw.iterrows():
        excel_row = idx + 2
        title = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        row_to_title[excel_row] = title

    # ⑤ コピーしたExcelのSheet3を書き換え
    wb = openpyxl.load_workbook(output_path)

    if "Sheet3" in wb.sheetnames:
        ws = wb["Sheet3"]
    else:
        ws = wb.create_sheet("Sheet3")

    # Row3以降のデータをクリア（Row1-2のヘッダーは残す）
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # ⑥ JANコードごとにSheet3の行を生成
    sq = f"'{sheet2_name}'"  # シート名参照用（数式内）
    out_row = 3
    unique_jans = df["jan_code"].unique()

    matched = 0
    for jan in unique_jans:
        if jan not in ean_to_row:
            continue

        s2r = ean_to_row[jan]  # Sheet2の行番号
        n = out_row
        matched += 1

        # Sheet2への参照（元Sheet3と同一パターン）
        ws.cell(row=n, column=4, value=f"={sq}!AR{s2r}")     # D: 商品名日本語
        title = row_to_title.get(s2r, "")
        ws.cell(row=n, column=5, value=title)                  # E: タイトル（固定値）
        ws.cell(row=n, column=7, value=f"={sq}!B{s2r}")       # G: ASIN
        ws.cell(row=n, column=9, value=f"={sq}!E{s2r}")       # I: 型番/PartNumber
        ws.cell(row=n, column=11, value=1)                     # K: セット内容
        ws.cell(row=n, column=14, value=f"={sq}!AM{s2r}")     # N: 卸（Keepaエクスポート）
        ws.cell(row=n, column=15, value=f"={sq}!F{s2r}")      # O: 30日ランク下落
        ws.cell(row=n, column=16, value=f"={sq}!Y{s2r}")      # P: FBAセラー数
        ws.cell(row=n, column=19, value=5)                     # S: 初回仕入れ
        ws.cell(row=n, column=20, value=f"={sq}!O{s2r}")      # T: Buy Box価格
        ws.cell(row=n, column=22, value=f"={sq}!AQ{s2r}")     # V: FBA手数料
        ws.cell(row=n, column=24, value=f"={sq}!AD{s2r}")     # X: Weight

        # 計算式（元Sheet3と完全に同一）
        ws.cell(row=n, column=23, value=f"=K{n}*N{n}*1.1")                # W: 仕入値
        ws.cell(row=n, column=21, value=f"=(W{n}+(X{n}*$AI$2))/$AH$2")    # U: 仕入＋送料
        ws.cell(row=n, column=25, value=f"=(T{n}-U{n}-V{n})")             # Y: 損益
        ws.cell(row=n, column=26, value=f"=(O{n}/(P{n}+1))*Y{n}")         # Z: 利益額
        ws.cell(row=n, column=27, value=f"=Y{n}/T{n}")                    # AA: 利益率
        ws.cell(row=n, column=28, value=f"=Y{n}*S{n}")                    # AB: 利益額2
        ws.cell(row=n, column=29, value=f"=S{n}*T{n}*$AI$2")             # AC: 売上
        ws.cell(row=n, column=30, value=f"=V{n}*$AI$2*S{n}")             # AD: 手数料
        ws.cell(row=n, column=31, value=f"=S{n}*W{n}")                    # AE: 原価
        ws.cell(row=n, column=32, value=f"=S{n}*X{n}*$AI$2")             # AF: 原価送料
        ws.cell(row=n, column=33, value=f"=AB{n}*$AH$2")                  # AG: 利益/円

        out_row += 1

    wb.save(output_path)
    wb.close()

    return {"total_jans": len(unique_jans), "matched": matched, "output_path": str(output_path)}


def generate_sheet3_from_api(
    wholesaler_df: pd.DataFrame,
    keepa_results: dict[str, dict],
    output_path: str | Path,
    exchange_rate: float = 150.0,
):
    """
    Keepa API結果からExcelを新規作成し、Sheet3に計算値を書き込む。
    Sheet2がないExcel用（API取得モード）。
    """
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    output_path = Path(output_path)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "リサーチ結果"

    # ヘッダー
    _BLUE = PatternFill("solid", fgColor="0000FF")
    _GREEN = PatternFill("solid", fgColor="B6D7A8")
    _RED = PatternFill("solid", fgColor="FF0000")
    _LIME = PatternFill("solid", fgColor="00FF00")
    _ORANGE = PatternFill("solid", fgColor="F9CB9C")
    _HEADER_FONT = Font(bold=True, size=10)

    headers = [
        (1, "重複確認", _BLUE), (2, "日付", _BLUE), (3, "購入先問屋", _BLUE),
        (4, "商品名", _GREEN), (5, "タイトル", _GREEN), (6, "SKU", _GREEN),
        (7, "ASIN", _GREEN), (8, "備考", _GREEN), (9, "型番", _GREEN),
        (10, "購入在庫", _GREEN), (11, "売価最新更新日", _GREEN),
        (12, "セット内容", _GREEN), (13, "BBセラー", _GREEN),
        (14, "卸価格", _GREEN), (15, "30キーパ", _GREEN),
        (16, "セラー数（FBA）", _GREEN), (17, "セラー数（FBM）", _GREEN),
        (18, "販売数", PatternFill("solid", fgColor="FFF2CC")),
        (19, "初回仕入れ", _BLUE), (20, "売価USD", _GREEN),
        (21, "仕入＋送料（FBA）", _ORANGE), (22, "Amazon手数料", _GREEN),
        (23, "仕入値", _GREEN), (24, "Weight（FBA）", _GREEN),
        (25, "損益", _RED), (26, "利益額", _RED), (27, "利益率", _RED),
        (28, "利益額2", _LIME), (29, "売上", _LIME), (30, "手数料", _LIME),
        (31, "原価", _LIME), (32, "原価送料", _LIME), (33, "利益/円", _LIME),
    ]
    for col_num, text, fill in headers:
        cell = ws.cell(row=1, column=col_num, value=text)
        cell.font = _HEADER_FONT
        cell.fill = fill
        ws.column_dimensions[get_column_letter(col_num)].width = 13

    # 定数
    AH, AI = exchange_rate, 3
    ws.cell(row=1, column=34, value=AH)
    ws.cell(row=1, column=35, value=AI)

    from datetime import date
    today = date.today().isoformat()

    unique_jans = wholesaler_df["jan_code"].unique()
    matched = 0
    out_row = 2

    for jan in unique_jans:
        kp = keepa_results.get(jan)
        if not kp or not kp.get("found"):
            continue

        ws_data = wholesaler_df[wholesaler_df["jan_code"] == jan].iloc[0]
        matched += 1
        n = out_row

        N = ws_data.get("wholesale_price", 0)
        T = kp.get("buy_box_price_usd")
        V = kp.get("total_amazon_fee_usd")
        X = kp.get("package_weight_g") or 0
        O = kp.get("sales_rank_drops_30")
        P = kp.get("fba_seller_count") or 0
        S = 5
        L = 1
        W = L * N * 1.1
        U = (W + X * AI) / AH if AH > 0 else 0

        ws.cell(row=n, column=2, value=today)
        ws.cell(row=n, column=4, value=ws_data.get("product_name_jp", ""))
        ws.cell(row=n, column=5, value=kp.get("title", ""))
        ws.cell(row=n, column=7, value=kp.get("asin", ""))
        ws.cell(row=n, column=9, value=ws_data.get("part_number", ""))
        ws.cell(row=n, column=12, value=L)
        ws.cell(row=n, column=13, value=kp.get("buy_box_seller", ""))
        ws.cell(row=n, column=14, value=N)
        ws.cell(row=n, column=15, value=O)
        ws.cell(row=n, column=16, value=P)
        ws.cell(row=n, column=17, value=kp.get("fbm_seller_count"))
        ws.cell(row=n, column=19, value=S)
        ws.cell(row=n, column=20, value=T)
        ws.cell(row=n, column=22, value=V)
        ws.cell(row=n, column=24, value=X if X > 0 else None)

        # 計算値を直接書き込み
        ws.cell(row=n, column=23, value=round(W) if W else None)
        ws.cell(row=n, column=21, value=round(U, 2) if U else None)

        if T and T > 0 and V is not None:
            Y = T - U - V
            Z = (O / (P + 1)) * Y if O else 0
            AA = Y / T
            AB = Y * S
            AG = AB * AH
            ws.cell(row=n, column=25, value=round(Y, 2))
            ws.cell(row=n, column=26, value=round(Z, 2))
            c = ws.cell(row=n, column=27, value=round(AA, 4))
            c.number_format = "0.00%"
            ws.cell(row=n, column=28, value=round(AB, 2))
            ws.cell(row=n, column=29, value=round(S * T * AI, 2))
            ws.cell(row=n, column=30, value=round(V * AI * S, 2))
            ws.cell(row=n, column=31, value=round(S * W, 2))
            ws.cell(row=n, column=32, value=round(S * X * AI, 2))
            ws.cell(row=n, column=33, value=round(AG, 2))

        out_row += 1

    wb.save(output_path)
    wb.close()
    return {"total_jans": len(unique_jans), "matched": matched, "output_path": str(output_path)}


def build_dashboard_data(
    wholesaler_df: pd.DataFrame,
    ean_to_row: dict[str, int],
    file_buffer,
    sheet_name: str | int = 1,
    exchange_rate: float = 150.0,
) -> pd.DataFrame:
    """
    Sheet1 + Sheet2 のデータを結合し、利益計算済みのDataFrameを返す。
    ダッシュボード表示・フィルタリング用。
    """
    raw = pd.read_excel(
        file_buffer, sheet_name=sheet_name,
        header=0, dtype=str, engine="openpyxl",
    )

    # Sheet2の行番号 → データ辞書
    row_to_data: dict[int, dict] = {}
    for idx, row in raw.iterrows():
        excel_row = idx + 2
        row_to_data[excel_row] = {
            "title": str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else "",
            "asin": str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else "",
            "buy_box_usd": _safe_float(row.iloc[14]),
            "drops_30": _safe_float(row.iloc[5]),
            "fba_sellers": _safe_float(row.iloc[24]) or 0,
            "fbm_sellers": _safe_float(row.iloc[25]) or 0,
            "weight_g": _safe_float(row.iloc[29]) or 0,
            "amazon_fee": _safe_float(row.iloc[42]) if len(row) > 42 else None,
            "wholesale_keepa": _safe_float(row.iloc[38]) if len(row) > 38 else None,
        }

    AH = exchange_rate
    AI = 3
    results = []

    for jan in wholesaler_df["jan_code"].unique():
        if jan not in ean_to_row:
            continue
        s2r = ean_to_row[jan]
        s2 = row_to_data.get(s2r)
        if not s2:
            continue

        ws = wholesaler_df[wholesaler_df["jan_code"] == jan].iloc[0]

        # N値（Keepa卸優先、乖離大なら卸データ）
        keepa_ws = s2.get("wholesale_keepa") or 0
        input_ws = ws.get("wholesale_price", 0)
        if keepa_ws > 0 and input_ws > 0:
            ratio = max(keepa_ws, input_ws) / min(keepa_ws, input_ws)
            N = keepa_ws if ratio < 3 else input_ws
        else:
            N = keepa_ws or input_ws

        T = s2["buy_box_usd"]
        V = s2["amazon_fee"]
        X = s2["weight_g"]
        W = N * 1.1
        U = (W + X * AI) / AH if AH > 0 else 0
        Y = (T - U - V) if T and V else None
        P = s2["fba_sellers"]
        O = s2["drops_30"]
        AA = (Y / T) if Y is not None and T and T > 0 else None

        results.append({
            "JAN": jan,
            "ASIN": s2["asin"],
            "タイトル": s2["title"][:60],
            "商品名": ws.get("product_name_jp", "")[:30],
            "売価USD": T,
            "仕入値(税込)": round(W) if W else None,
            "仕入+送料USD": round(U, 2) if U else None,
            "手数料USD": V,
            "損益USD": round(Y, 2) if Y is not None else None,
            "利益率": round(AA, 4) if AA is not None else None,
            "30日drop": O,
            "FBAセラー": int(P),
            "FBMセラー": int(s2["fbm_sellers"]),
            "重量g": int(X) if X else None,
        })

    return pd.DataFrame(results)


def _safe_float(val) -> float | None:
    try:
        v = float(val)
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None
