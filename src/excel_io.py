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
    """卸問屋Excelから JAN コード・品名・卸価格などを読み取る。"""
    raw = pd.read_excel(
        file_buffer,
        sheet_name=0,
        usecols=[3, 4, 5, 7, 8, 9],
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
        ean = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        if not ean or not ean.replace(" ", "").isdigit():
            continue
        ean = ean.replace(" ", "")
        excel_row = idx + 2  # pandas 0-indexed → Excel 1-indexed + header

        buy_box = _safe_float(row.iloc[14]) or 0

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


def _safe_float(val) -> float | None:
    try:
        v = float(val)
        return v if v > 0 else None
    except (TypeError, ValueError):
        return None
