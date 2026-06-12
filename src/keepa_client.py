"""
Keepa API ラッパーモジュール

注意: product_finder は Keepa の Professional プラン以上が必要です。
"""

from __future__ import annotations

import os
import keepa
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# JPY は円そのままで格納（÷1）、USD はセント格納のため÷100
_DIVISOR_JP = 1
_DIVISOR_US = 100

# csv配列の価格タイプインデックス
_IDX_AMAZON   = 0   # Amazon 直販
_IDX_NEW      = 1   # 新品（マーケットプレイス）
_IDX_SALES    = 3   # 販売ランク
_IDX_NEW_FBA  = 10  # FBA 新品
_IDX_BUYBOX   = 18  # バイボックス価格


@st.cache_resource
def _get_api(api_key: str) -> keepa.Keepa:
    """Keepa API クライアントを初期化する（同じキーならキャッシュを再利用）"""
    return keepa.Keepa(api_key)


def _resolve_api_key(api_key: str | None) -> str:
    """UI 入力キー → .env の順で有効なキーを返す。どちらもなければ例外を投げる"""
    key = api_key or os.getenv("KEEPA_API_KEY", "")
    if not key:
        raise ValueError(
            "Keepa API キーが設定されていません。"
            "サイドバーにキーを入力するか、.env に KEEPA_API_KEY を設定してください。"
        )
    return key


def search_products(params: dict, api_key: str = "") -> list[dict]:
    """
    日本 Amazon で商品を検索し、米国 Amazon にも出品されている商品のみ返す。
    EAN/UPC で JP↔US をクロスチェックする。
    """
    api = _get_api(_resolve_api_key(api_key))

    # ① Product Finder で JP Amazon から ASIN を取得
    # JPY は円そのまま格納のため ×100 不要
    product_parms = {
        "current_SALES_gte": 1,
        "current_SALES_lte": params["sales_rank_max"],
        "current_NEW_gte": params["price_min"],
        "current_NEW_lte": params["price_max"],
    }
    asins = api.product_finder(product_parms, domain="JP", wait=True)
    if not asins:
        return []

    # フィルタ後に必要件数を確保するため多めに取得
    fetch_count = min(len(asins), params.get("max_results", 20) * 2)
    asins = list(asins)[:fetch_count]

    # ② JP 商品詳細を取得（キャッシュ優先で高速化）
    jp_raw = api.query(asins, domain="JP", history=True, update=0, wait=True)
    jp_results = _parse_jp_products(jp_raw)

    # ③ EAN リストを収集して US Amazon で一括検索（チェックボックスONのときのみ）
    ean_to_us_price: dict[str, float] = {}
    if params.get("check_us_listing", True):
        all_eans: list[str] = []
        for p in jp_results:
            all_eans.extend(p.get("ean_list") or [])
        all_eans = list(dict.fromkeys(all_eans))[:100]  # 重複除去・上限100件

        if all_eans:
            us_raw = api.query(all_eans, domain="US", history=True, update=0, wait=True)
            for us_p in (us_raw or []):
                if not us_p:
                    continue
                us_price = _get_us_price(us_p)
                if us_price is None:
                    continue
                for ean in (us_p.get("eanList") or []):
                    ean_to_us_price[str(ean)] = us_price

    # ④ US チェックONなら出品あり商品のみ / OFFなら全商品を対象にする
    matched: list[dict] = []
    if params.get("check_us_listing", True):
        for p in jp_results:
            for ean in (p.get("ean_list") or []):
                if str(ean) in ean_to_us_price:
                    p["us_actual_price_usd"] = ean_to_us_price[str(ean)]
                    p["has_us_listing"] = True
                    matched.append(p)
                    break
    else:
        matched = jp_results  # チェックOFFは全商品をそのまま使う

    # ⑤ 評価・レビューフィルタ（0 = データなしはスキップ）
    filtered = [
        p for p in matched
        if (p["rating"] == 0 or p["rating"] >= params["rating_min"])
        and (p["review_count"] == 0 or p["review_count"] >= params["review_count_min"])
    ]

    return filtered[: params.get("max_results", 20)]


def _parse_jp_products(products_raw: list) -> list[dict]:
    """Keepa の JP 生レスポンスを扱いやすい辞書形式に変換する"""
    results = []

    for p in products_raw or []:
        if not p:
            continue

        csv = p.get("csv") or []

        # 現在の新品価格（新品 → FBA → Amazon 直販 の順で試みる）
        price_jpy = _latest_price(csv, _IDX_NEW, _DIVISOR_JP)
        if price_jpy is None:
            price_jpy = _latest_price(csv, _IDX_NEW_FBA, _DIVISOR_JP)
        if price_jpy is None:
            price_jpy = _latest_price(csv, _IDX_AMAZON, _DIVISOR_JP)
        if price_jpy is None:
            continue  # 価格不明はスキップ

        # 現在の販売ランク
        rank_raw = _latest_price(csv, _IDX_SALES, 1)
        sales_rank = int(rank_raw) if rank_raw and rank_raw > 0 else None

        # 商品サムネイル URL（最初の画像）
        images_csv = p.get("imagesCSV") or ""
        first_img = images_csv.split(",")[0] if images_csv else ""
        image_url = (
            f"https://images-na.ssl-images-amazon.com/images/I/{first_img}"
            if first_img else None
        )

        # カテゴリ名（最末端）
        category_tree = p.get("categoryTree") or []
        category = category_tree[-1].get("name", "") if category_tree else ""

        results.append({
            "asin": p.get("asin", ""),
            "title": p.get("title") or "タイトル不明",
            "category": category,
            "image_url": image_url,
            "rating": (p.get("avgRating") or 0) / 10,
            "review_count": p.get("reviewCount") or 0,
            "sales_rank": sales_rank,
            "price_jp_jpy": price_jpy,
            "weight_g": p.get("packageWeight"),
            "ean_list": [str(e) for e in (p.get("eanList") or []) if e],
            "has_us_listing": False,
            "us_actual_price_usd": None,
        })

    return results


def _get_us_price(us_product: dict) -> float | None:
    """US 商品からバイボックス → FBA → 新品 の順で価格（USD）を取得する"""
    csv = us_product.get("csv") or []
    price = _latest_price(csv, _IDX_BUYBOX, _DIVISOR_US)
    if price is None:
        price = _latest_price(csv, _IDX_NEW_FBA, _DIVISOR_US)
    if price is None:
        price = _latest_price(csv, _IDX_NEW, _DIVISOR_US)
    if price is None:
        price = _latest_price(csv, _IDX_AMAZON, _DIVISOR_US)
    return price


def _latest_price(csv: list, price_type: int, divisor: int) -> float | None:
    """
    csv配列の指定タイプから最新価格を取得する。
    keepa csv形式: [time0, price0, time1, price1, ...] のフラット配列。
    最新価格は末尾要素（奇数インデックス位置）。
    """
    if not csv or price_type >= len(csv):
        return None
    history = csv[price_type]
    if not history or len(history) < 2:
        return None
    raw = history[-1]
    if raw is None or raw < 0:
        return None
    return raw / divisor
