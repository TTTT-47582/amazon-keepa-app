"""
Keepa API ラッパーモジュール

注意: product_finder は Keepa の Professional プラン以上が必要です。
     API トークンを節約するため検索結果は 1 時間キャッシュします。
"""

import os
import keepa
import streamlit as st
from dotenv import load_dotenv

load_dotenv()

# Keepa は全通貨で価格を (実際の価格 × 100) で格納している
# 例: ¥500 → 50000、$9.99 → 999
_KEEPA_PRICE_DIVISOR = 100

# stats['current'] リストのインデックス（価格タイプ）
_PRICE_AMAZON = 0   # Amazon 直販
_PRICE_NEW = 1      # 新品（サードパーティ）
_PRICE_FBA = 7      # FBA 新品


@st.cache_resource
def _get_api() -> keepa.Keepa:
    """Keepa API クライアントを初期化する（シングルトン）"""
    api_key = os.getenv("KEEPA_API_KEY")
    if not api_key:
        raise ValueError(
            "KEEPA_API_KEY が設定されていません。"
            "プロジェクトルートに .env ファイルを作成して KEEPA_API_KEY を設定してください。"
        )
    return keepa.Keepa(api_key)


@st.cache_data(ttl=3600, show_spinner=False)
def search_products(params: dict) -> list[dict]:
    """
    Keepa Product Finder で日本 Amazon から商品を検索する。
    同じ条件での再検索は 1 時間キャッシュしてトークン消費を抑える。
    """
    api = _get_api()

    # Product Finder パラメータ（価格は Keepa 単位: 円 × 100）
    product_parms = {
        "current_SALES_gte": 1,
        "current_SALES_lte": params["sales_rank_max"],
        "current_NEW_gte": params["price_min"] * _KEEPA_PRICE_DIVISOR,
        "current_NEW_lte": params["price_max"] * _KEEPA_PRICE_DIVISOR,
        "avg_RATING_gte": int(params["rating_min"] * 10),
        "reviewCount_gte": params["review_count_min"],
    }

    # Product Finder は ASIN のリストを返す
    asins = api.product_finder(product_parms, domain="JP", wait=True)

    if not asins:
        return []

    # トークン節約のため最大件数に切り詰める
    asins = list(asins)[: params.get("max_results", 20)]

    # ASIN から商品詳細を一括取得（history=False で価格履歴を省略してトークン節約）
    products_raw = api.query(asins, domain="JP", history=False, wait=True)

    return _parse_products(products_raw)


def _parse_products(products_raw: list) -> list[dict]:
    """Keepa の生レスポンスを扱いやすい辞書形式に変換する"""
    results = []

    for p in products_raw or []:
        if not p:
            continue

        stats = p.get("stats") or {}
        current_prices = stats.get("current") or []

        # 現在の新品価格を取得（サードパーティ → FBA の順で試みる）
        price_jpy = _safe_price(current_prices, _PRICE_NEW)
        if price_jpy is None:
            price_jpy = _safe_price(current_prices, _PRICE_FBA)
        if price_jpy is None:
            # Amazon 直販価格でも試みる
            price_jpy = _safe_price(current_prices, _PRICE_AMAZON)
        if price_jpy is None:
            continue  # 価格が取得できない商品はスキップ

        # 現在の販売ランク（-1 は取得不可）
        sales_rank = stats.get("salesRankCurrent", -1)
        sales_rank = sales_rank if sales_rank and sales_rank > 0 else None

        # 商品サムネイル URL（最初の画像のみ使用）
        images_csv = p.get("imagesCSV") or ""
        first_image_id = images_csv.split(",")[0] if images_csv else ""
        image_url = (
            f"https://images-na.ssl-images-amazon.com/images/I/{first_image_id}"
            if first_image_id
            else None
        )

        # カテゴリ名（最末端カテゴリを表示）
        category_tree = p.get("categoryTree") or []
        category = category_tree[-1].get("name", "") if category_tree else ""

        results.append(
            {
                "asin": p.get("asin", ""),
                "title": p.get("title") or "タイトル不明",
                "category": category,
                "image_url": image_url,
                "rating": (p.get("avgRating") or 0) / 10,
                "review_count": p.get("reviewCount") or 0,
                "sales_rank": sales_rank,
                "price_jp_jpy": price_jpy,
                "weight_g": p.get("packageWeight"),  # None の場合は profit_calc でデフォルト値を使用
            }
        )

    return results


def _safe_price(prices: list, index: int) -> float | None:
    """Keepa の価格リストから安全に価格を取得する（-1 や範囲外は None を返す）"""
    if not prices or index >= len(prices):
        return None
    raw = prices[index]
    if raw is None or raw < 0:
        return None
    return raw / _KEEPA_PRICE_DIVISOR
