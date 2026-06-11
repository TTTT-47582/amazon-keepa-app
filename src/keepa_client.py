"""
Keepa API ラッパーモジュール

注意: product_finder は Keepa の Professional プラン以上が必要です。
     API トークンを節約するため検索結果は 1 時間キャッシュします。
"""

from __future__ import annotations

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
    Keepa Product Finder で日本 Amazon から商品を検索する。
    """
    api = _get_api(_resolve_api_key(api_key))

    # Product Finder パラメータ（価格は Keepa 単位: 円 × 100）
    # ※ 評価・レビュー数は product_finder 非対応のため取得後にフィルタリングする
    product_parms = {
        "current_SALES_gte": 1,
        "current_SALES_lte": params["sales_rank_max"],
        "current_NEW_gte": params["price_min"] * _KEEPA_PRICE_DIVISOR,
        "current_NEW_lte": params["price_max"] * _KEEPA_PRICE_DIVISOR,
    }

    st.info(f"🔍 検索パラメータ: {product_parms}")

    # Product Finder は ASIN のリストを返す
    asins = api.product_finder(product_parms, domain="JP", wait=True)
    st.info(f"📦 取得 ASIN 数: {len(asins) if asins else 0}")

    if not asins:
        return []

    # 評価・レビューフィルタ後に必要な件数を確保するため多めに取得する
    fetch_count = min(len(asins), params.get("max_results", 20) * 3)
    asins = list(asins)[:fetch_count]

    # ASIN から商品詳細を一括取得
    products_raw = api.query(asins, domain="JP", history=False, wait=True)
    st.info(f"📋 商品詳細取得数: {len(products_raw) if products_raw else 0}")

    # 最初の商品の stats 構造をデバッグ表示
    if products_raw:
        p0 = products_raw[0]
        stats0 = p0.get("stats") or {}
        st.write("🔎 stats keys:", list(stats0.keys()) if stats0 else "なし")
        st.write("🔎 stats['current']:", stats0.get("current"))
        st.write("🔎 csv keys count:", len(p0.get("csv") or []))

    results = _parse_products(products_raw)
    st.info(f"✅ フィルタ後: {len(results)} 件")
    results = [
        p for p in results
        if p["rating"] >= params["rating_min"]
        and p["review_count"] >= params["review_count_min"]
    ]
    st.info(f"✅ 評価・レビューフィルタ後: {len(results)} 件")
    return results[: params.get("max_results", 20)]


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
