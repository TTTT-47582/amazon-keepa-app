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

    # トークン節約のため最大件数に切り詰める（history=True は消費が多いため少なめに）
    fetch_count = min(len(asins), params.get("max_results", 20))
    asins = list(asins)[:fetch_count]

    # history=True で csv（価格履歴）を含む全データを取得する
    # history=False だと stats/csv が返らず価格が取得できないため True が必須
    products_raw = api.query(asins, domain="JP", history=True, wait=True)
    st.info(f"📋 商品詳細取得数: {len(products_raw) if products_raw else 0}")

    results = _parse_products(products_raw)
    st.info(f"✅ パース後: {len(results)} 件")
    # rating=0 は「Keepaに評価データなし」を意味するため、0の場合はフィルタをスキップする
    results = [
        p for p in results
        if (p["rating"] == 0 or p["rating"] >= params["rating_min"])
        and (p["review_count"] == 0 or p["review_count"] >= params["review_count_min"])
    ]
    st.info(f"✅ フィルタ後: {len(results)} 件")
    return results


def _parse_products(products_raw: list) -> list[dict]:
    """Keepa の生レスポンスを扱いやすい辞書形式に変換する"""
    results = []

    for p in products_raw or []:
        if not p:
            continue

        csv = p.get("csv") or []

        # csv はフラット配列 [time0, price0, time1, price1, ...] の形式
        # 末尾の価格（最新値）は csv[type_index][-1]
        # まず新品（index 1）→ FBA（index 10）→ Amazon直販（index 0）の順で試みる
        price_jpy = _latest_price_from_csv(csv, _PRICE_NEW)
        if price_jpy is None:
            price_jpy = _latest_price_from_csv(csv, _PRICE_FBA)
        if price_jpy is None:
            price_jpy = _latest_price_from_csv(csv, _PRICE_AMAZON)
        if price_jpy is None:
            continue  # 価格が取得できない商品はスキップ

        # 現在の販売ランク（csv[3] の最新値）
        rank_raw = _latest_price_from_csv(csv, 3)
        sales_rank = int(rank_raw) if rank_raw and rank_raw > 0 else None

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


def _latest_price_from_csv(csv: list, price_type: int) -> float | None:
    """
    csv配列の指定価格タイプから最新価格を取得する。
    keepa csv形式: [time0, price0, time1, price1, ...] のフラット配列。
    最新価格は末尾の要素（奇数インデックス）。
    """
    if not csv or price_type >= len(csv):
        return None
    history = csv[price_type]
    if not history or len(history) < 2:
        return None
    # 末尾が価格（奇数インデックス）、その前がタイムスタンプ
    raw = history[-1]
    if raw is None or raw < 0:
        return None
    return raw / _KEEPA_PRICE_DIVISOR
