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

# csv 配列の価格タイプインデックス（stats['current'] にも同じインデックスを使う）
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


@st.cache_data(ttl=1800, show_spinner=False)
def search_products(params: dict, api_key: str = "") -> list[dict]:
    """
    日本 Amazon で商品を検索し、利益計算に必要な情報を返す。
    history=False で現在価格のみ取得することで高速化。
    同じ条件は 30 分キャッシュしてトークン消費を抑える。
    """
    api = _get_api(_resolve_api_key(api_key))
    max_results = params.get("max_results", 5)

    # ① Product Finder で JP Amazon から ASIN を取得
    product_parms = {
        "current_SALES_gte": 1,
        "current_SALES_lte": params["sales_rank_max"],
        "current_NEW_gte": params["price_min"],
        "current_NEW_lte": params["price_max"],
    }
    asins = api.product_finder(product_parms, domain="JP", wait=True)
    if not asins:
        return []

    # 件数を絞ってデータ転送量を抑える
    asins = list(asins)[:max_results]

    # ② JP 商品詳細を取得（history=False で現在価格のみ・高速）
    jp_raw = api.query(
        asins, domain="JP",
        history=False,      # 全履歴不要 → 大幅に高速化
        update=0,           # Keepa キャッシュを使用（クロールしない）
        wait=True,
    )
    jp_results = _parse_jp_products(jp_raw)

    # ③ EAN で US Amazon を検索（チェックボックス ON のときのみ・上限 5 件）
    ean_to_us_price: dict[str, float] = {}
    if params.get("check_us_listing", True):
        all_eans: list[str] = []
        for p in jp_results:
            all_eans.extend(p.get("ean_list") or [])
        all_eans = list(dict.fromkeys(all_eans))[:5]

        if all_eans:
            us_raw = api.query(
                all_eans, domain="US",
                history=False,
                update=0,
                wait=True,
            )
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
        matched = jp_results

    # ⑤ 評価・レビューフィルタ（0 = データなしはスキップ）
    filtered = [
        p for p in matched
        if (p["rating"] == 0 or p["rating"] >= params["rating_min"])
        and (p["review_count"] == 0 or p["review_count"] >= params["review_count_min"])
    ]

    return filtered[:max_results]


def _parse_jp_products(products_raw: list) -> list[dict]:
    """Keepa の JP 生レスポンスを扱いやすい辞書形式に変換する"""
    results = []

    for p in products_raw or []:
        if not p:
            continue

        # 価格取得: stats['current'] → csv の順で試みる
        price_jpy = _get_jp_price(p)
        if price_jpy is None:
            continue  # 価格不明はスキップ

        # 販売ランク
        sales_rank = _get_sales_rank(p)

        # サムネイル URL
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


def _get_jp_price(p: dict) -> float | None:
    """
    JP 商品から現在の新品価格（JPY）を取得する。
    stats['current'] → stats['avg'] → csv の順で試みる。
    """
    stats = p.get("stats") or {}

    # stats['current'] から取得（history=False でも使える・高速）
    current = stats.get("current") or []
    for idx in [_IDX_NEW, _IDX_NEW_FBA, _IDX_AMAZON]:
        val = current[idx] if idx < len(current) else None
        if val and val > 0:
            return val / _DIVISOR_JP

    # stats['avg'] から取得（90日平均）
    avg = stats.get("avg") or []
    for idx in [_IDX_NEW, _IDX_NEW_FBA, _IDX_AMAZON]:
        val = avg[idx] if idx < len(avg) else None
        if val and val > 0:
            return val / _DIVISOR_JP

    # csv から取得（history=True の場合のフォールバック）
    csv = p.get("csv") or []
    for idx in [_IDX_NEW, _IDX_NEW_FBA, _IDX_AMAZON]:
        price = _latest_price(csv, idx, _DIVISOR_JP)
        if price is not None:
            return price

    return None


def _get_sales_rank(p: dict) -> int | None:
    """JP 商品から現在の販売ランクを取得する"""
    stats = p.get("stats") or {}
    current = stats.get("current") or []

    rank_val = current[_IDX_SALES] if len(current) > _IDX_SALES else None
    if rank_val and rank_val > 0:
        return int(rank_val)

    # csv フォールバック
    csv = p.get("csv") or []
    rank_raw = _latest_price(csv, _IDX_SALES, 1)
    return int(rank_raw) if rank_raw and rank_raw > 0 else None


def _get_us_price(us_product: dict) -> float | None:
    """US 商品からバイボックス → FBA → 新品 の順で価格（USD）を取得する"""
    stats = us_product.get("stats") or {}

    # stats['current'] から取得
    current = stats.get("current") or []
    for idx in [_IDX_BUYBOX, _IDX_NEW_FBA, _IDX_NEW, _IDX_AMAZON]:
        val = current[idx] if idx < len(current) else None
        if val and val > 0:
            return val / _DIVISOR_US

    # stats['avg'] から取得
    avg = stats.get("avg") or []
    for idx in [_IDX_BUYBOX, _IDX_NEW_FBA, _IDX_NEW, _IDX_AMAZON]:
        val = avg[idx] if idx < len(avg) else None
        if val and val > 0:
            return val / _DIVISOR_US

    # csv フォールバック
    csv = us_product.get("csv") or []
    for idx in [_IDX_BUYBOX, _IDX_NEW_FBA, _IDX_NEW, _IDX_AMAZON]:
        price = _latest_price(csv, idx, _DIVISOR_US)
        if price is not None:
            return price

    return None


def _latest_price(csv: list, price_type: int, divisor: int) -> float | None:
    """
    csv 配列の指定タイプから最新価格を取得する。
    keepa csv 形式: [time0, price0, time1, price1, ...] のフラット配列。
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
