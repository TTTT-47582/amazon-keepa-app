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

_DIVISOR_US = 100
_IDX_BUYBOX = 18


@st.cache_resource
def _get_api(api_key: str) -> keepa.Keepa:
    """Keepa API クライアントを初期化する（同じキーならキャッシュを再利用）"""
    return keepa.Keepa(api_key)


def _resolve_api_key(api_key: str | None) -> str:
    """UI 入力キー → .env の順で有効なキーを返す"""
    key = api_key or os.getenv("KEEPA_API_KEY", "")
    if not key:
        raise ValueError(
            "Keepa API キーが設定されていません。"
            "サイドバーにキーを入力するか、.env に KEEPA_API_KEY を設定してください。"
        )
    return key


def query_jan_codes_us(
    jan_codes: list[str],
    api_key: str = "",
    batch_size: int = 100,
    use_offers: bool = True,
    progress_callback=None,
) -> dict[str, dict]:
    """
    JANコード(EAN)のリストで US Amazon の商品データを一括取得する。

    Returns:
        {jan_code: product_dict} のマッピング。
        US に未出品なら product_dict["found"] = False。
    """
    api = _get_api(_resolve_api_key(api_key))
    results: dict[str, dict] = {}
    total = len(jan_codes)

    for start in range(0, total, batch_size):
        batch = jan_codes[start : start + batch_size]

        try:
            raw = api.query(
                batch,
                domain="US",
                history=False,
                stats=90,
                offers=20 if use_offers else None,
                product_code_is_asin=False,
                wait=True,
            )
        except Exception as e:
            for jan in batch:
                results.setdefault(jan, _empty_result(jan, str(e)))
            if progress_callback:
                progress_callback(min(start + batch_size, total), total)
            continue

        # レスポンスを JAN コードにマッピング
        ean_to_product: dict[str, dict] = {}
        for p in (raw or []):
            if not p:
                continue
            parsed = _parse_us_product(p, use_offers)
            for ean in (p.get("eanList") or []):
                ean_str = str(ean)
                if ean_str not in ean_to_product:
                    ean_to_product[ean_str] = parsed

        for jan in batch:
            if jan in ean_to_product:
                results[jan] = ean_to_product[jan]
            else:
                results.setdefault(jan, _not_found_result())

        if progress_callback:
            progress_callback(min(start + batch_size, total), total)

    return results


def _parse_us_product(p: dict, use_offers: bool) -> dict:
    """Keepa の US 商品レスポンスを必要な項目に変換する"""
    stats = p.get("stats") or {}
    current = stats.get("current") or []

    # Buy Box 価格（USD）
    buy_box_raw = current[_IDX_BUYBOX] if len(current) > _IDX_BUYBOX else None
    buy_box_price = buy_box_raw / _DIVISOR_US if buy_box_raw and buy_box_raw > 0 else None

    # Buy Box セラー
    buy_box_seller = ""
    buy_box_seller_id = stats.get("buyBoxSellerId") or ""
    if buy_box_seller_id:
        buy_box_seller = buy_box_seller_id

    # 30日ランク下落回数
    drops_30 = stats.get("salesRankDrops30")

    # FBA / FBM セラー数（offers から集計）
    fba_count = None
    fbm_count = None
    if use_offers:
        offers = p.get("offers") or []
        fba_count = sum(1 for o in offers if o.get("isFBA"))
        fbm_count = sum(1 for o in offers if not o.get("isFBA"))

    # FBA 手数料
    fba_fees = p.get("fbaFees") or {}
    pick_pack_raw = fba_fees.get("pickAndPackFee")
    fba_pick_pack = pick_pack_raw / _DIVISOR_US if pick_pack_raw and pick_pack_raw > 0 else None

    # 紹介料率
    referral_pct = p.get("referralFeePercentage")
    referral_fee = None
    if referral_pct and buy_box_price:
        referral_fee = round(buy_box_price * referral_pct / 100, 2)

    # 合算手数料（FBA Pick&Pack + 紹介料）= 元Sheet3のV列と同一
    total_fee = None
    if fba_pick_pack is not None:
        ref = referral_fee or 0
        total_fee = round(fba_pick_pack + ref, 2)

    return {
        "found": True,
        "asin": p.get("asin", ""),
        "title": p.get("title") or "",
        "buy_box_price_usd": buy_box_price,
        "buy_box_seller": buy_box_seller,
        "fba_seller_count": fba_count,
        "fbm_seller_count": fbm_count,
        "sales_rank_drops_30": drops_30,
        "package_weight_g": p.get("packageWeight"),
        "total_amazon_fee_usd": total_fee,
    }


def _not_found_result() -> dict:
    return {"found": False, "asin": "", "title": "", "buy_box_price_usd": None,
            "buy_box_seller": "", "fba_seller_count": None, "fbm_seller_count": None,
            "sales_rank_drops_30": None, "package_weight_g": None,
            "total_amazon_fee_usd": None}


def _empty_result(jan: str, error: str = "") -> dict:
    r = _not_found_result()
    r["error"] = error
    return r


def estimate_tokens(n_items: int, use_offers: bool = True) -> int:
    """Keepa API トークン消費量の見積もり"""
    per_item = 3 if use_offers else 1
    return n_items * per_item
