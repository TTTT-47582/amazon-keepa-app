"""
利益計算モジュール

計算式:
  推定米国販売価格 = 仕入れ(JPY) ÷ 為替レート × 販売倍率
  国際送料        = 重量(kg) × 送料単価(JPY/kg) ÷ 為替レート
  Amazon 紹介料   = 米国販売価格 × 紹介料率(%)
  FBA 手数料      = 重量・サイズベースの概算値
  純利益(USD)     = 販売価格 - 仕入れ - 国際送料 - 紹介料 - FBA 手数料
  利益率(%)       = 純利益 ÷ 販売価格 × 100
  ROI(%)          = 純利益 ÷ 仕入れ × 100
"""


def calculate_profit(product: dict, params: dict) -> dict:
    """商品データと計算パラメータから推定利益を計算して返す"""
    exchange_rate: float = params["exchange_rate"]
    price_jpy: float = product.get("price_jp_jpy") or 0

    if price_jpy <= 0 or exchange_rate <= 0:
        return _empty_profit()

    # 仕入れ価格（JPY → USD 換算）
    purchase_usd = price_jpy / exchange_rate

    # 米国販売価格：実際の US 出品価格があればそちらを優先、なければ倍率で推定
    if product.get("us_actual_price_usd"):
        us_sell_price_usd = product["us_actual_price_usd"]
    else:
        us_sell_price_usd = purchase_usd * params["us_price_markup"]

    # 国際送料（重量不明の場合は 500g と仮定）
    weight_kg = ((product.get("weight_g") or 500)) / 1000
    shipping_jpy = weight_kg * params["shipping_cost_per_kg_jpy"]
    shipping_usd = shipping_jpy / exchange_rate

    # Amazon US 紹介料（販売価格 × 紹介料率）
    referral_fee_usd = us_sell_price_usd * (params["fba_referral_fee_percent"] / 100)

    # FBA フルフィルメント手数料（重量から概算）
    fba_fulfillment_usd = _estimate_fba_fee(weight_kg)

    # 合計費用（USD）
    total_fees_usd = shipping_usd + referral_fee_usd + fba_fulfillment_usd

    # 純利益
    profit_usd = us_sell_price_usd - purchase_usd - total_fees_usd
    profit_jpy = profit_usd * exchange_rate

    # 利益率・ROI
    profit_margin = (profit_usd / us_sell_price_usd * 100) if us_sell_price_usd > 0 else 0
    roi = (profit_usd / purchase_usd * 100) if purchase_usd > 0 else 0

    return {
        "purchase_usd": round(purchase_usd, 2),
        "us_sell_price_usd": round(us_sell_price_usd, 2),
        "us_sell_price_jpy": round(us_sell_price_usd * exchange_rate),
        "shipping_cost_jpy": round(shipping_jpy),
        "shipping_cost_usd": round(shipping_usd, 2),
        "referral_fee_usd": round(referral_fee_usd, 2),
        "fba_fulfillment_usd": round(fba_fulfillment_usd, 2),
        "total_fees_usd": round(total_fees_usd, 2),
        "profit_usd": round(profit_usd, 2),
        "profit_jpy": round(profit_jpy),
        "profit_margin_percent": round(profit_margin, 1),
        "roi_percent": round(roi, 1),
    }


def _estimate_fba_fee(weight_kg: float) -> float:
    """
    重量（kg）から FBA フルフィルメント手数料を概算する。
    2024年 Amazon US FBA 料金表をベースにした近似値。
    重量をポンドに変換してサイズ区分を判定する。
    """
    weight_lb = weight_kg * 2.20462

    if weight_lb <= 0.5:
        return 3.22     # Small standard（小型標準）
    elif weight_lb <= 1.0:
        return 3.86     # Standard
    elif weight_lb <= 2.0:
        return 4.75     # Standard（重め）
    elif weight_lb <= 3.0:
        return 5.50     # Large standard（大型標準・軽め）
    elif weight_lb <= 20.0:
        # 大型標準：3lb 超過分は 1lb ごとに加算
        return 5.50 + (weight_lb - 3.0) * 0.16
    else:
        # 特大・重量物
        return 9.73 + weight_lb * 0.42


def _empty_profit() -> dict:
    """価格が取得できない商品に対して返すゼロ値"""
    return {
        "purchase_usd": 0,
        "us_sell_price_usd": 0,
        "us_sell_price_jpy": 0,
        "shipping_cost_jpy": 0,
        "shipping_cost_usd": 0,
        "referral_fee_usd": 0,
        "fba_fulfillment_usd": 0,
        "total_fees_usd": 0,
        "profit_usd": 0,
        "profit_jpy": 0,
        "profit_margin_percent": 0,
        "roi_percent": 0,
    }
