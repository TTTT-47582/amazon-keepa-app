"""
利益計算モジュール（卸仕入れ → US Amazon FBA 販売）

計算式:
  仕入USD         = 卸価格(JPY) ÷ 為替レート
  仕入＋送料      = 仕入USD + FBA Pick&Pack
  Amazon手数料    = 紹介料 + FBA Pick&Pack
  損益(USD)       = 売価USD − 仕入＋送料 − 紹介料
  利益率          = 損益 ÷ 売価USD
"""

from __future__ import annotations


def calculate_profit_us(item: dict, params: dict) -> dict:
    """
    卸仕入れ → US FBA 販売の利益計算。
    送料・FBA手数料はKeepaデータから取得。
    """
    exchange_rate = params.get("exchange_rate", 150.0)

    wholesale_jpy = item.get("wholesale_price") or 0
    sell_usd = item.get("buy_box_price_usd")
    weight_g = item.get("package_weight_g") or 500
    amazon_fee = item.get("total_amazon_fee_usd")
    shipping_per_g = 3.0

    if not sell_usd or sell_usd <= 0 or wholesale_jpy <= 0 or exchange_rate <= 0:
        return _zero()

    # Amazon手数料がKeepaデータにない場合は概算
    if amazon_fee is None:
        fba = _estimate_fba_fee_us(weight_g)
        referral = round(sell_usd * 0.15, 2)
        amazon_fee = round(fba + referral, 2)

    # 元Sheet3と同一の計算式: U = (W + X*3) / 為替レート
    wholesale_incl_tax = wholesale_jpy * 1.1
    purchase_plus_shipping = round((wholesale_incl_tax + weight_g * shipping_per_g) / exchange_rate, 2)
    profit_usd = round(sell_usd - purchase_plus_shipping - amazon_fee, 2)
    profit_jpy = round(profit_usd * exchange_rate)
    margin = round(profit_usd / sell_usd, 4) if sell_usd > 0 else 0

    return {
        "purchase_plus_shipping_usd": purchase_plus_shipping,
        "amazon_fees_usd": amazon_fee,
        "profit_usd": profit_usd,
        "profit_jpy": profit_jpy,
        "profit_margin": margin,
    }


def _estimate_fba_fee_us(weight_g: float) -> float:
    """重量から FBA フルフィルメント手数料を概算（USD）"""
    weight_lb = weight_g / 453.59
    if weight_lb <= 0.5:
        return 3.22
    elif weight_lb <= 1.0:
        return 3.86
    elif weight_lb <= 2.0:
        return 4.75
    elif weight_lb <= 3.0:
        return 5.50
    elif weight_lb <= 20.0:
        return 5.50 + (weight_lb - 3.0) * 0.16
    else:
        return 9.73 + weight_lb * 0.42


def _zero() -> dict:
    return {
        "purchase_plus_shipping_usd": None,
        "amazon_fees_usd": None,
        "profit_usd": None,
        "profit_jpy": None,
        "profit_margin": None,
    }
