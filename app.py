"""
Amazon 商品リサーチアプリ（日本 → 米国 FBA）

Keepa API で日本 Amazon から商品を取得し、
米国 FBA での推定利益を計算してリスト表示する。

起動コマンド: streamlit run app.py
"""

import os
import requests
import streamlit as st
from src.keepa_client import search_products
from src.profit_calc import calculate_profit
from src.config_manager import load_filters


@st.cache_data(ttl=3600)
def _fetch_exchange_rate() -> tuple[float, str]:
    """Frankfurter API から USD→JPY レートを取得する（1時間キャッシュ）"""
    try:
        resp = requests.get(
            "https://api.frankfurter.app/latest?from=USD&to=JPY",
            timeout=5,
        )
        resp.raise_for_status()
        data = resp.json()
        rate = float(data["rates"]["JPY"])
        date = data.get("date", "")
        return rate, date
    except Exception:
        return 150.0, ""

# ページ設定（必ず最初に呼び出す）
st.set_page_config(
    page_title="Amazon 商品リサーチ",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded",
)

# 不動産サイト風のリスト表示に近づけるカスタム CSS
st.markdown(
    """
    <style>
        /* リストヘッダー行 */
        .list-header {
            display: flex;
            background: #f5f5f5;
            padding: 6px 8px;
            border-radius: 4px;
            font-size: 12px;
            color: #888;
            font-weight: bold;
            margin-bottom: 4px;
        }
        /* 商品行の区切り線を細くする */
        hr { margin: 4px 0 !important; border-color: #f0f0f0 !important; }
        /* サイドバー幅を固定 */
        section[data-testid="stSidebar"] { min-width: 270px; max-width: 270px; }
    </style>
    """,
    unsafe_allow_html=True,
)


def main():
    st.title("🛒 Amazon 商品リサーチ（日本 → 米国 FBA）")
    st.caption(
        "Keepa データをもとに、日本 Amazon で仕入れて米国 Amazon FBA で販売できる商品を自動検索します。"
    )

    # 設定ファイルのデフォルト値を読み込む
    try:
        config = load_filters()
    except FileNotFoundError:
        st.error("設定ファイル config/screening_filters.yaml が見つかりません。")
        return

    df = config.get("filters", {})
    dp = config.get("profit", {})

    # ================================================
    # サイドバー：スクリーニング条件（後から追加可能）
    # ================================================
    with st.sidebar:
        st.header("🔍 スクリーニング条件")
        st.caption("条件は config/screening_filters.yaml でも変更できます")

        # .env に KEEPA_API_KEY があれば自動入力、なければ手入力できる
        env_key = os.getenv("KEEPA_API_KEY", "")
        api_key_input = st.text_input(
            "Keepa API キー",
            value=env_key,
            type="password",
            placeholder=".env 未設定の場合はここに入力",
            help="Keepa API キー。.env の KEEPA_API_KEY が優先されます。",
        )
        # UI 入力 → .env の優先順位でキーを決定
        api_key = api_key_input or env_key

        st.divider()

        check_us_listing = st.checkbox(
            "🇺🇸 米国 Amazon に出品されている商品のみ表示",
            value=False,
            help="EAN/UPC で US Amazon の出品有無をチェックします。ONにするとAPIトークンを多く消費します。",
        )

        with st.expander("📊 販売条件", expanded=True):
            sales_rank_max = st.number_input(
                "販売ランク 上限",
                min_value=1,
                max_value=1_000_000,
                value=df.get("sales_rank_max", 200000),
                step=5000,
                help="数値が小さいほど売れている商品です（例: 200,000）",
            )
            col1, col2 = st.columns(2)
            price_min = col1.number_input(
                "仕入れ下限(円)",
                min_value=0,
                value=df.get("price_min", 1000),
                step=100,
            )
            price_max = col2.number_input(
                "仕入れ上限(円)",
                min_value=0,
                value=df.get("price_max", 20000),
                step=100,
            )
            max_weight_g = st.number_input(
                "商品重量 上限 (g)",
                min_value=100,
                max_value=5000,
                value=df.get("max_weight_g", 500),
                step=100,
                help="重い商品は国際送料が高くなり赤字になりやすいです。500g以下推奨。",
            )

        with st.expander("🗂 カテゴリ絞り込み", expanded=False):
            st.caption("チェックなし＝すべてのカテゴリを対象")
            _category_options = [
                "エレクトロニクス",
                "ホーム&キッチン",
                "おもちゃ",
                "スポーツ",
                "ビューティー",
                "ヘルス",
                "ベビー",
                "ペット用品",
                "楽器",
                "文房具",
                "カー&バイク",
                "DIY・工具",
                "ファッション",
                "食品",
                "アウトドア",
            ]
            allowed_categories = []
            col_a, col_b = st.columns(2)
            for i, cat in enumerate(_category_options):
                col = col_a if i % 2 == 0 else col_b
                if col.checkbox(cat, key=f"cat_{cat}"):
                    allowed_categories.append(cat)

        with st.expander("📈 詳細条件（Keepa スクリーニング）", expanded=False):
            st.caption("0 = 制限なし（空白と同じ）")

            st.markdown("**🏪 セラー数（JP 新品出品者数）**")
            col1, col2 = st.columns(2)
            seller_min = col1.number_input(
                "最小", min_value=0, value=0, step=1, key="seller_min",
                help="最低セラー数。1以上にすると在庫なし商品を除外できます",
            )
            seller_max = col2.number_input(
                "最大", min_value=0, value=0, step=1, key="seller_max",
                help="最大セラー数。小さくするほど競合が少ない商品に絞れます",
            )

            st.markdown("**📉 売れ筋ランキング 過去90日間の下落回数**")
            st.caption("回数が多い＝その期間に多く売れた商品")
            col3, col4 = st.columns(2)
            rank_drops_min = col3.number_input(
                "最小", min_value=0, value=0, step=1, key="rank_drops_min",
                help="例: 3以上にすると90日で3回以上売れた商品のみ",
            )
            rank_drops_max = col4.number_input(
                "最大", min_value=0, value=0, step=1, key="rank_drops_max",
            )

            st.markdown("**📦 在庫切れ率 過去90日間 (%)**")
            st.caption("低いほど安定して在庫がある商品")
            col5, col6 = st.columns(2)
            oos_min = col5.number_input(
                "最小 (%)", min_value=0, value=0, step=5, key="oos_min",
            )
            oos_max = col6.number_input(
                "最大 (%)", min_value=0, value=0, step=5, key="oos_max",
                help="例: 20以下にすると90日の20%以下しか在庫切れしていない商品のみ",
            )

        with st.expander("⭐ レビュー条件", expanded=True):
            rating_min = st.slider(
                "最低評価",
                min_value=1.0,
                max_value=5.0,
                value=float(df.get("rating_min", 3.5)),
                step=0.5,
            )
            review_count_min = st.number_input(
                "最低レビュー数",
                min_value=0,
                value=df.get("review_count_min", 10),
                step=5,
            )

        with st.expander("💰 利益計算設定", expanded=True):
            live_rate, rate_date = _fetch_exchange_rate()
            auto_rate = st.checkbox(
                "為替レートを自動取得する",
                value=True,
                help="ONにすると現在のUSD/JPYレートを自動で反映します",
            )
            if auto_rate:
                if rate_date:
                    st.caption(f"💱 現在レート: 1 USD = ¥{live_rate:.2f}（{rate_date} 更新）")
                else:
                    st.caption("⚠️ レート取得失敗。¥150を使用します。")
                exchange_rate = live_rate
            else:
                exchange_rate = st.number_input(
                    "為替レート (1 USD = X 円)",
                    min_value=100.0,
                    max_value=250.0,
                    value=float(dp.get("exchange_rate_jpy_per_usd", 150.0)),
                    step=1.0,
                )
            us_price_markup = st.slider(
                "米国販売価格の倍率",
                min_value=1.5,
                max_value=5.0,
                value=2.5,
                step=0.1,
                help="仕入れ USD × この倍率 = 推定米国販売価格",
            )
            target_margin = st.number_input(
                "目標利益率 (%)",
                min_value=1,
                max_value=100,
                value=dp.get("target_margin_percent", 30),
                step=5,
                help="この利益率以上の商品を緑色でハイライト",
            )
            filter_by_margin = st.checkbox(
                "目標利益率未満を除外する",
                value=True,
                help="ONにすると目標利益率を下回る商品は表示しません",
            )

        max_results = st.number_input(
            "最大表示件数",
            min_value=5,
            max_value=15,
            value=10,
            step=5,
        )

        st.divider()
        search_btn = st.button(
            "🔍 商品を検索する",
            type="primary",
            use_container_width=True,
            disabled=not api_key,  # キー未入力時はボタンを無効化
        )

    profit_params = {
        "exchange_rate": exchange_rate,
        "us_price_markup": us_price_markup,
        "target_margin": target_margin,
        "filter_by_margin": filter_by_margin,
        "shipping_cost_per_kg_jpy": dp.get("shipping_cost_per_kg_jpy", 1500),
        "fba_referral_fee_percent": dp.get("fba_referral_fee_percent", 15),
        "fba_fulfillment_fee_usd": dp.get("fba_fulfillment_fee_base_usd", 4.75),
    }

    # ================================================
    # メインエリア：検索結果
    # ================================================
    if not api_key:
        st.warning("サイドバーに Keepa API キーを入力してください。")
        return

    if search_btn:
        # 検索実行して結果を session_state に保存する（画面遷移後も表示を維持するため）
        search_params = {
            "sales_rank_max": sales_rank_max,
            "price_min": price_min,
            "price_max": price_max,
            "max_weight_g": int(max_weight_g),
            "rating_min": rating_min,
            "review_count_min": review_count_min,
            "max_results": int(max_results),
            "check_us_listing": check_us_listing,
            "allowed_categories": allowed_categories,
            # 詳細条件（0 = 制限なし）
            "seller_min": int(seller_min),
            "seller_max": int(seller_max),
            "rank_drops_90_min": int(rank_drops_min),
            "rank_drops_90_max": int(rank_drops_max),
            "oos_90_max": int(oos_max),
            "oos_90_min": int(oos_min),
        }
        _run_search(search_params, profit_params, api_key)

    # 検索済みの結果があれば表示する（ボタンを押していない場合も維持）
    elif "search_results" in st.session_state and st.session_state["search_results"]:
        _render_summary_and_list(
            st.session_state["search_results"],
            st.session_state.get("profit_params", profit_params),
        )
    else:
        st.info("👈 左のサイドバーで検索条件を設定して「商品を検索する」を押してください")


def _run_search(search_params: dict, profit_params: dict, api_key: str = ""):
    """商品を検索して利益計算し、結果を session_state に保存して表示する"""
    with st.spinner("Keepa で商品を検索中..."):
        try:
            products = search_products(search_params, api_key=api_key)
        except Exception as e:
            st.error(f"検索エラー: {e}")
            return

    if not products:
        st.warning(
            "条件に一致する商品が見つかりませんでした。"
            "販売ランク・価格帯・レビュー条件を緩めて再検索してください。"
        )
        return

    # 利益計算してソート
    results = [
        {**product, **calculate_profit(product, profit_params)}
        for product in products
    ]
    results.sort(key=lambda x: x.get("profit_margin_percent", 0), reverse=True)

    # 目標利益率フィルタ（チェックONのとき目標未満を除外）
    if profit_params.get("filter_by_margin"):
        target = profit_params["target_margin"]
        results = [r for r in results if r.get("profit_margin_percent", 0) >= target]

    if not results:
        st.warning(
            f"目標利益率 {profit_params['target_margin']}% 以上の商品が見つかりませんでした。"
            "目標利益率を下げるか、「目標利益率未満を除外する」をOFFにして再検索してください。"
        )
        return

    # 結果を session_state に保存（画面を操作しても消えないようにする）
    st.session_state["search_results"] = results
    st.session_state["profit_params"] = profit_params

    _render_summary_and_list(results, profit_params)


def _render_summary_and_list(results: list, profit_params: dict):
    """サマリーと商品リストを表示する"""
    profitable = [r for r in results if r.get("profit_jpy", 0) > 0]
    target_ok = [
        r for r in results if r.get("profit_margin_percent", 0) >= profit_params["target_margin"]
    ]

    st.success(
        f"✅ **{len(results)} 件**見つかりました  "
        f"（利益あり: {len(profitable)} 件 ／ 目標利益率達成: {len(target_ok)} 件）"
    )

    margins = [r.get("profit_margin_percent", 0) for r in results]
    profits = [r.get("profit_jpy", 0) for r in results]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("平均利益率", f"{sum(margins)/len(margins):.1f}%")
    c2.metric("最高利益率", f"{max(margins):.1f}%")
    c3.metric("平均推定利益", f"¥{sum(profits)/len(profits):,.0f}")
    c4.metric("目標達成件数", f"{len(target_ok)} 件")

    st.divider()
    _render_product_list(results, profit_params["target_margin"])


def _render_product_list(results: list[dict], target_margin: float):
    """商品を不動産サイト風のリスト形式で一覧表示する"""

    # カラムヘッダー
    header = st.columns([0.7, 3.5, 1.6, 1.6, 1.8, 2.0])
    labels = ["画像", "商品情報", "仕入れ価格（JP）", "推定販売価格（US）", "諸費用（USD）", "推定利益"]
    for col, label in zip(header, labels):
        col.markdown(f"<small style='color:#999'>{label}</small>", unsafe_allow_html=True)

    st.divider()

    for item in results:
        margin = item.get("profit_margin_percent", 0)
        profit_jpy = item.get("profit_jpy", 0)
        is_target = margin >= target_margin
        is_profitable = profit_jpy > 0

        cols = st.columns([0.7, 3.5, 1.6, 1.6, 1.8, 2.0])

        # ① サムネイル
        with cols[0]:
            if item.get("image_url"):
                st.image(item["image_url"], width=60)
            else:
                st.markdown("📦")

        # ② 商品名・属性（タイトルクリックで日本 Amazon 商品ページへ）
        with cols[1]:
            title = item.get("title", "タイトル不明")
            asin = item.get("asin", "")
            jp_url = f"https://www.amazon.co.jp/dp/{asin}" if asin else "#"
            st.markdown(f"**[{title[:60]}{'…' if len(title) > 60 else ''}]({jp_url})**")
            rank = item.get("sales_rank")
            rank_str = f"{rank:,}" if rank else "不明"
            st.caption(
                f"ASIN: {asin}　｜　{item.get('category', '-')}"
                f"　｜　ランク: {rank_str}"
                f"　｜　⭐ {item.get('rating', 0):.1f}（{item.get('review_count', 0):,} 件）"
            )

        # ③ 仕入れ価格（日本）
        with cols[2]:
            jp_price = item.get("price_jp_jpy", 0)
            st.markdown(f"**¥{jp_price:,.0f}**")
            st.caption(f"≈ ${item.get('purchase_usd', 0):.2f}")

        # ④ 米国販売価格（実際の出品価格 or 推定）
        with cols[3]:
            us_price = item.get("us_sell_price_usd", 0)
            is_actual = bool(item.get("us_actual_price_usd"))
            label = "実績価格" if is_actual else "推定価格"
            st.markdown(f"**${us_price:.2f}** `{label}`")
            st.caption(f"≈ ¥{item.get('us_sell_price_jpy', 0):,}")

        # ⑤ 諸費用の内訳
        with cols[4]:
            st.markdown(f"**${item.get('total_fees_usd', 0):.2f}**")
            st.caption(
                f"送料 ¥{item.get('shipping_cost_jpy', 0):,}　"
                f"紹介料 ${item.get('referral_fee_usd', 0):.2f}　"
                f"FBA ${item.get('fba_fulfillment_usd', 0):.2f}"
            )

        # ⑥ 推定利益（目標達成=緑、利益あり=通常、赤字=赤）
        with cols[5]:
            if is_target:
                st.markdown(f"**:green[¥{profit_jpy:,.0f}]**")
            elif is_profitable:
                st.markdown(f"**¥{profit_jpy:,.0f}**")
            else:
                st.markdown(f"**:red[¥{profit_jpy:,.0f}]**")
            st.caption(
                f"利益率 {margin:.1f}%　｜　ROI {item.get('roi_percent', 0):.1f}%"
            )

        st.divider()


if __name__ == "__main__":
    main()
