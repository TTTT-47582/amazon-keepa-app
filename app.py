"""
Amazon 商品リサーチアプリ（日本 → 米国 FBA）

Keepa API で日本 Amazon から商品を取得し、
米国 FBA での推定利益を計算してリスト表示する。

起動コマンド: streamlit run app.py
"""

import streamlit as st
from src.keepa_client import search_products
from src.profit_calc import calculate_profit
from src.config_manager import load_filters

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

        with st.expander("📊 販売条件", expanded=True):
            sales_rank_max = st.number_input(
                "販売ランク 上限",
                min_value=1,
                max_value=1_000_000,
                value=df.get("sales_rank_max", 50000),
                step=5000,
                help="数値が小さいほど売れている商品です（例: 50,000）",
            )
            col1, col2 = st.columns(2)
            price_min = col1.number_input(
                "仕入れ下限(円)",
                min_value=0,
                value=df.get("price_min", 500),
                step=100,
            )
            price_max = col2.number_input(
                "仕入れ上限(円)",
                min_value=0,
                value=df.get("price_max", 5000),
                step=100,
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

        max_results = st.number_input(
            "最大表示件数",
            min_value=5,
            max_value=100,
            value=20,
            step=5,
        )

        st.divider()
        search_btn = st.button(
            "🔍 商品を検索する", type="primary", use_container_width=True
        )

    # ================================================
    # メインエリア：検索結果
    # ================================================
    if search_btn:
        _run_search(
            search_params={
                "sales_rank_max": sales_rank_max,
                "price_min": price_min,
                "price_max": price_max,
                "rating_min": rating_min,
                "review_count_min": review_count_min,
                "max_results": int(max_results),
            },
            profit_params={
                "exchange_rate": exchange_rate,
                "us_price_markup": us_price_markup,
                "target_margin": target_margin,
                "shipping_cost_per_kg_jpy": dp.get("shipping_cost_per_kg_jpy", 1500),
                "fba_referral_fee_percent": dp.get("fba_referral_fee_percent", 15),
                "fba_fulfillment_fee_usd": dp.get("fba_fulfillment_fee_base_usd", 4.75),
            },
        )
    else:
        _show_usage_guide()


def _run_search(search_params: dict, profit_params: dict):
    """商品を検索して利益計算し、結果を表示する"""
    with st.spinner("Keepa で商品を検索中... （初回は数秒かかります）"):
        try:
            products = search_products(search_params)
        except ValueError as e:
            # API キー未設定など設定エラー
            st.error(str(e))
            return
        except Exception as e:
            st.error(f"検索中にエラーが発生しました: {e}")
            return

    if not products:
        st.warning(
            "条件に一致する商品が見つかりませんでした。"
            "販売ランク・価格帯・レビュー条件を緩めて再検索してください。"
        )
        return

    # 各商品の利益を計算してマージ
    results = [
        {**product, **calculate_profit(product, profit_params)}
        for product in products
    ]

    # 利益率の高い順にソート
    results.sort(key=lambda x: x.get("profit_margin_percent", 0), reverse=True)

    # サマリー集計
    profitable = [r for r in results if r.get("profit_jpy", 0) > 0]
    target_ok = [
        r for r in results if r.get("profit_margin_percent", 0) >= profit_params["target_margin"]
    ]

    st.success(
        f"✅ **{len(results)} 件**見つかりました  "
        f"（利益あり: {len(profitable)} 件 ／ 目標利益率達成: {len(target_ok)} 件）"
    )

    # サマリーメトリクス
    margins = [r.get("profit_margin_percent", 0) for r in results]
    profits = [r.get("profit_jpy", 0) for r in results]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("平均利益率", f"{sum(margins)/len(margins):.1f}%")
    c2.metric("最高利益率", f"{max(margins):.1f}%")
    c3.metric("平均推定利益", f"¥{sum(profits)/len(profits):,.0f}")
    c4.metric("目標達成件数", f"{len(target_ok)} 件")

    st.divider()

    # 商品リスト表示
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

        # ② 商品名・属性
        with cols[1]:
            title = item.get("title", "タイトル不明")
            st.markdown(f"**{title[:60]}{'…' if len(title) > 60 else ''}**")
            rank = item.get("sales_rank")
            rank_str = f"{rank:,}" if rank else "不明"
            st.caption(
                f"ASIN: {item.get('asin', '-')}　｜　{item.get('category', '-')}"
                f"　｜　ランク: {rank_str}"
                f"　｜　⭐ {item.get('rating', 0):.1f}（{item.get('review_count', 0):,} 件）"
            )

        # ③ 仕入れ価格（日本）
        with cols[2]:
            jp_price = item.get("price_jp_jpy", 0)
            st.markdown(f"**¥{jp_price:,.0f}**")
            st.caption(f"≈ ${item.get('purchase_usd', 0):.2f}")

        # ④ 推定米国販売価格
        with cols[3]:
            us_price = item.get("us_sell_price_usd", 0)
            st.markdown(f"**${us_price:.2f}**")
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


def _show_usage_guide():
    """初期画面に使い方ガイドを表示する"""
    st.info("👈 左のサイドバーで検索条件を設定して「商品を検索する」を押してください")

    with st.expander("📖 使い方・利益計算の仕組み", expanded=True):
        st.markdown(
            """
            #### セットアップ
            1. プロジェクトルートに `.env` ファイルを作成
            2. `KEEPA_API_KEY=your_key_here` を記載して保存
            3. `pip install -r requirements.txt` で依存パッケージをインストール
            4. `streamlit run app.py` でアプリを起動

            #### 検索の流れ
            1. サイドバーで **販売ランク・価格帯・レビュー条件** を設定
            2. 「商品を検索する」ボタンを押すと Keepa から商品を取得
            3. 結果は **利益率の高い順** にリスト表示されます

            #### 利益計算の内訳
            | 項目 | 内容 |
            |------|------|
            | 仕入れ価格 | 日本 Amazon の現在最安値（Keepa より取得） |
            | 推定販売価格 | 仕入れ USD × 販売倍率（サイドバーで調整） |
            | 国際送料 | 重量（kg）× 送料単価（config で変更可） |
            | Amazon 紹介料 | 販売価格 × 15%（概算） |
            | FBA 手数料 | 重量・サイズに応じた 2024 年 US 料金表ベース |
            | **推定利益** | **販売価格 − 仕入れ − 国際送料 − Amazon 各手数料** |

            > スクリーニング条件は `config/screening_filters.yaml` を編集して追加・変更できます
            """
        )


if __name__ == "__main__":
    main()
