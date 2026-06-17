"""
JAN コード → US Amazon リサーチツール

卸問屋のJANコードExcelをアップロードすると、
Keepa APIでUS Amazonの商品データを取得し、
利益計算済みExcelを出力する。

起動コマンド: streamlit run app.py
"""

from __future__ import annotations

import os
import requests
import streamlit as st
from src.keepa_client import query_jan_codes_us, estimate_tokens
from src.profit_calc import calculate_profit_us
from src.excel_io import read_wholesaler_excel, write_output_excel

st.set_page_config(
    page_title="JAN → US Amazon リサーチ",
    page_icon="📦",
    layout="wide",
)


@st.cache_data(ttl=3600)
def _fetch_exchange_rate() -> tuple[float, str]:
    """Frankfurter API から USD→JPY レートを取得する"""
    try:
        resp = requests.get(
            "https://api.frankfurter.app/latest?from=USD&to=JPY",
            timeout=5,
        )
        resp.raise_for_status()
        data = resp.json()
        return float(data["rates"]["JPY"]), data.get("date", "")
    except Exception:
        return 150.0, ""


def main():
    st.title("📦 JAN → US Amazon リサーチツール")
    st.caption("卸問屋のJANコードExcelをアップロード → Keepa で US Amazon データ取得 → 利益計算Excel出力")

    # ── サイドバー ──
    with st.sidebar:
        st.header("⚙️ 設定")

        env_key = os.getenv("KEEPA_API_KEY", "")
        api_key = st.text_input(
            "Keepa API キー",
            value=env_key,
            type="password",
            placeholder=".env 未設定の場合はここに入力",
        )

        st.divider()

        wholesaler_name = st.text_input("購入先問屋名", value="多摩電子")

        live_rate, rate_date = _fetch_exchange_rate()
        auto_rate = st.checkbox("為替レートを自動取得", value=True)
        if auto_rate:
            exchange_rate = live_rate
            if rate_date:
                st.caption(f"💱 1 USD = ¥{live_rate:.2f}（{rate_date}）")
        else:
            exchange_rate = st.number_input(
                "為替レート (1 USD = X 円)",
                min_value=100.0, max_value=250.0, value=150.0, step=1.0,
            )

        shipping_per_g = st.number_input(
            "国際送料 (円/g)",
            min_value=0.0, max_value=20.0, value=3.0, step=0.5,
            help="EMS: ~4円/g、SAL便: ~2円/g、船便: ~1円/g",
        )

        st.divider()

        use_offers = st.checkbox(
            "セラー数を取得する",
            value=True,
            help="ONにするとFBA/FBMセラー数を取得できますがトークン消費が約3倍になります",
        )

        max_items = st.number_input(
            "処理上限件数",
            min_value=10, max_value=5000, value=100, step=50,
            help="一度に処理するJANコードの上限。多いほど時間がかかります。",
        )

    profit_params = {
        "exchange_rate": exchange_rate,
        "shipping_cost_per_g_jpy": shipping_per_g,
    }

    # ── メインエリア ──
    if not api_key:
        st.warning("サイドバーに Keepa API キーを入力してください。")
        return

    # Phase 1: アップロード
    uploaded = st.file_uploader(
        "卸元の Excel ファイルをアップロード（JAN コード入り）",
        type=["xlsx"],
    )

    if uploaded is None:
        st.info("👆 Excelファイルをドラッグ＆ドロップまたは選択してください")
        return

    # Excel 読み込み（キャッシュ）
    if "uploaded_df" not in st.session_state or st.session_state.get("uploaded_name") != uploaded.name:
        with st.spinner("Excel を読み込み中..."):
            df = read_wholesaler_excel(uploaded)
            st.session_state["uploaded_df"] = df
            st.session_state["uploaded_name"] = uploaded.name
            st.session_state["results"] = None

    df = st.session_state["uploaded_df"]

    # Phase 2: データ概要表示
    unique_jans = df["jan_code"].nunique()
    st.success(f"✅ **{len(df):,} 行**読み込み（ユニーク JAN: **{unique_jans:,} 件**）")

    with st.expander("📋 データプレビュー（先頭20行）", expanded=False):
        st.dataframe(df.head(20), use_container_width=True)

    # 処理対象
    process_count = min(unique_jans, int(max_items))
    tokens = estimate_tokens(process_count, use_offers)
    st.info(
        f"🔍 処理対象: **{process_count:,} 件**（全{unique_jans:,}件中）\n\n"
        f"💰 推定トークン消費: **{tokens:,}** トークン　｜　"
        f"⏱ 推定時間: **約{max(tokens // 50, 1)}分**（50トークン/分の場合）"
    )

    # Phase 3: 処理実行
    start_btn = st.button(
        "🚀 Keepa データ取得 開始",
        type="primary",
        use_container_width=True,
    )

    if start_btn:
        _run_processing(df, process_count, api_key, profit_params, wholesaler_name, use_offers)

    # Phase 4: 結果表示 & ダウンロード
    if st.session_state.get("results"):
        _show_results(st.session_state["results"], wholesaler_name, profit_params)


def _run_processing(
    df, max_count: int, api_key: str, profit_params: dict,
    wholesaler_name: str, use_offers: bool,
):
    """JANコードをKeepaで一括取得し、利益計算して session_state に保存する"""

    # ユニークJANコード抽出（重複排除）
    unique_jans = df["jan_code"].unique().tolist()[:max_count]

    # 卸データを JAN → row のマッピングに変換（先頭行を優先）
    jan_to_wholesale = {}
    for _, row in df.iterrows():
        jan = row["jan_code"]
        if jan not in jan_to_wholesale:
            jan_to_wholesale[jan] = row.to_dict()

    progress_bar = st.progress(0, text="Keepa API に問い合わせ中...")
    status_text = st.empty()

    def on_progress(done, total):
        progress_bar.progress(done / total, text=f"処理中... {done:,}/{total:,} 件")
        status_text.caption(f"完了: {done:,} / {total:,}")

    try:
        keepa_results = query_jan_codes_us(
            unique_jans,
            api_key=api_key,
            batch_size=100,
            use_offers=use_offers,
            progress_callback=on_progress,
        )
    except Exception as e:
        st.error(f"Keepa API エラー: {e}")
        return

    progress_bar.progress(1.0, text="完了！")

    # Keepa結果 + 卸データをマージして利益計算
    results = []
    for jan in unique_jans:
        ws = jan_to_wholesale.get(jan, {})
        kp = keepa_results.get(jan, {})

        merged = {
            "jan_code": jan,
            "product_name_jp": ws.get("product_name_jp", ""),
            "part_number": ws.get("part_number", ""),
            "wholesale_price": ws.get("wholesale_price", 0),
            "box_qty": ws.get("box_qty", 1),
            **kp,
        }

        profit = calculate_profit_us(merged, profit_params)
        merged.update(profit)
        merged["_exchange_rate"] = profit_params["exchange_rate"]
        results.append(merged)

    st.session_state["results"] = results

    found_count = sum(1 for r in results if r.get("found"))
    st.success(
        f"🎉 完了！ **{len(results):,} 件**処理　"
        f"（US出品あり: **{found_count:,} 件** / 未出品: {len(results) - found_count:,} 件）"
    )


def _show_results(results: list[dict], wholesaler_name: str, profit_params: dict):
    """結果サマリーとダウンロードボタンを表示する"""

    st.divider()
    st.subheader("📊 結果サマリー")

    found = [r for r in results if r.get("found")]
    profitable = [r for r in found if (r.get("profit_usd") or 0) > 0]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("処理件数", f"{len(results):,}")
    c2.metric("US出品あり", f"{len(found):,}")
    c3.metric("利益あり", f"{len(profitable):,}")
    if profitable:
        avg_margin = sum(r.get("profit_margin", 0) for r in profitable) / len(profitable)
        c4.metric("平均利益率", f"{avg_margin:.1%}")
    else:
        c4.metric("平均利益率", "-")

    # プレビューテーブル（US出品ありのみ、利益率順）
    if found:
        with st.expander(f"📋 US出品あり商品一覧（{len(found)}件）", expanded=True):
            preview = []
            for r in sorted(found, key=lambda x: x.get("profit_usd") or -9999, reverse=True):
                preview.append({
                    "ASIN": r.get("asin", ""),
                    "タイトル": (r.get("title") or "")[:50],
                    "商品名(JP)": (r.get("product_name_jp") or "")[:20],
                    "売価USD": r.get("buy_box_price_usd"),
                    "仕入値(円)": r.get("wholesale_price"),
                    "損益USD": r.get("profit_usd"),
                    "利益率": f"{r.get('profit_margin', 0):.1%}" if r.get("profit_margin") else "",
                    "FBAセラー": r.get("fba_seller_count"),
                    "30日drop": r.get("sales_rank_drops_30"),
                })
            st.dataframe(preview, use_container_width=True)

    # Excel ダウンロード
    st.divider()
    excel_buf = write_output_excel(results, wholesaler_name)
    st.download_button(
        "📥 結果 Excel をダウンロード",
        data=excel_buf,
        file_name="us_research_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
