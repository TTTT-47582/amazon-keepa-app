"""
JAN コード → US Amazon リサーチツール

卸問屋のJANコードExcelをアップロードすると、
Sheet2のKeepaデータとマッチングし、
Sheet3に数式ごと自動生成する。

起動コマンド: streamlit run app.py
"""

from __future__ import annotations

import os
import tempfile
import subprocess
import platform
from pathlib import Path
import requests
import streamlit as st
from src.excel_io import read_wholesaler_excel, build_ean_to_sheet2_row, generate_sheet3

st.set_page_config(
    page_title="Amapro - 仕入れリサーチ",
    page_icon="🔶",
    layout="wide",
)

# Amazon風カラーテーマ
st.markdown("""
<style>
    /* ヘッダーバー */
    header[data-testid="stHeader"] {
        background-color: #131921 !important;
    }
    /* 不要なヘッダーアイコン（星・ペン・GitHub）を非表示 */
    header [data-testid="stToolbar"] {
        display: none !important;
    }
    /* サイドバーの「app」→「HOME」に変更 */
    [data-testid="stSidebarNav"] li:first-child a span {
        font-size: 0 !important;
    }
    [data-testid="stSidebarNav"] li:first-child a span::after {
        content: "HOME";
        font-size: 14px !important;
        color: #FFFFFF !important;
    }
    /* サイドバー */
    section[data-testid="stSidebar"] {
        background-color: #232F3E !important;
    }
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stNumberInput label,
    section[data-testid="stSidebar"] .stTextInput label {
        color: #FF9900 !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: #3B4859 !important;
    }
    /* メインエリア */
    .main .block-container {
        background-color: #FFFFFF;
    }
    /* タイトル */
    h1 {
        color: #131921 !important;
    }
    /* プライマリボタン → Amazon オレンジ */
    button[kind="primary"], .stButton > button[kind="primary"] {
        background-color: #FF9900 !important;
        border-color: #E88B00 !important;
        color: #131921 !important;
        font-weight: bold !important;
    }
    button[kind="primary"]:hover {
        background-color: #E88B00 !important;
    }
    /* ダウンロードボタン */
    .stDownloadButton > button {
        background-color: #FFD814 !important;
        border-color: #FCD200 !important;
        color: #0F1111 !important;
        font-weight: bold !important;
    }
    .stDownloadButton > button:hover {
        background-color: #F7CA00 !important;
    }
    /* 成功メッセージ */
    .stSuccess {
        background-color: #F0FFF0 !important;
        border-left-color: #FF9900 !important;
    }
    /* メトリクス */
    [data-testid="stMetricValue"] {
        color: #131921 !important;
        font-weight: bold !important;
    }
    [data-testid="stMetricLabel"] {
        color: #565959 !important;
    }
    /* ファイルアップローダー */
    [data-testid="stFileUploader"] {
        border-color: #FF9900 !important;
    }
    /* リンク色 */
    a {
        color: #007185 !important;
    }
    a:hover {
        color: #C7511F !important;
    }
</style>
""", unsafe_allow_html=True)


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
    st.markdown("""
    <div style="text-align:center; padding: 10px 0 40px 0;">
        <div style="display:inline-block; position:relative;">
            <span style="font-size:48px; font-weight:bold; color:#131921; letter-spacing:-1px;
                         font-family:'Amazon Ember','Helvetica Neue',Arial,sans-serif;">
                ama<span style="color:#FF9900;">pro</span>
            </span>
            <svg width="120" height="20" viewBox="0 0 120 20"
                 style="display:block; margin:-8px auto 0 auto;">
                <path d="M10 12 Q60 28 110 8" stroke="#FF9900" stroke-width="3"
                      fill="none" stroke-linecap="round"/>
                <polygon points="107,3 115,7 107,11" fill="#FF9900"/>
            </svg>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── サイドバー ──
    with st.sidebar:
        st.markdown("<h3 style='color:#FF9900;'>⚙️ 設定</h3>", unsafe_allow_html=True)

        live_rate, rate_date = _fetch_exchange_rate()
        if rate_date:
            st.caption(f"💱 現在レート: 1 USD = ¥{live_rate:.2f}（{rate_date}）")

    # ── メインエリア ──
    uploaded = st.file_uploader(
        "卸元の Excel ファイルをアップロード（JAN コード + Keepa エクスポート入り）",
        type=["xlsx"],
    )

    if uploaded is None:
        st.info("👆 Excelファイルをドラッグ＆ドロップまたは選択してください")
        st.markdown("""
        **必要なシート構成：**
        - **シート1**（商品データ）: JANコード・品名・卸価格
        - **シート2**（Keepaエクスポート）: Keepa Webからエクスポートしたデータ
        - **シート3**: 利益計算シート（自動で上書きされます）
        """)
        return

    # Excel読み込み
    if "uploaded_df" not in st.session_state or st.session_state.get("uploaded_name") != uploaded.name:
        with st.spinner("Excel を読み込み中..."):
            df = read_wholesaler_excel(uploaded)
            uploaded.seek(0)
            ean_to_row = build_ean_to_sheet2_row(uploaded)
            st.session_state["uploaded_df"] = df
            st.session_state["ean_to_row"] = ean_to_row
            st.session_state["uploaded_name"] = uploaded.name
            st.session_state["generated"] = False

    df = st.session_state.get("uploaded_df")
    ean_to_row = st.session_state.get("ean_to_row")
    if df is None or ean_to_row is None:
        return

    unique_jans = df["jan_code"].nunique()
    matched = sum(1 for jan in df["jan_code"].unique() if jan in ean_to_row)

    st.success(f"✅ **{len(df):,} 行**読み込み（ユニーク JAN: **{unique_jans:,} 件**）")
    st.info(f"🔍 Sheet2 とマッチ: **{matched:,} 件** / {unique_jans:,} 件")

    with st.expander("📋 データプレビュー（先頭20行）", expanded=False):
        st.dataframe(df.head(20), use_container_width=True)

    # 処理実行
    start_btn = st.button(
        "🚀 Sheet3 を自動生成",
        type="primary",
        use_container_width=True,
    )

    if start_btn:
        _run_generation(uploaded)

    if st.session_state.get("generated"):
        output_path = st.session_state.get("output_path")
        result = st.session_state.get("gen_result", {})
        _show_result(output_path, result)


def _run_generation(uploaded_file):
    """元Excelをコピーし、Sheet3を自動生成する"""
    with st.spinner("Sheet3 を生成中..."):
        # アップロードファイルを一時保存
        uploaded_file.seek(0)
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_in:
            tmp_in.write(uploaded_file.read())
            tmp_in_path = tmp_in.name

        # 出力先（ローカルはDesktop、Cloudは一時ディレクトリ）
        desktop = Path.home() / "Desktop"
        if desktop.exists():
            output_path = desktop / "リサーチ結果.xlsx"
        else:
            output_path = Path(tempfile.gettempdir()) / "リサーチ結果.xlsx"

        # Sheet2のシート名を取得
        import openpyxl
        wb_tmp = openpyxl.load_workbook(tmp_in_path, read_only=True)
        sheet2_name = wb_tmp.sheetnames[1] if len(wb_tmp.sheetnames) > 1 else "Sheet2"
        wb_tmp.close()

        try:
            result = generate_sheet3(
                source_path=tmp_in_path,
                output_path=output_path,
                sheet2_name=sheet2_name,
            )
        except Exception as e:
            st.error(f"生成エラー: {e}")
            import traceback
            st.code(traceback.format_exc())
            return
        finally:
            os.unlink(tmp_in_path)

    st.session_state["generated"] = True
    st.session_state["output_path"] = str(output_path)
    st.session_state["gen_result"] = result

    st.success(
        f"🎉 完了！ **{result['matched']:,} 件**をSheet3に生成 "
        f"（全{result['total_jans']:,} JAN中）"
    )


def _show_result(output_path: str, result: dict):
    """結果表示とファイルオープン"""
    st.divider()

    c1, c2, c3 = st.columns(3)
    c1.metric("全JANコード", f"{result.get('total_jans', 0):,}")
    c2.metric("Sheet2マッチ", f"{result.get('matched', 0):,}")
    c3.metric("未マッチ", f"{result.get('total_jans', 0) - result.get('matched', 0):,}")

    st.success(f"📂 保存先: **{output_path}**")

    # 自動でExcelを開く
    try:
        if platform.system() == "Darwin":
            subprocess.Popen(["open", output_path])
        elif platform.system() == "Windows":
            subprocess.Popen(["start", "", output_path], shell=True)
    except Exception:
        pass

    # ダウンロードボタンも用意
    try:
        with open(output_path, "rb") as f:
            st.download_button(
                "📥 結果 Excel を手動ダウンロード",
                data=f.read(),
                file_name="リサーチ結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    except Exception:
        pass


if __name__ == "__main__":
    main()
