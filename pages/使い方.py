"""使い方ガイドページ"""

import streamlit as st

st.set_page_config(page_title="Amapro - 使い方", page_icon="🔶", layout="wide")

# Amazon風テーマ（app.pyと共通）
st.markdown("""
<style>
    header[data-testid="stHeader"] {
        background-color: #131921 !important;
    }
    header [data-testid="stToolbar"] {
        display: none !important;
    }
    section[data-testid="stSidebar"] {
        background-color: #232F3E !important;
    }
    section[data-testid="stSidebar"] * {
        color: #FFFFFF !important;
    }
    section[data-testid="stSidebar"] hr {
        border-color: #3B4859 !important;
    }
    [data-testid="stSidebarNav"] {
        display: none !important;
    }
    h1, h2, h3 { color: #131921 !important; }
    a { color: #007185 !important; }
</style>
""", unsafe_allow_html=True)

# サイドバーにナビ
with st.sidebar:
    st.markdown("""
    <div style="background:#37475A; margin:-1rem -1rem 1rem -1rem; padding:12px 16px;">
        <a href="/" target="_self"
           style="color:#DDDDDD; text-decoration:none; font-size:14px;
                  display:block; padding:6px 0 6px 15px;">
            🏠 HOME
        </a>
        <a href="/使い方" target="_self"
           style="color:#FFFFFF; text-decoration:none; font-weight:bold; font-size:15px;
                  display:block; padding:6px 0; border-left:3px solid #FF9900; padding-left:12px;">
            📖 使い方ガイド
        </a>
    </div>
    """, unsafe_allow_html=True)

st.markdown("""
<div style="text-align:center; padding: 10px 0 20px 0;">
    <span style="font-size:32px; font-weight:bold; color:#131921; letter-spacing:-1px;
                 font-family:'Amazon Ember','Helvetica Neue',Arial,sans-serif;">
        ama<span style="color:#FF9900;">pro</span>
    </span>
    <span style="color:#565959; font-size:16px; margin-left:12px;">使い方ガイド</span>
</div>
""", unsafe_allow_html=True)

st.header("このツールでできること")
st.markdown("""
卸問屋から入手した **JANコード入りExcel** をアップロードするだけで、
Sheet2（Keepaエクスポート）のデータとJANコードを自動マッチングし、
**Sheet3に利益計算付きのリストを自動生成**します。
""")

st.divider()

st.header("使い方（3ステップ）")
st.markdown("""
### ① Excelファイルをアップロード
卸問屋から受け取ったExcelをドラッグ＆ドロップします。

| シート | 内容 | 必須 |
|--------|------|------|
| シート1（商品データ） | JANコード・品名・卸価格 | ✅ |
| シート2（Keepaエクスポート） | Keepa Webからエクスポートしたデータ | ✅ |
| シート3 | 利益計算シート（自動で上書きされます） | - |

### ② 「Sheet3 を自動生成」を押す
Sheet1のJANコードとSheet2のEANを自動マッチングし、Sheet3に数式ごと生成します。
APIキーもトークンも不要です。

### ③ 結果Excelを取得
- **ローカル実行時**: デスクトップに自動保存され、Excelが開きます
- **Streamlit Cloud**: 「手動ダウンロード」ボタンから取得
""")

st.divider()

st.header("Sheet3に生成される内容")
st.markdown("""
| 列 | 項目 | データ元 |
|----|------|---------|
| D | 商品名 | Sheet2 → 商品名日本語 |
| E | タイトル | Sheet2 → US Amazon商品名 |
| G | ASIN | Sheet2 → ASIN |
| N | 卸価格 | Sheet2 → 卸 (AM列) |
| O | 30キーパ | Sheet2 → 30日間ランク下落回数 |
| P | セラー数（FBA） | Sheet2 → FBAオファー数 |
| T | 売価USD | Sheet2 → Buy Box価格 |
| V | Amazon手数料 | Sheet2 → FBA手数料 |
| X | Weight | Sheet2 → パッケージ重さ |
""")

st.markdown("""
**自動計算される列（元Sheet3と同一の数式）：**

| 列 | 数式 | 内容 |
|----|------|------|
| W | `=K×N×1.1` | 仕入値（税込） |
| U | `=(W+(X×3))/為替レート` | 仕入＋送料（FBA） |
| Y | `=T-U-V` | 損益（USD） |
| Z | `=(O/(P+1))×Y` | 利益額 |
| AA | `=Y/T` | 利益率 |
| AB | `=Y×S` | 利益額2（初回仕入れ分） |
| AG | `=AB×為替レート` | 利益/円 |
""")

st.divider()

st.header("固定値・定数")
st.markdown("""
| セル | 値 | 説明 |
|------|-----|------|
| K列 | 1 | セット内容（デフォルト） |
| S列 | 5 | 初回仕入れ数（デフォルト） |
| AH2 | 150 | 為替レート（JPY/USD） |
| AI2 | 3 | 国際送料単価（円/g） |

これらはExcel上で変更すると全行が自動で再計算されます。
""")

st.divider()

st.header("注意事項")
st.info("""
**マッチングについて**

- Sheet1のJANコードがSheet2のEAN列に存在する商品のみSheet3に出力されます
- 同じEANに複数の商品がある場合、Buy Box価格が最も高い商品が選ばれます
- Sheet2にないJANコードの商品はSheet3に含まれません
""")

st.warning("""
**利益計算の精度について**

- 売価USD（Buy Box価格）はKeepaエクスポート時点の値です
- 卸価格はSheet2のAM列（Keepaに結合された卸データ）から取得します
- 最終判断は実際のAmazon USページで確認してください
""")
