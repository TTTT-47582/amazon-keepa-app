"""使い方ガイドページ"""

import os
import streamlit as st

st.set_page_config(page_title="Amapro - 使い方", page_icon="🔶", layout="wide")

# Amazon風テーマ
st.markdown("""
<style>
    header[data-testid="stHeader"] { background-color: #131921 !important; }
    header [data-testid="stToolbar"] { display: none !important; }
    section[data-testid="stSidebar"] { background-color: #232F3E !important; }
    section[data-testid="stSidebar"] * { color: #FFFFFF !important; }
    section[data-testid="stSidebar"] hr { border-color: #3B4859 !important; }
    [data-testid="stSidebarNav"] { display: none !important; }
    h1, h2, h3 { color: #131921 !important; }
    a { color: #007185 !important; }
</style>
""", unsafe_allow_html=True)

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
    <span style="font-size:32px; font-weight:bold; color:#131921;
                 font-family:'Amazon Ember','Helvetica Neue',Arial,sans-serif;">
        ama<span style="color:#FF9900;">pro</span>
    </span>
    <span style="color:#565959; font-size:16px; margin-left:12px;">使い方ガイド</span>
</div>
""", unsafe_allow_html=True)

st.header("このツールでできること")
st.markdown("""
卸問屋の **JANコード入りExcel** をアップロードするだけで、
US Amazonの商品データを取得し、**利益計算付きのSheet3を自動生成**します。

データの取得方法は **2つのモード** から選べます。
""")

st.divider()

st.header("2つのデータ取得モード")
c1, c2 = st.columns(2)

with c1:
    st.markdown("""
    <div style="background:#F7F8FA; border:1px solid #D5D9D9; border-radius:8px;
                padding:16px; min-height:280px;">
        <h4 style="color:#FF9900;">📄 Excel内Keepaデータ（Sheet2）</h4>
        <p><strong>おすすめ・APIキー不要</strong></p>
        <ul>
            <li>トークン消費ゼロ</li>
            <li>Sheet2のKeepaエクスポートデータと自動マッチング</li>
            <li>数秒〜数十秒で完了</li>
            <li>元Sheet3と同一の数式で出力</li>
        </ul>
        <p><strong>必要なExcel構成：</strong></p>
        <ul>
            <li>シート1: JANコード・品名・卸価格</li>
            <li>シート2: Keepa Webからのエクスポートデータ</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

with c2:
    st.markdown("""
    <div style="background:#F7F8FA; border:1px solid #D5D9D9; border-radius:8px;
                padding:16px; min-height:280px;">
        <h4 style="color:#FF9900;">🌐 Keepa APIで取得</h4>
        <p><strong>Sheet2なしでもOK</strong></p>
        <ul>
            <li>Keepa APIで自動取得（APIキー必要）</li>
            <li>JANコードだけのExcelで動作</li>
            <li>1件あたり約1トークン消費</li>
            <li>Proプラン推奨（50トークン/分）</li>
        </ul>
        <p><strong>必要なExcel構成：</strong></p>
        <ul>
            <li>シート1: JANコード・品名・卸価格</li>
            <li>（シート2は不要）</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

st.divider()

st.header("使い方（3ステップ）")
st.markdown("""
### ① サイドバーでモードを選択
- **📄 Excel内Keepaデータ**: Sheet2入りExcelがある場合
- **🌐 Keepa APIで取得**: JANコードだけのExcelの場合（APIキーを入力）

### ② Excelファイルをアップロード
卸問屋のExcelをドラッグ＆ドロップします。
JANコード列は自動検出されるので、列の位置が異なるExcelでもOKです。

### ③ 「Sheet3 を自動生成」or「Keepa API で取得＆生成」を押す
処理完了後：
- **ローカル**: デスクトップに自動保存＆Excelが開きます
- **Streamlit Cloud**: 「手動ダウンロード」ボタンから取得
""")

st.divider()

st.header("出力される列と計算式")
st.markdown("""
**データ列（Keepaから取得）：**

| 列 | 項目 | 内容 |
|----|------|------|
| D | 商品名 | 日本語の商品名 |
| E | タイトル | US Amazonの英語タイトル |
| G | ASIN | US AmazonのASIN |
| N | 卸価格 | 仕入れ単価 |
| O | 30キーパ | 30日間のランク下落回数（≒販売回数） |
| P | セラー数（FBA） | FBA出品者数 |
| T | 売価USD | Buy Box価格（米ドル） |
| V | Amazon手数料 | FBA手数料 + 紹介料 |
| X | Weight | パッケージ重量（g） |

**計算列（元Sheet3と同一の数式）：**

| 列 | 数式 | 内容 |
|----|------|------|
| W | `=K×N×1.1` | 仕入値（税込） |
| U | `=(W+(X×3))/為替レート` | 仕入＋送料（FBA） |
| Y | `=T-U-V` | 損益（USD） |
| Z | `=(O/(P+1))×Y` | 利益額 |
| AA | `=Y/T` | 利益率 |
| AB | `=Y×S` | 利益額2（初回仕入れ分） |
| AG | `=AB×為替レート` | 利益/円 |

**固定値（Excel上で変更可能）：**

| セル | デフォルト | 説明 |
|------|-----------|------|
| K列 | 1 | セット内容 |
| S列 | 5 | 初回仕入れ数 |
| AH2 | 150 | 為替レート（JPY/USD） |
| AI2 | 3 | 国際送料単価（円/g） |
""")

st.divider()

st.header("注意事項")
st.warning("""
**APIモード利用時のトークンについて**

- Keepa Proプランで約50トークン/分の補充速度です
- 処理開始前にトークン残高を自動チェックします
- 不足時はエラーメッセージが表示されます
- 残高確認: https://keepa.com/#!api
""")

st.info("""
**マッチングについて**

- 同じEANに複数の商品がある場合、Buy Box価格が最も高い商品が選ばれます
- JANコード列はヘッダー名（「JAN」を含む列）から自動検出されます
- 列の位置が異なるExcelでも自動で対応します
""")
