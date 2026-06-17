"""使い方ガイドページ"""

import streamlit as st

st.set_page_config(page_title="使い方", page_icon="📖", layout="wide")

st.title("📖 使い方ガイド")

st.header("このツールでできること")
st.markdown("""
卸問屋から入手した **JANコード入りExcel** をアップロードするだけで、
US Amazonでの販売価格・セラー数・利益率などを自動取得し、
**利益計算済みのExcelファイル** を出力します。
""")

st.divider()

st.header("使い方（3ステップ）")
st.markdown("""
### ① Excelファイルをアップロード
卸問屋から受け取ったJANコード入りExcelをドラッグ＆ドロップします。

| シート | 内容 |
|--------|------|
| シート1（商品データ） | JANコード・品名・卸価格 |
| シート2（Keepaエクスポート） | Keepa Webからエクスポートしたデータ（あれば） |

### ② 設定を確認して「マッチング開始」を押す
サイドバーで以下を確認します：
- **データ取得モード**: Excel内Keepaデータ使用（推奨） or Keepa APIで取得
- **為替レート**: 自動取得（1時間ごと更新）or 手動入力

### ③ 結果Excelを取得
- **ローカル実行時**: デスクトップに自動保存され、Excelが開きます
- **Streamlit Cloud**: 「手動ダウンロード」ボタンから取得
""")

st.divider()

st.header("2つのデータ取得モード")
c1, c2 = st.columns(2)

with c1:
    st.subheader("📄 Excel内Keepaデータ使用")
    st.success("推奨・APIキー不要")
    st.markdown("""
    - **トークン消費ゼロ**
    - シート2のKeepaエクスポートデータを自動マッチング
    - 数秒〜数十秒で完了
    - 日常のリサーチ作業に最適
    """)

with c2:
    st.subheader("🌐 Keepa APIで取得")
    st.warning("APIキー＋トークン必要")
    st.markdown("""
    - 1件あたり1〜3トークン消費
    - **最新**のUS Amazonデータを取得
    - Keepaエクスポートがない新規JANに対応
    - Proプラン推奨（50トークン/分補充）
    """)

st.divider()

st.header("出力Excelの主な項目")
st.markdown("""
| 項目 | 説明 |
|------|------|
| 商品名 | 卸データの品名（日本語） |
| タイトル | US Amazonの商品名（英語） |
| ASIN | US AmazonのASIN |
| 売価USD | Buy Box価格（米ドル） |
| 仕入値 | 卸価格（円） |
| BBセラー | Buy Boxを獲得しているセラー |
| セラー数（FBA / FBM） | FBA出品者数 / 無在庫出品者数 |
| 30キーパ | 30日間のランキング下落回数（≒販売回数） |
| Weight（FBA） | パッケージ重量（g） |
| Amazon手数料 | 紹介料 + FBA Pick&Pack料金 |
| 損益 | 売価 − 仕入 − 手数料（USD） |
| 利益率 | 損益 ÷ 売価 |
""")

st.divider()

st.header("注意事項")
st.warning("""
**APIモード利用時のトークンについて**

- Keepa Proプランで約50トークン/分の補充速度です
- 大量処理（1000件超）はトークン残高を確認してから実行してください
- 残高確認: https://keepa.com/#!api
- トークンが0になると補充まで数時間待つ必要があります
""")

st.info("""
**利益計算の精度について**

- 売価USD（Buy Box価格）はKeepaのデータ取得時点の値です
- FBA手数料はKeepaデータがない場合は重量から概算します
- 紹介料はKeepaデータがない場合は売価の15%で概算します
- 最終判断は実際のAmazon USページで確認してください
""")
