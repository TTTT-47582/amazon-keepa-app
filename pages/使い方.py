"""使い方ガイドページ"""

import os
import streamlit as st

st.set_page_config(page_title="Amapro - 使い方", page_icon="static/favicon-32.png", layout="wide")

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
    button[data-testid="stSidebarCollapseButton"] { display: none !important; }
</style>
""", unsafe_allow_html=True)

import base64

with st.sidebar:
    with open("static/logo-lockup-dark.png", "rb") as f:
        sb_logo = base64.b64encode(f.read()).decode()
    st.markdown(f"""
    <div style="background:#37475A; margin:-1rem -1rem 1rem -1rem; padding:16px;">
        <div style="text-align:center; margin-bottom:12px;">
            <a href="/" target="_self"><img src="data:image/png;base64,{sb_logo}" height="40" alt="amapro" style="cursor:pointer;"></a>
        </div>
        <a href="/使い方" target="_self"
           style="color:#FFFFFF; text-decoration:none; font-weight:bold; font-size:14px;
                  display:block; padding:6px 0; border-left:3px solid #FF9900; padding-left:12px;">
            📖 使い方ガイド
        </a>
    </div>
    """, unsafe_allow_html=True)

with open("static/logo-lockup.png", "rb") as f:
    main_logo = base64.b64encode(f.read()).decode()
st.markdown(
    f'<div style="text-align:center; padding:10px 0 20px 0;">'
    f'<img src="data:image/png;base64,{main_logo}" height="70" alt="amapro">'
    f'<span style="color:#565959; font-size:16px; margin-left:12px;">使い方ガイド</span>'
    f'</div>',
    unsafe_allow_html=True,
)

# ── 全体の流れ ──
st.header("全体の流れ")
st.markdown("""
```
卸データExcel + KeepaエクスポートExcel
        ↓ アップロード
    Sheet3 自動生成（リサーチ結果Excel）
        ↓
    スクリーニング（30日drop・利益率・Amazon除外）
        ↓
    スクリーニング済みExcel ダウンロード ← 最終成果物
```
""")

st.divider()

# ── ステップ1 ──
st.header("ステップ1: Keepaでデータを準備する")
st.markdown("""
### Keepaで商品データをエクスポート

1. [keepa.com](https://keepa.com) にログイン
2. 上部メニュー **「検索」** → **「製品ビューア」** をクリック
3. **「UPC / EAN / GTINコードのリスト」** を選択
4. 卸データExcelの**JAN列をコピー＆ペースト**（最大10,000件ずつ）
5. **「リストを読み込む」** をクリック
6. テーブル上部の **「列を設定」** で **「過去30日間の減少」にチェック** ← 重要！
7. **「エクスポート」** でExcelをダウンロード

**94,000件の場合：** 10回に分けてエクスポート → Amaproで全ファイルまとめてアップロード可能
""")

st.divider()

# ── ステップ2 ──
st.header("ステップ2: Amaproにアップロード")
st.markdown("""
### ファイルのアップロード方法

| アップロード欄 | 入れるファイル | 必須？ |
|---|---|---|
| **① 卸データExcel** | 問屋のJANコード入りExcel | ✅ |
| **② Keepaエクスポート** | ステップ1でダウンロードしたファイル（複数OK） | ✅ |

- ②は**複数ファイルを同時にアップロード**できます（Ctrl/Cmdを押しながら選択）
- 1つのExcelに3シート入り（卸データ+Keepa+Sheet3）の場合は①だけでOK
""")

st.divider()

# ── ステップ3 ──
st.header("ステップ3: 為替レートを設定してSheet3を生成")
st.markdown("""
1. **為替レート** を確認（デフォルト155円）
2. **「🚀 Sheet3 を自動生成」** ボタンをクリック
3. 処理完了後、リサーチ結果Excelが生成されます

| 環境 | 結果の受け取り方 |
|------|----------------|
| ローカル | デスクトップに自動保存＆Excelが開く |
| Streamlit Cloud | **「📥 結果Excelを手動ダウンロード」** ボタンをクリック |
""")

st.divider()

# ── ステップ4（スクリーニング） ──
st.header("ステップ4: スクリーニングで絞り込む")
st.markdown("""
Sheet3生成後、画面下部に **🔍 スクリーニング** セクションが表示されます。
""")

st.markdown("""
### フィルタ条件

| フィルタ | 意味 | デフォルト |
|---------|------|-----------|
| **30日ランク下落（最低回数）** | 数値が大きいほど売れている | 19以上 |
| **利益率（最低%）** | 売価に対する利益の割合 | 10%以上 |
| **🚫 Amazon本体を除外** | Amazon直販商品を除外 | ON |

スライダーを動かすと**リアルタイム**で結果が変わります。
""")

st.markdown("""
### スクリーニング済みExcelのダウンロード

フィルタで絞り込んだ後、テーブルの下にある

> **📥 スクリーニング済み Excel（○○件）**

をクリックすると、**絞り込んだ商品だけ**のExcelがダウンロードされます。

このExcelがリサーチの**最終成果物**です。
""")

st.warning("""
**注意：** 「結果Excelを手動ダウンロード」と「スクリーニング済みExcel」は別物です。

| ボタン | 内容 |
|--------|------|
| 📥 結果Excel | 全商品（フィルタなし） |
| 📥 スクリーニング済みExcel | **フィルタ後の商品のみ** ← こちらが最終成果物 |
""")

st.divider()

# ── 出力Excelの列説明 ──
st.header("出力Excelの列と計算式")
st.markdown("""
| 列 | 項目 | 内容 |
|----|------|------|
| D | 商品名 | 日本語の商品名 |
| E | タイトル | US Amazonの英語タイトル |
| G | ASIN | US AmazonのASIN |
| N | 卸価格 | 仕入れ単価 |
| O | 30キーパ | 過去30日間のランク下落回数（≒販売回数） |
| P | セラー数 | 合計オファー数 |
| T | 売価USD | Buy Box価格（米ドル） |
| V | Amazon手数料 | FBA手数料（紹介料込み） |
| X | Weight | パッケージ重量（g） |

**計算列：**

| 列 | 数式 | 内容 |
|----|------|------|
| W | `=L×N×1.1` | 仕入値（税込） |
| U | `=(W+(X×3))÷為替レート` | 仕入＋送料（FBA） |
| Y | `=T-U-V` | 損益（USD） |
| Z | `=(O÷(P+1))×Y` | 利益額 |
| AA | `=Y÷T` | 利益率 |
| AB〜AG | 利益額2・売上・手数料・原価・原価送料・利益/円 | |

**定数（AH2=為替レート, AI2=送料3円/g）はExcel上で変更すると全行再計算されます。**
""")

st.divider()

# ── Keepaエクスポートの注意 ──
st.header("Keepaエクスポートの注意点")
st.info("""
**「過去30日間の減少」列を必ず含めてください**

Keepa製品ビューアで「列を設定」→「過去30日間の減少」にチェック。
この列がないと30キーパの値が空になります。
""")

st.info("""
**10,000件制限について**

Keepa製品ビューアは1回の検索で最大10,000件です。
94,000件ある場合は10回に分けてエクスポートし、
Amaproの②に全ファイルをまとめてアップロードしてください。
""")
