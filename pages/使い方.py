"""使い方・利益計算の仕組みページ"""

import streamlit as st

st.set_page_config(page_title="使い方", page_icon="📖", layout="wide")

st.title("📖 使い方・利益計算の仕組み")

st.header("セットアップ")
st.markdown("""
1. プロジェクトルートに `.env` ファイルを作成
2. `KEEPA_API_KEY=your_key_here` を記載して保存
3. `pip install -r requirements.txt` で依存パッケージをインストール
4. `streamlit run app.py` でアプリを起動
""")

st.divider()

st.header("検索の流れ")
st.markdown("""
1. サイドバーに **Keepa API キー** を入力
2. **販売ランク・価格帯・レビュー条件** を設定
3. 「商品を検索する」ボタンを押すと Keepa から商品を取得
4. 結果は **利益率の高い順** にリスト表示されます
""")

st.divider()

st.header("利益計算の内訳")
st.markdown("""
| 項目 | 内容 |
|------|------|
| 仕入れ価格 | 日本 Amazon の現在最安値（Keepa より取得） |
| 推定販売価格 | 仕入れ USD × 販売倍率（サイドバーで調整） |
| 国際送料 | 重量（kg）× 送料単価（config で変更可） |
| Amazon 紹介料 | 販売価格 × 15%（概算） |
| FBA 手数料 | 重量・サイズに応じた 2024 年 US 料金表ベース |
| **推定利益** | **販売価格 − 仕入れ − 国際送料 − Amazon 各手数料** |
""")

st.divider()

st.header("スクリーニング条件の追加・変更")
st.markdown("""
`config/screening_filters.yaml` を直接編集するだけで条件を追加できます。
アプリを再起動すると反映されます。

```yaml
filters:
  sales_rank_max: 50000   # 販売ランク上限
  price_min: 500          # 仕入れ下限（円）
  price_max: 5000         # 仕入れ上限（円）
  rating_min: 3.5         # 最低評価
  review_count_min: 10    # 最低レビュー数
```
""")
