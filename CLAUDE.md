# Amazon Keepa App

## GitHubリポジトリ

https://github.com/TTTT-47582/amazon-keepa-app.git

## プロジェクト概要

Keepa API からデータを取得し、**日本 Amazon で仕入れ → 米国 Amazon FBA で販売**できる商品を自動リサーチするツール。
販売ランク・価格・レビューなどの条件でフィルタリングし、利益計算結果をリスト表示する。

## 技術スタック

- **言語**: Python 3.11+
- **UI フレームワーク**: Streamlit（無料デプロイ: Streamlit Community Cloud）
- **Keepa クライアント**: `keepa` ライブラリ
- **設定管理**: PyYAML（`config/screening_filters.yaml`）
- **パッケージ管理**: pip + venv（または `uv`）
- **環境変数**: `python-dotenv`

## ディレクトリ構成

```
amazon-keepa-app/
├── app.py                          # Streamlit メインアプリ（起動エントリポイント）
├── src/
│   ├── __init__.py
│   ├── keepa_client.py             # Keepa API ラッパー（検索・データ取得）
│   ├── profit_calc.py              # 利益計算ロジック
│   └── config_manager.py          # YAML 設定ファイルの読み書き
├── config/
│   └── screening_filters.yaml     # スクリーニング条件（後から追加・変更可能）
├── .env                            # APIキー（git 管理外）
├── .env.example                    # APIキーのテンプレート
├── .gitignore
├── requirements.txt
└── CLAUDE.md
```

## セットアップ

```bash
# 依存パッケージのインストール
pip install -r requirements.txt

# .env ファイルを作成して API キーを設定
cp .env.example .env
# → .env を開いて KEEPA_API_KEY を設定

# アプリ起動
streamlit run app.py
```

## 環境変数

`.env` ファイルで管理（**絶対に git commit しない**）:

```
KEEPA_API_KEY=your_keepa_api_key_here
```

> Keepa API キーは https://keepa.com/#!api で取得。
> `product_finder` は Professional プラン以上が必要。

## スクリーニング条件の追加・変更

`config/screening_filters.yaml` を直接編集するだけで条件を追加できます。
アプリを再起動すると反映されます。

## Git 運用ルール

### 基本方針

- **コードを変更するたびに GitHub へ push する**
- `main` ブランチを常に動作する状態に保つ
- 機能追加・バグ修正は feature ブランチで行い、完了後に `main` へマージ

### コミット手順

```bash
git add <変更ファイル>      # -A や . は使わない（.env の誤コミット防止）
git commit -m "簡潔な変更内容"
git push origin <ブランチ名>
```

### コミットメッセージ規則

- `feat:` 新機能追加
- `fix:` バグ修正
- `refactor:` リファクタリング
- `docs:` ドキュメント変更
- `test:` テスト追加・修正

例: `feat: add price drop filter to screening`

### ブランチ命名

- `feature/<機能名>` — 新機能
- `fix/<バグ内容>` — バグ修正

### 必須: push 前チェック

1. `.env` や API キーを含むファイルが `git status` に含まれていないか確認
2. `git diff --staged` で変更内容を確認してから commit
3. commit 直後に `git push` を実行

## 開発ガイドライン

- Keepa API はトークン消費が大きいため、`@st.cache_data(ttl=3600)` で 1 時間キャッシュする
- API キーは `.env` で管理し、コード内にハードコードしない
- コメントは日本語で記載する
- 型ヒントを積極的に使う（Python 3.11+ の `X | Y` 記法を使用）
