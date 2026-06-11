# Amazon Keepa App

## GitHubリポジトリ

https://github.com/TTTT-47582/amazon-keepa-app.git

Keepa API からデータを抽出し、Amazon 商品検索を自動化するツール。

## プロジェクト概要

- Keepa API を使って商品価格履歴・ランキング・在庫情報を取得
- 条件指定による商品の自動リサーチ・絞り込み
- 取得データの保存・分析・可視化

## 技術スタック

- **言語**: Python 3.11+
- **Keepa クライアント**: `keepa` ライブラリ
- **データ処理**: pandas, numpy
- **保存**: SQLite（ローカル）または PostgreSQL
- **可視化 / UI**: Streamlit または CLI（Click）
- **スケジューリング**: APScheduler または cron
- **パッケージ管理**: `uv`（推奨）または pip + venv

## ディレクトリ構成

```
amazon-keepa-app/
├── src/
│   ├── keepa_client.py   # Keepa API ラッパー
│   ├── search.py         # 商品検索・フィルタリングロジック
│   ├── storage.py        # データ保存・取得
│   └── scheduler.py      # 自動化スケジューラ
├── scripts/              # 単発実行スクリプト
├── tests/
├── .env.example          # 環境変数テンプレート（APIキー等）
├── requirements.txt
└── README.md
```

## 環境変数

`.env` ファイルで管理（**絶対に git commit しない**）:

```
KEEPA_API_KEY=your_api_key_here
```

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

例: `feat: add price drop filter to product search`

### ブランチ命名

- `feature/<機能名>` — 新機能
- `fix/<バグ内容>` — バグ修正

### 必須: push 前チェック

1. `.env` や APIキーを含むファイルが `git status` に含まれていないか確認
2. `git diff --staged` で変更内容を確認してから commit
3. commit 直後に `git push` を実行

## 開発ガイドライン

- Keepa API はトークン消費が大きいため、開発中はレスポンスをローカルにキャッシュする
- APIキーは `.env` で管理し、コード内にハードコードしない
- テストはモック or キャッシュ済みデータを使い、実 API を叩かない
- 型ヒントを積極的に使う
