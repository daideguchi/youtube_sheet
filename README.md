# YouTube チャンネル分析ツール

Google Apps Script (GAS) を使った YouTube チャンネルのベンチマーク分析ツールです。

## 機能

- **チャンネル情報取得**: ハンドル名（@username）からチャンネル情報を自動取得
- **統計分析**: 登録者数、視聴回数、動画数の統計計算
- **ベンチマークレポート**: 比較分析レポートの自動生成
- **ジャンル別分析**: カテゴリごとの詳細分析

## ファイル構成

### GAS 実行ファイル

- `2_benchmark.gs` - メインの YouTube ベンチマークツール
- `4_channelCheck.gs` - 詳細チャンネル分析機能
- `oauth_handler.gs.gs` - OAuth 認証処理

### 開発環境（Cursor/VSCode 用）

- `package.json` - 依存関係とスクリプト定義
- `test-setup.js` - GAS モック設定
- `benchmark-functions.js` - テスト用関数抽出
- `__tests__/` - Jest テストファイル
- `.eslintrc.json` - ESLint 設定
- `.vscode/settings.json` - エディタ設定

## 使い方

### GAS ブラウザでの実行

1. [Google Apps Script](https://script.google.com) にアクセス
2. 新しいプロジェクトを作成
3. `*.gs` ファイルの内容をコピー&ペースト
4. YouTube Data API キーを設定
5. スプレッドシートで実行

### 開発環境でのテスト

```bash
# 依存関係インストール
npm install

# テスト実行
npm test

# リアルタイムテスト
npm run test:watch

# コードフォーマット
npm run format

# Lint実行
npm run lint
```

## セットアップ

### 1. YouTube Data API キー取得

1. [Google Cloud Console](https://console.cloud.google.com/)でプロジェクト作成
2. YouTube Data API v3 を有効化
3. API キーを作成
4. GAS のメニューから「① API 設定・テスト」でキーを設定

### 2. スプレッドシート準備

1. 新しい Google スプレッドシートを作成
2. A 列: ジャンル（任意）
3. B 列: YouTube ハンドル名（@から始まる）

### 3. 実行

1. 「② チャンネル情報取得」を実行
2. 「③ ベンチマークレポート作成」でレポート生成

## API 制限

- YouTube Data API の 1 日あたりの呼び出し制限があります
- 大量のチャンネルを分析する場合は注意してください

## ライセンス

MIT

## 作成者

Claude AI
