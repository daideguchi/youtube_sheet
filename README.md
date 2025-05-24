# 📊 YouTube Channel Analysis Tool

Google Sheets 上で YouTube チャンネルのデータ分析を行うための Google Apps Script (GAS) ツールです。

## 🎯 主要機能

- **📈 チャンネル統計分析** - 登録者数、視聴回数、動画数の取得・分析
- **📊 動画別詳細分析** - 各動画の再生回数、高評価率、コメント数分析
- **⚡ 自動データ更新** - 定期的なデータ取得とスプレッドシート更新
- **📋 レポート生成** - 分析結果の可視化とレポート出力
- **🔧 ベンチマーク機能** - パフォーマンス測定と最適化

## 🚀 セットアップ

### 1. Google Apps Script プロジェクト作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 新しいプロジェクトを作成
3. プロジェクト名を「YouTube Channel Analysis」に設定

### 2. スクリプトファイルのアップロード

```bash
# Google Apps Script CLI (clasp) を使用
npm install -g @google/clasp
clasp login
clasp create --type sheets --title "YouTube Channel Analysis"
clasp push
```

### 3. YouTube Data API 設定

1. [Google Cloud Console](https://console.cloud.google.com/) でプロジェクト作成
2. YouTube Data API v3 を有効化
3. API キーを取得
4. スクリプトプロパティに `YOUTUBE_API_KEY` として設定

## 📁 ファイル構成

```
youtube-sheet-gas/
├── 2_benchmark.gs           # ベンチマーク・パフォーマンス測定
├── 4_channelCheck.gs        # チャンネル分析メイン機能
├── package.json             # プロジェクト設定
├── .clasp.json             # Google Apps Script設定
├── appsscript.json         # GASマニフェスト
├── __tests__/              # テストファイル
├── .cursor/rules/          # 開発ルール
└── @todo.md               # タスク管理
```

## 🛠️ 開発環境

### 必要なツール

- **Node.js** (v16 以上)
- **Google Apps Script CLI (clasp)**
- **Google Sheets** (データ表示用)

### セットアップコマンド

```bash
# 依存関係インストール
npm install

# リンティング
npm run lint

# テスト実行
npm test

# 自動同期（開発時）
./auto-clasp-sync.sh
```

## 📊 使用方法

### 1. 基本的なチャンネル分析

```javascript
// Google Sheetsのセルから実行
=channelAnalysis("チャンネルID")
```

### 2. 詳細な動画分析

```javascript
// 特定チャンネルの全動画分析
=videoAnalysis("チャンネルID", 50)  // 最新50動画
```

### 3. ベンチマーク実行

```javascript
// パフォーマンス測定
=runBenchmark()
```

## 📈 機能詳細

### チャンネル統計

- 登録者数の推移
- 総視聴回数
- 動画投稿頻度
- 平均視聴回数

### 動画分析

- 動画別再生回数
- 高評価・低評価率
- コメント数と反応率
- 投稿時期と視聴数の相関

### レポート機能

- データの可視化（グラフ生成）
- 週次・月次レポート
- パフォーマンス比較分析

## 🔧 設定

### Google Apps Script プロパティ

| プロパティ名       | 説明                  | 必須 |
| ------------------ | --------------------- | ---- |
| `YOUTUBE_API_KEY`  | YouTube Data API キー | ✅   |
| `REFRESH_INTERVAL` | データ更新間隔（分）  | ❌   |
| `MAX_VIDEOS`       | 最大取得動画数        | ❌   |

### スプレッドシート設定

1. 新しい Google Sheets を作成
2. GAS プロジェクトをスプレッドシートにバインド
3. 必要なシートタブを作成（Dashboard, Videos, Analytics）

## 🧪 テスト

```bash
# 全テスト実行
npm test

# 監視モード
npm run test:watch

# 特定ファイルのテスト
npx jest benchmark-functions.test.js
```

## 📋 開発ガイドライン

### コーディング規約

- Google Apps Script のベストプラクティスに従う
- JSDoc コメントの記述
- エラーハンドリングの実装
- ログ出力の統一

### Git ワークフロー

```bash
# 機能開発
git checkout -b feature/new-analysis
git commit -m "feat: add new analysis function"
git push origin feature/new-analysis

# リリース
git checkout main
git merge feature/new-analysis
git tag v1.1.0
```

## 🔄 自動更新

プロジェクトには自動同期スクリプトが含まれています：

```bash
# バックグラウンドで自動同期開始
./auto-clasp-sync.sh &

# 手動同期
clasp push
```

## 📝 ライセンス

MIT License

## 🤝 コントリビューション

1. このリポジトリをフォーク
2. 機能ブランチを作成 (`git checkout -b feature/AmazingFeature`)
3. 変更をコミット (`git commit -m 'Add some AmazingFeature'`)
4. ブランチにプッシュ (`git push origin feature/AmazingFeature`)
5. プルリクエストを作成

## 📞 サポート

問題や質問がある場合は、GitHub の Issues でお知らせください。
