# 🤖 AI Tools Directory

このディレクトリには、YouTube Channel Analysis Tool で使用する AI 関連のファイルが整理されています。

## 📁 ディレクトリ構造

```
ai-tools/
├── global-system/          # グローバル AI 切り替えシステム
│   ├── global-ai-switcher.sh   # メイン AI 切り替えスクリプト
│   └── install-ai-switcher.sh  # ワンライナーインストール
├── project-config/         # プロジェクト固有設定
│   ├── .claude-code-config.json # Claude Code 設定
│   └── .claudecodeignore       # Claude Code 除外設定
├── workflows/              # ワークフロー & 使用方法
│   └── claude-workflows.md    # Claude Code ワークフローガイド
├── docs/                   # ドキュメント
│   └── README-AI-SWITCHER.md  # 詳細ドキュメント
└── README.md              # このファイル
```

## 🚀 クイックスタート

### Cursor AI（軽量・高速）

```bash
# Cursor IDE でプロジェクトを開く
cursor .

# AI チャット開始: Ctrl+L (macOS: Cmd+L)
# .cursor/rules/ の設定が自動適用されます
```

### Claude Code（大規模分析）

```bash
# 使用例
claude "このプロジェクトの全体構造を分析して"
claude "4_channelCheck.gsのパフォーマンスを最適化して"
claude "テスト戦略を立案して"
```

## 🎯 AI 使い分けガイド

| 用途                     | 推奨 AI     | 理由                     | 例                               |
| ------------------------ | ----------- | ------------------------ | -------------------------------- |
| **日常コーディング**     | Cursor AI   | 軽量・高速・リアルタイム | バグ修正、小規模機能追加         |
| **大規模ファイル分析**   | Claude Code | 深い分析・包括的理解     | 7876 行の 4_channelCheck.gs 分析 |
| **アーキテクチャ設計**   | Claude Code | 構造的な改善提案         | モジュール分割、責任分離         |
| **パフォーマンス最適化** | Claude Code | 全体最適化視点           | 実行時間削減、メモリ効率化       |
| **デバッグ・エラー解決** | Cursor AI   | 素早い対応               | エラーメッセージ解釈、即座の修正 |

## 🔧 高度な使用方法

### プロジェクト間での AI システム共有

```bash
# 他のプロジェクトに AI システムを導入
cd /path/to/another-project
ai-setup

# プロジェクトタイプ選択
# 1) Google Apps Script
# 2) Web Development
# 3) Python
# 4) Node.js
# 5) Generic
```

### Claude Code 設定のカスタマイズ

`project-config/.claude-code-config.json` を編集：

```json
{
  "project_context": {
    "name": "YouTube Channel Analysis Tool",
    "type": "google-apps-script",
    "main_files": ["4_channelCheck.gs", "2_benchmark.gs"],
    "constraints": {
      "execution_time_limit": "6 minutes",
      "quota_management": true
    }
  }
}
```

## 📖 詳細ドキュメント

- [詳細設定ガイド](docs/README-AI-SWITCHER.md)
- [Claude Code ワークフロー](workflows/claude-workflows.md)

## 🛠 トラブルシューティング

### Claude Code が動作しない

```bash
# API キー確認
echo $ANTHROPIC_API_KEY

# 再インストール
npm install -g @anthropic-ai/claude-code
```

### Cursor AI の設定が反映されない

```bash
# Cursor を再起動
# .cursor/rules/ ディレクトリを確認
ls -la .cursor/rules/
```
