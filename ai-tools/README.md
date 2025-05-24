# 🤖 AI Tools Directory

このディレクトリには、YouTube Channel Analysis Tool で使用する AI 関連のファイルが整理されています。

## 📁 ディレクトリ構造

```
ai-tools/
├── global-system/          # グローバル AI 切り替えシステム
├── project-config/         # プロジェクト固有設定
├── workflows/              # ワークフロー & 使用方法
├── docs/                   # ドキュメント
└── README.md              # このファイル
```

## 🚀 使用方法

### Cursor AI (軽量・高速)

- Cursor IDE でプロジェクトを開く
- `Ctrl+L` (macOS: `Cmd+L`) で AI チャット開始
- `.cursor/rules/` の設定が自動適用

### Claude Code (大規模分析)

```bash
# 使用例
claude "このプロジェクトを分析して"
claude "4_channelCheck.gsを最適化して"
```

## 🤝 AI 使い分けガイド

| 用途               | 推奨 AI     | 理由                     |
| ------------------ | ----------- | ------------------------ |
| 日常コーディング   | Cursor AI   | 軽量・高速・リアルタイム |
| 大規模ファイル分析 | Claude Code | 深い分析・包括的理解     |
| リファクタリング   | Claude Code | 構造的な改善提案         |
| バグ修正           | Cursor AI   | 素早い対応               |
