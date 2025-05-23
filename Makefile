# YouTube Channel Analysis Project - Makefile
# プロジェクトのGit自動管理とタスク実行用

.PHONY: help rules-commit config-commit watch-files auto-commit push status clean install clasp-push clasp-pull clasp-watch clasp-info

# デフォルトターゲット
help: ## このヘルプメッセージを表示
	@echo "YouTube Channel Analysis Project - Git Auto Management"
	@echo "====================================================="
	@echo ""
	@echo "利用可能なコマンド:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-20s\033[0m %s\n", $$1, $$2}'
	@echo ""
	@echo "使用例:"
	@echo "  make rules-commit    # rules.mdcの変更をコミット"
	@echo "  make clasp-watch     # Google Apps Script自動同期"
	@echo "  make auto-commit     # 自動コミット実行"

# =============================================================================
# Git 自動管理コマンド
# =============================================================================

# rules.mdcファイルの自動コミット
rules-commit: ## rules.mdcファイルの変更をGitにコミット・プッシュ
	@echo "📝 rules.mdcファイルの変更をコミットします..."
	@./auto-commit-rules.sh rules

# 設定ファイルの自動コミット
config-commit: ## 設定ファイル全体の変更をGitにコミット・プッシュ
	@echo "⚙️ 設定ファイルの変更をコミットします..."
	@./auto-commit-rules.sh config

# ファイル監視モード
watch-files: ## ファイル変更の監視モードを開始
	@echo "👀 ファイル監視モードを開始します..."
	@./auto-commit-rules.sh watch

# 自動コミット（デフォルト）
auto-commit: ## 変更ファイルの自動検出・コミット・プッシュ
	@echo "🚀 自動コミットを実行します..."
	@./auto-commit-rules.sh auto

# 手動プッシュ
push: ## リモートリポジトリに手動プッシュ
	@echo "⬆️ リモートリポジトリにプッシュします..."
	@git push origin main

# =============================================================================
# Google Apps Script (clasp) 管理コマンド
# =============================================================================

# clasp 手動プッシュ
clasp-push: ## Google Apps Scriptプロジェクトに手動プッシュ
	@echo "🚀 Google Apps Scriptにプッシュします..."
	@chmod +x auto-clasp-sync.sh
	@./auto-clasp-sync.sh push

# clasp 手動プル
clasp-pull: ## Google Apps Scriptプロジェクトから手動プル
	@echo "⬇️ Google Apps Scriptからプルします..."
	@chmod +x auto-clasp-sync.sh
	@./auto-clasp-sync.sh pull

# clasp 自動監視
clasp-watch: ## Google Apps Scriptファイルの自動監視・プッシュ
	@echo "👀 Google Apps Scriptファイルの自動監視を開始します..."
	@chmod +x auto-clasp-sync.sh
	@./auto-clasp-sync.sh watch

# clasp プロジェクト情報
clasp-info: ## Google Apps Scriptプロジェクト情報を表示
	@echo "📋 Google Apps Scriptプロジェクト情報..."
	@chmod +x auto-clasp-sync.sh
	@./auto-clasp-sync.sh info

# =============================================================================
# その他のコマンド
# =============================================================================

# Git状態確認
status: ## Git状態とプロジェクト情報を表示
	@echo "📊 プロジェクト状態"
	@echo "===================="
	@echo ""
	@echo "🔍 Git Status:"
	@git status --short
	@echo ""
	@echo "📝 最新コミット:"
	@git log --oneline -3
	@echo ""
	@echo "📁 監視対象ファイル:"
	@ls -la .cursor/rules/rules.mdc README.md 2>/dev/null || echo "一部ファイルが見つかりません"

# クリーンアップ
clean: ## 一時ファイルとバックアップファイルを削除
	@echo "🧹 一時ファイルを削除します..."
	@find . -name "*.bak*" -type f -delete
	@find . -name ".DS_Store" -type f -delete
	@echo "✅ クリーンアップ完了"

# 開発環境セットアップ
install: ## 必要なツールのインストール確認
	@echo "🔧 開発環境を確認します..."
	@echo ""
	@echo "Git: $(shell git --version 2>/dev/null || echo '❌ インストールされていません')"
	@echo "Bash: $(shell bash --version | head -1 2>/dev/null || echo '❌ インストールされていません')"
	@echo "fswatch: $(shell fswatch --version 2>/dev/null || echo '⚠️ オプション（ファイル監視用）')"
	@echo "clasp: $(shell clasp --version 2>/dev/null || echo '❌ npm install -g @google/clasp')"
	@echo ""
	@echo "💡 clasp をインストールするには:"
	@echo "   npm install -g @google/clasp"
	@echo "💡 fswatch をインストールするには:"
	@echo "   brew install fswatch"

# rules.mdc編集後の自動実行
rules-auto: rules-commit ## rules.mdc編集後の推奨操作
	@echo "✅ rules.mdcファイルの変更が完了しました"

# 開発完了時の一括処理
dev-complete: auto-commit clasp-push ## 開発完了時の一括Git・clasp操作
	@echo "🎉 開発完了処理が完了しました"

# プロジェクト情報表示
info: ## プロジェクト情報を表示
	@echo "📋 YouTube Channel Analysis Project"
	@echo "===================================="
	@echo ""
	@echo "📁 プロジェクトディレクトリ: $(PWD)"
	@echo "🌿 Git ブランチ: $(shell git branch --show-current 2>/dev/null || echo 'unknown')"
	@echo "📝 rules.mdc更新日: $(shell stat -f '%Sm' .cursor/rules/rules.mdc 2>/dev/null || echo '未確認')"
	@echo "🔧 Git自動化スクリプト: ./auto-commit-rules.sh"
	@echo "🚀 clasp自動化スクリプト: ./auto-clasp-sync.sh"
	@echo ""
	@echo "📖 使い方:"
	@echo "  Git: make watch-files で自動監視"
	@echo "  GAS: make clasp-watch で自動同期" 