#!/bin/bash

# YouTube Channel Analysis Project - Auto Git Management Script
# このスクリプトは .cursor/rules/rules.mdc ファイルの変更を自動でGitに反映します

set -e  # エラー時に終了

# 設定
RULES_FILE=".cursor/rules/rules.mdc"
WATCH_FILES=(".cursor/rules/rules.mdc" "README.md" "package.json")

# 色設定
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# ログ関数
log_info() {
    echo -e "${BLUE}[INFO]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[SUCCESS]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[WARNING]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Gitリポジトリの確認
check_git_repo() {
    if ! git rev-parse --git-dir > /dev/null 2>&1; then
        log_error "Gitリポジトリではありません"
        exit 1
    fi
    log_info "Gitリポジトリを確認しました"
}

# 作業ディレクトリの状態確認
check_working_tree() {
    if ! git diff-index --quiet HEAD --; then
        log_warning "作業ディレクトリに未コミットの変更があります"
        git status --porcelain
        return 1
    fi
    return 0
}

# rules.mdcファイルの変更をコミット
commit_rules_changes() {
    local commit_message="Update project rules: $(date '+%Y-%m-%d %H:%M')"
    
    # ファイルの存在確認
    if [[ ! -f "$RULES_FILE" ]]; then
        log_error "rules.mdcファイルが見つかりません: $RULES_FILE"
        return 1
    fi
    
    # 変更の確認
    if git diff --quiet "$RULES_FILE"; then
        log_info "rules.mdcファイルに変更はありません"
        return 0
    fi
    
    log_info "rules.mdcファイルの変更を検出しました"
    
    # ステージング
    git add "$RULES_FILE"
    log_info "ファイルをステージングしました: $RULES_FILE"
    
    # コミット
    git commit -m "$commit_message"
    log_success "コミットしました: $commit_message"
    
    # プッシュ
    if git push origin main; then
        log_success "リモートリポジトリにプッシュしました"
    else
        log_error "プッシュに失敗しました"
        return 1
    fi
    
    return 0
}

# 指定ファイルの変更を一括コミット
commit_config_changes() {
    local changed_files=()
    local commit_message="Update configuration files: $(date '+%Y-%m-%d %H:%M')"
    
    # 変更されたファイルを検出
    for file in "${WATCH_FILES[@]}"; do
        if [[ -f "$file" ]] && ! git diff --quiet "$file" 2>/dev/null; then
            changed_files+=("$file")
        fi
    done
    
    if [[ ${#changed_files[@]} -eq 0 ]]; then
        log_info "監視対象ファイルに変更はありません"
        return 0
    fi
    
    log_info "変更されたファイル: ${changed_files[*]}"
    
    # ステージング
    git add "${changed_files[@]}"
    log_info "ファイルをステージングしました"
    
    # コミット
    git commit -m "$commit_message"
    log_success "コミットしました: $commit_message"
    
    # プッシュ
    if git push origin main; then
        log_success "リモートリポジトリにプッシュしました"
    else
        log_error "プッシュに失敗しました"
        return 1
    fi
    
    return 0
}

# ファイル監視モード
watch_files() {
    log_info "ファイル監視モードを開始します..."
    log_info "監視対象: ${WATCH_FILES[*]}"
    log_info "Ctrl+Cで終了します"
    
    # fswatch が利用可能かチェック
    if ! command -v fswatch &> /dev/null; then
        log_warning "fswatch が見つかりません。代替手段でポーリングします"
        watch_files_polling
        return
    fi
    
    # fswatch を使用してファイル監視
    fswatch -o "${WATCH_FILES[@]}" | while read num; do
        log_info "ファイル変更を検出しました"
        sleep 1  # ファイル書き込み完了を待つ
        commit_config_changes
    done
}

# ポーリングによるファイル監視（fswatch がない場合）
watch_files_polling() {
    local -A last_modified
    
    # 初期の更新時刻を記録
    for file in "${WATCH_FILES[@]}"; do
        if [[ -f "$file" ]]; then
            last_modified["$file"]=$(stat -f "%m" "$file" 2>/dev/null || echo "0")
        fi
    done
    
    while true; do
        sleep 5  # 5秒間隔でチェック
        
        for file in "${WATCH_FILES[@]}"; do
            if [[ -f "$file" ]]; then
                current_modified=$(stat -f "%m" "$file" 2>/dev/null || echo "0")
                if [[ "${last_modified["$file"]:-0}" != "$current_modified" ]]; then
                    log_info "ファイル変更を検出しました: $file"
                    last_modified["$file"]=$current_modified
                    commit_config_changes
                    break
                fi
            fi
        done
    done
}

# メイン処理
main() {
    log_info "YouTube Channel Analysis Project - Auto Git Management"
    log_info "=================================================="
    
    # Gitリポジトリの確認
    check_git_repo
    
    # 引数によるモード切り替え
    case "${1:-auto}" in
        "rules")
            log_info "rules.mdcファイルの変更をコミットします"
            commit_rules_changes
            ;;
        "config")
            log_info "設定ファイルの変更をコミットします"
            commit_config_changes
            ;;
        "watch")
            watch_files
            ;;
        "auto"|*)
            log_info "自動モード: 設定ファイルの変更をチェックしてコミットします"
            commit_config_changes
            ;;
    esac
}

# スクリプト実行
main "$@" 