#!/bin/bash

# YouTube Channel Analysis Project - Auto clasp Sync Script
# Google Apps Script ファイルの変更を自動でclaspにプッシュします

set -e

# 色設定
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

# 監視対象ファイル
WATCH_FILES=(
    "0_dashboard.gs"
    "2_benchmark.gs" 
    "4_channelCheck.gs"
    "test_dashboard.gs"
    "oauth_handler.gs.gs"
    "1_roadmap.gs"
    "appsscript.json"
)

# ログ関数
log_info() {
    echo -e "${BLUE}[CLASP-SYNC]${NC} $1"
}

log_success() {
    echo -e "${GREEN}[CLASP-SYNC]${NC} $1"
}

log_warning() {
    echo -e "${YELLOW}[CLASP-SYNC]${NC} $1"
}

log_error() {
    echo -e "${RED}[CLASP-SYNC]${NC} $1"
}

# clasp 認証確認
check_clasp_auth() {
    if ! clasp login --status &>/dev/null; then
        log_error "clasp 認証が必要です: clasp login を実行してください"
        exit 1
    fi
    log_info "clasp 認証確認完了"
}

# ファイル変更時の自動プッシュ
auto_push() {
    log_info "Google Apps Script ファイルの変更を検出しました"
    
    # 少し待ってからプッシュ（ファイル書き込み完了を待つ）
    sleep 2
    
    if clasp push; then
        log_success "Google Apps Script プロジェクトにプッシュしました: $(date)"
    else
        log_error "clasp push に失敗しました"
        return 1
    fi
}

# ファイル監視モード（fswatch使用）
watch_files_fswatch() {
    log_info "fswatch を使用してファイル監視を開始します..."
    log_info "監視対象: ${WATCH_FILES[*]}"
    log_info "Ctrl+Cで終了します"
    
    fswatch -o "${WATCH_FILES[@]}" | while read num; do
        auto_push
    done
}

# ファイル監視モード（ポーリング）
watch_files_polling() {
    log_warning "fswatch が見つかりません。ポーリングモードで監視します"
    log_info "監視対象: ${WATCH_FILES[*]}"
    log_info "Ctrl+Cで終了します"
    
    local -A last_modified
    
    # 初期の更新時刻を記録
    for file in "${WATCH_FILES[@]}"; do
        if [[ -f "$file" ]]; then
            last_modified["$file"]=$(stat -f "%m" "$file" 2>/dev/null || echo "0")
        fi
    done
    
    while true; do
        sleep 3  # 3秒間隔でチェック
        
        for file in "${WATCH_FILES[@]}"; do
            if [[ -f "$file" ]]; then
                current_modified=$(stat -f "%m" "$file" 2>/dev/null || echo "0")
                if [[ "${last_modified["$file"]:-0}" != "$current_modified" ]]; then
                    log_info "ファイル変更検出: $file"
                    last_modified["$file"]=$current_modified
                    auto_push
                    break
                fi
            fi
        done
    done
}

# 手動プッシュ
manual_push() {
    log_info "手動プッシュを実行します..."
    auto_push
}

# 手動プル
manual_pull() {
    log_info "Google Apps Script プロジェクトからプルします..."
    
    if clasp pull; then
        log_success "Google Apps Script プロジェクトからプルしました: $(date)"
    else
        log_error "clasp pull に失敗しました"
        return 1
    fi
}

# プロジェクト情報表示
show_info() {
    log_info "Google Apps Script プロジェクト情報"
    echo "=================================="
    clasp list 2>/dev/null || log_warning "プロジェクト一覧の取得に失敗"
    echo ""
    if [[ -f ".clasp.json" ]]; then
        echo "現在のプロジェクト設定:"
        cat .clasp.json
    else
        log_warning ".clasp.json ファイルが見つかりません"
    fi
}

# メイン処理
main() {
    log_info "YouTube Channel Analysis - Auto clasp Sync"
    log_info "============================================="
    
    # clasp認証確認
    check_clasp_auth
    
    # 引数によるモード切り替え
    case "${1:-watch}" in
        "push")
            manual_push
            ;;
        "pull")
            manual_pull
            ;;
        "watch")
            if command -v fswatch &> /dev/null; then
                watch_files_fswatch
            else
                watch_files_polling
            fi
            ;;
        "info")
            show_info
            ;;
        *)
            echo "使用方法: $0 [push|pull|watch|info]"
            echo ""
            echo "  push  - 手動プッシュ"
            echo "  pull  - 手動プル"  
            echo "  watch - ファイル監視モード（デフォルト）"
            echo "  info  - プロジェクト情報表示"
            ;;
    esac
}

# スクリプト実行
main "$@" 