/* eslint-disable */
/**
 * YouTube事業管理統合ダッシュボード
 * プロフェッショナル向けYouTube事業の包括的管理プラットフォーム
 *
 * 作成者: Claude AI
 * バージョン: 2.0
 * 最終更新: 2025-01-22
 */
/* eslint-enable */

/**
 * YouTube事業管理ダッシュボードのメイン起動
 */
function createOrShowMainDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getSheetByName("🎬 YouTube事業管理");
    
    if (!dashboardSheet) {
      createYouTubeBusinessDashboard();
    } else {
      ss.setActiveSheet(dashboardSheet);
      refreshDashboardData();
    }
  } catch (error) {
    Logger.log("ダッシュボード起動エラー: " + error.toString());
  }
}

/**
 * YouTube事業管理ダッシュボードを作成
 */
function createYouTubeBusinessDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    // 既存ダッシュボードがあれば削除
    var existingDashboard = ss.getSheetByName("🎬 YouTube事業管理");
    if (existingDashboard) {
      ss.deleteSheet(existingDashboard);
    }
    
    // 新しいダッシュボードシートを作成
    var dashboard = ss.insertSheet("🎬 YouTube事業管理", 0);
    
    // ========== ヘッダー：事業概要 ==========
    dashboard.getRange("A1:L1").merge();
    dashboard.getRange("A1").setValue("🎬 YouTube事業管理ダッシュボード");
    
    dashboard.getRange("A2:L2").merge();
    dashboard.getRange("A2").setValue("包括的チャンネル運営・競合分析・成長戦略管理プラットフォーム");
    
    dashboard.getRange("A3:L3").merge();
    dashboard.getRange("A3").setValue("最終更新: " + new Date().toLocaleString() + " | 総合管理画面");
    
    // ========== セクション1：チャンネル運営管理 ==========
    dashboard.getRange("A5").setValue("📊 チャンネル運営管理");
    
    // 運営管理ボタン
    var managementButtons = [
      ["🎯", "自チャンネル分析", "現在のパフォーマンス詳細分析", "analyzeOwnChannel", "A7"],
      ["📈", "成長指標追跡", "登録者・視聴数・収益の推移", "trackGrowthMetrics", "A8"],
      ["🎨", "コンテンツ戦略", "動画企画・投稿スケジュール管理", "manageContentStrategy", "A9"],
      ["💰", "収益分析", "AdSense・スポンサー・商品売上", "analyzeRevenue", "A10"]
    ];
    
    createButtonSection(dashboard, managementButtons, 7);
    
    // ========== セクション2：競合・市場分析 ==========
    dashboard.getRange("A12").setValue("🔍 競合・市場分析");
    
    var competitorButtons = [
      ["⚔️", "競合チャンネル分析", "同業界チャンネルとの詳細比較", "analyzeCompetitors", "A14"],
      ["📊", "ベンチマーク作成", "業界標準との比較レポート", "createBenchmarkReport", "A15"],
      ["🎭", "ニッチ分析", "特定分野での市場ポジション", "analyzeNiche", "A16"],
      ["📱", "トレンド調査", "急成長キーワード・話題分析", "researchTrends", "A17"]
    ];
    
    createButtonSection(dashboard, competitorButtons, 14);
    
    // ========== セクション3：成長戦略・計画 ==========
    dashboard.getRange("A19").setValue("🚀 成長戦略・計画");
    
    var strategyButtons = [
      ["🗺️", "成長ロードマップ", "6ヶ月〜1年の戦略計画", "createGrowthRoadmap", "A21"],
      ["🎪", "コラボ機会", "他チャンネルとの連携可能性", "findCollabOpportunities", "A22"],
      ["🏆", "目標設定・KPI", "具体的数値目標と追跡", "setKPITargets", "A23"],
      ["💡", "改善提案", "AIによる成長施策提案", "generateImprovements", "A24"]
    ];
    
    createButtonSection(dashboard, strategyButtons, 21);
    
    // ========== 右側：リアルタイムデータエリア ==========
    dashboard.getRange("H5").setValue("📈 事業パフォーマンス概要");
    
    // 主要KPI表示
    var kpiLabels = [
      "総登録者数", "月間総視聴時間", "月間収益", "投稿頻度",
      "平均エンゲージメント率", "チャンネル登録率", "視聴継続率", "収益単価（RPM）"
    ];
    
    for (var i = 0; i < kpiLabels.length; i++) {
      dashboard.getRange(7 + i, 8).setValue(kpiLabels[i]);
      dashboard.getRange(7 + i, 10).setValue("取得中...");
      dashboard.getRange(7 + i, 11).setValue(""); // ステータス
    }
    
    // ========== 業界ベンチマーク情報 ==========
    dashboard.getRange("H16").setValue("🏆 業界ベンチマーク（2024年基準）");
    
    var benchmarks = [
      ["小規模（〜1万人）", "エンゲージメント: 5%+, 月4本投稿"],
      ["中規模（1-10万人）", "エンゲージメント: 3%+, 週1本投稿"],
      ["大規模（10万人+）", "エンゲージメント: 1%+, 定期投稿"],
      ["収益化目安", "1000人+4000h視聴時間（基本条件）"],
      ["成長率目安", "月5-10%成長（健全な成長）"]
    ];
    
    for (var i = 0; i < benchmarks.length; i++) {
      dashboard.getRange(18 + i, 8).setValue(benchmarks[i][0]);
      dashboard.getRange(18 + i, 10, 1, 2).merge();
      dashboard.getRange(18 + i, 10).setValue(benchmarks[i][1]);
    }
    
    // ========== アクションアイテム・アラート ==========
    dashboard.getRange("H24").setValue("⚠️ 要注意アラート・推奨アクション");
    
    dashboard.getRange("H26").setValue("🔴 緊急対応:");
    dashboard.getRange("J26").setValue("確認中...");
    
    dashboard.getRange("H27").setValue("🟡 改善推奨:");
    dashboard.getRange("J27").setValue("確認中...");
    
    dashboard.getRange("H28").setValue("🟢 次のアクション:");
    dashboard.getRange("J28").setValue("確認中...");
    
    // ========== クイックアクセスツール ==========
    dashboard.getRange("A26").setValue("⚡ クイックアクセス");
    
    var quickTools = [
      ["🔑", "API設定", "setApiKey"],
      ["📥", "データ取得", "processHandles"],
      ["📋", "レポート生成", "createBenchmarkReport"],
      ["🔄", "更新", "refreshDashboardData"]
    ];
    
    for (var i = 0; i < quickTools.length; i++) {
      dashboard.getRange(28, 1 + (i * 2)).setValue(quickTools[i][0]);
      dashboard.getRange(28, 2 + (i * 2)).setValue(quickTools[i][1]);
      dashboard.getRange(28, 2 + (i * 2)).setBackground("#4285F4").setFontColor("white");
      dashboard.getRange(29, 2 + (i * 2)).setValue(quickTools[i][2]); // 隠し関数名
    }
    
    // フォーマットを適用
    formatYouTubeBusinessDashboard(dashboard);
    
    // 初期データを読み込み
    refreshDashboardData();
    
    // アクティブにする
    ss.setActiveSheet(dashboard);
    
    // 初回ガイド
    ui.alert(
      "🎬 YouTube事業管理ダッシュボード",
      "プロフェッショナル向けYouTube事業管理ダッシュボードが作成されました。\n\n" +
      "📊 チャンネル運営: 自チャンネルの詳細分析・管理\n" +
      "🔍 競合分析: 業界動向・競合比較\n" +
      "🚀 成長戦略: 戦略的計画・目標管理\n\n" +
      "まずは「API設定」から始めて、各機能をお試しください。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("YouTube事業ダッシュボード作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "エラー",
      "ダッシュボード作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ボタンセクションを作成する汎用関数
 */
function createButtonSection(sheet, buttons, startRow) {
  for (var i = 0; i < buttons.length; i++) {
    var row = startRow + i;
    
    // アイコン
    sheet.getRange(row, 1).setValue(buttons[i][0]);
    sheet.getRange(row, 1).setFontSize(16);
    
    // ボタンタイトル
    sheet.getRange(row, 2).setValue(buttons[i][1]);
    sheet.getRange(row, 2).setFontWeight("bold");
    
    // 説明
    sheet.getRange(row, 3, 1, 3).merge();
    sheet.getRange(row, 3).setValue(buttons[i][2]);
    sheet.getRange(row, 3).setFontSize(10);
    
    // 実行ボタン
    sheet.getRange(row, 6).setValue("▶ 実行");
    sheet.getRange(row, 6).setBackground("#34A853").setFontColor("white");
    sheet.getRange(row, 6).setHorizontalAlignment("center");
    
    // 隠し関数名
    sheet.getRange(row, 7).setValue(buttons[i][3]);
    sheet.getRange(row, 7).setFontColor("white").setFontSize(1);
  }
}

/**
 * YouTube事業ダッシュボードのフォーマット設定
 */
function formatYouTubeBusinessDashboard(sheet) {
  // ========== ヘッダー ==========
  sheet.getRange("A1").setFontSize(24).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("white")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A2").setFontSize(14)
    .setBackground("#e8f0fe").setFontColor("#1a73e8")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A3").setFontSize(11).setFontStyle("italic")
    .setHorizontalAlignment("center").setFontColor("#5f6368");
  
  // ========== セクションヘッダー ==========
  var sectionHeaders = ["A5", "A12", "A19", "A26", "H5", "H16", "H24"];
  for (var i = 0; i < sectionHeaders.length; i++) {
    sheet.getRange(sectionHeaders[i]).setFontSize(16).setFontWeight("bold")
      .setBackground("#f8f9fa")
      .setBorder(true, true, true, true, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // ========== ボタンエリア ==========
  // チャンネル運営管理
  sheet.getRange("A7:G10")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A7:A10").setBackground("#e8f5e8");
  
  // 競合・市場分析
  sheet.getRange("A14:G17")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A14:A17").setBackground("#fff3e0");
  
  // 成長戦略・計画
  sheet.getRange("A21:G24")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A21:A24").setBackground("#f3e5f5");
  
  // ========== データエリア ==========
  sheet.getRange("H7:L14")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("H7:H14").setBackground("#e3f2fd");
  
  // ベンチマークエリア
  sheet.getRange("H18:L22")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("H18:H22").setBackground("#fff8e1");
  
  // アラートエリア
  sheet.getRange("H26:L28")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("H26:H28").setBackground("#ffebee");
  
  // ========== クイックアクセス ==========
  sheet.getRange("A28:H28")
    .setBorder(true, true, true, true, false, false, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== 列幅設定 ==========
  sheet.setColumnWidth(1, 45);   // アイコン
  sheet.setColumnWidth(2, 140);  // タイトル
  sheet.setColumnWidth(3, 200);  // 説明
  sheet.setColumnWidth(4, 50);   // 説明続き
  sheet.setColumnWidth(5, 50);   // 説明続き
  sheet.setColumnWidth(6, 70);   // ボタン
  sheet.setColumnWidth(7, 1);    // 隠し列
  sheet.setColumnWidth(8, 140);  // ラベル
  sheet.setColumnWidth(9, 20);   // スペーサー
  sheet.setColumnWidth(10, 120); // 値
  sheet.setColumnWidth(11, 80);  // ステータス
  sheet.setColumnWidth(12, 50);  // 余白
  
  // ========== 行高設定 ==========
  sheet.setRowHeight(1, 45);
  sheet.setRowHeight(2, 30);
  
  // ボタン行
  for (var i = 7; i <= 24; i++) {
    if ([7,8,9,10,14,15,16,17,21,22,23,24].indexOf(i) !== -1) {
      sheet.setRowHeight(i, 35);
    }
  }
}

/**
 * ダッシュボードデータを更新
 */
function refreshDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("🎬 YouTube事業管理");
    
    if (!dashboard) return;
    
    // API接続状況確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    
    if (!apiKey) {
      dashboard.getRange("J26").setValue("APIキーが未設定です");
      dashboard.getRange("J27").setValue("「API設定」を実行してください");
      dashboard.getRange("J28").setValue("設定完了後にデータ取得開始");
      return;
    }
    
    // 全シートからデータを集計
    var sheets = ss.getSheets();
    var channelCount = 0;
    var totalSubscribers = 0;
    var totalViews = 0;
    var totalVideos = 0;
    var validCount = 0;
    
    // 分析シートとベンチマークデータを集計
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      
      // 個別分析シート
      if (sheetName.startsWith("分析_")) {
        channelCount++;
        try {
          var subscribers = extractNumber(sheets[i].getRange("C14").getValue());
          var views = extractNumber(sheets[i].getRange("C15").getValue());
          
          if (subscribers > 0) {
            totalSubscribers += subscribers;
            validCount++;
          }
          if (views > 0) totalViews += views;
        } catch (e) { /* エラー無視 */ }
      }
      
      // ベンチマークデータシート
      if (sheetName === "Sheet1" || sheetName === "シート1") {
        try {
          var data = sheets[i].getDataRange().getValues();
          for (var j = 1; j < data.length; j++) {
            if (data[j][2] && data[j][2] !== "チャンネルが見つかりません") {
              channelCount++;
              
              var subscribers = extractNumber(data[j][4]);
              var views = extractNumber(data[j][5]);
              var videos = extractNumber(data[j][6]);
              
              if (subscribers > 0) {
                totalSubscribers += subscribers;
                validCount++;
              }
              if (views > 0) totalViews += views;
              if (videos > 0) totalVideos += videos;
            }
          }
        } catch (e) { /* エラー無視 */ }
      }
    }
    
    // KPI値を更新
    if (validCount > 0) {
      var avgSubscribers = Math.round(totalSubscribers / validCount);
      var avgViews = Math.round(totalViews / validCount);
      var avgVideos = Math.round(totalVideos / validCount);
      var engagementRate = ((totalViews / validCount) / (totalSubscribers / validCount) * 100);
      
      // データを表示
      dashboard.getRange("J7").setValue(totalSubscribers.toLocaleString());
      dashboard.getRange("J8").setValue((totalViews / 1000000).toFixed(1) + "M回");
      dashboard.getRange("J9").setValue("計算中...");
      dashboard.getRange("J10").setValue((avgVideos * validCount / 4).toFixed(1) + "本/月");
      dashboard.getRange("J11").setValue(engagementRate.toFixed(2) + "%");
      dashboard.getRange("J12").setValue("計算中...");
      dashboard.getRange("J13").setValue("計算中...");
      dashboard.getRange("J14").setValue("計算中...");
      
      // ステータス判定
      var engagementStatus = engagementRate >= 3 ? "🟢 良好" : 
                            engagementRate >= 1 ? "🟡 平均" : "🔴 要改善";
      dashboard.getRange("K11").setValue(engagementStatus);
      
      // アラート・推奨アクションを更新
      updateAlerts(dashboard, channelCount, avgSubscribers, engagementRate);
      
    } else {
      // データなしの場合
      for (var i = 7; i <= 14; i++) {
        dashboard.getRange("J" + i).setValue("データなし");
      }
      dashboard.getRange("J26").setValue("分析データがありません");
      dashboard.getRange("J27").setValue("チャンネル分析を開始してください");
      dashboard.getRange("J28").setValue("「自チャンネル分析」を実行");
    }
    
    // 最終更新時間
    dashboard.getRange("A3").setValue("最終更新: " + new Date().toLocaleString() + " | 総合管理画面");
    
  } catch (error) {
    Logger.log("ダッシュボードデータ更新エラー: " + error.toString());
  }
}

/**
 * 数値を抽出するヘルパー関数
 */
function extractNumber(value) {
  if (!value) return 0;
  var numStr = value.toString().replace(/[,\s]/g, "");
  var num = parseInt(numStr);
  return isNaN(num) ? 0 : num;
}

/**
 * アラート・推奨アクションを更新
 */
function updateAlerts(sheet, channelCount, avgSubscribers, engagementRate) {
  var urgent = "";
  var recommended = "";
  var nextAction = "";
  
  // 緊急対応
  if (engagementRate < 1) {
    urgent = "エンゲージメント率が低下（" + engagementRate.toFixed(2) + "%）";
  } else if (channelCount === 0) {
    urgent = "分析データが不足しています";
  } else {
    urgent = "特に緊急事項はありません";
  }
  
  // 改善推奨
  if (avgSubscribers < 1000) {
    recommended = "収益化条件達成に向けた登録者増加施策";
  } else if (avgSubscribers < 10000) {
    recommended = "中規模チャンネル向け成長戦略の実行";
  } else {
    recommended = "ブランド化・収益多様化の検討";
  }
  
  // 次のアクション
  if (channelCount < 5) {
    nextAction = "競合チャンネル分析の拡充";
  } else {
    nextAction = "成長ロードマップの策定・実行";
  }
  
  sheet.getRange("J26").setValue(urgent);
  sheet.getRange("J27").setValue(recommended);
  sheet.getRange("J28").setValue(nextAction);
}

/**
 * 自チャンネル分析を実行
 */
function analyzeOwnChannel() {
  // 既存の個別分析機能を呼び出し
  analyzeExistingChannel();
}

/**
 * 成長指標追跡を実行
 */
function trackGrowthMetrics() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "成長指標追跡機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・時系列での登録者数推移\n" +
    "・視聴時間・収益の変化\n" +
    "・コンテンツパフォーマンス分析",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * コンテンツ戦略管理を実行
 */
function manageContentStrategy() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "コンテンツ戦略機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・動画企画スケジュール\n" +
    "・人気コンテンツ分析\n" +
    "・投稿最適化提案",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 収益分析を実行
 */
function analyzeRevenue() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "収益分析機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・AdSense収益分析\n" +
    "・スポンサーシップ管理\n" +
    "・商品売上追跡",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 競合チャンネル分析を実行
 */
function analyzeCompetitors() {
  // ベンチマーク分析ダッシュボードを表示
  showBenchmarkDashboard();
}

/**
 * ニッチ分析を実行
 */
function analyzeNiche() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "ニッチ分析機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・特定分野での市場ポジション\n" +
    "・ニッチキーワード分析\n" +
    "・ターゲット層の特定",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * トレンド調査を実行
 */
function researchTrends() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "トレンド調査機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・急上昇キーワード分析\n" +
    "・話題のコンテンツ調査\n" +
    "・シーズナルトレンド予測",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 成長ロードマップ作成を実行
 */
function createGrowthRoadmap() {
  // 既存のロードマップ機能を呼び出し
  createRoadmap();
}

/**
 * コラボ機会発見を実行
 */
function findCollabOpportunities() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "コラボ機会発見機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・相性の良いチャンネル発見\n" +
    "・コラボ企画提案\n" +
    "・パートナーシップ管理",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * KPI目標設定を実行
 */
function setKPITargets() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "KPI目標設定機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・具体的数値目標設定\n" +
    "・進捗追跡ダッシュボード\n" +
    "・目標達成度評価",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 改善提案生成を実行
 */
function generateImprovements() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "AI改善提案機能は現在開発中です。\n\n" +
    "近日実装予定:\n" +
    "・パフォーマンス分析に基づく提案\n" +
    "・成長施策の優先順位付け\n" +
    "・個別化された戦略提案",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * セルクリックイベントハンドラ
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // YouTube事業管理ダッシュボードでのクリックを処理
    if (sheet.getName() === "🎬 YouTube事業管理") {
      
      // 実行ボタン（F列）のクリック
      if (range.getColumn() === 6) {
        var row = range.getRow();
        var buttonValue = range.getValue();
        
        if (buttonValue === "▶ 実行") {
          // 隠し関数名を取得（G列）
          var functionName = sheet.getRange(row, 7).getValue();
          
          if (functionName) {
            try {
              // 関数を実行
              if (typeof eval(functionName) === 'function') {
                eval(functionName + '()');
              }
            } catch (error) {
              Logger.log("関数実行エラー: " + functionName + " - " + error.toString());
            }
          }
        }
      }
      
      // クイックアクセスボタン（28行目のB,D,F,Hセル）
      if (range.getRow() === 28) {
        var col = range.getColumn();
        if ([2, 4, 6, 8].indexOf(col) !== -1) {
          var functionName = sheet.getRange(29, col).getValue();
          
          if (functionName) {
            try {
              if (typeof eval(functionName) === 'function') {
                eval(functionName + '()');
              }
            } catch (error) {
              Logger.log("クイック関数実行エラー: " + functionName + " - " + error.toString());
            }
          }
        }
      }
    }
    
  } catch (error) {
    Logger.log("onEditエラー: " + error.toString());
  }
} 