/* eslint-disable */
/**
 * YouTube チャンネル分析統合ダッシュボード
 * 全機能へのアクセスポイントと統計表示
 *
 * 作成者: Claude AI
 * バージョン: 1.0
 * 最終更新: 2025-01-22
 */
/* eslint-enable */

/**
 * メインダッシュボードを作成または表示する
 */
function createOrShowMainDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getSheetByName("📊 統合ダッシュボード");
    
    // ダッシュボードが存在しない場合は作成
    if (!dashboardSheet) {
      createIntegratedDashboard();
    } else {
      // 既にある場合はアクティブにして統計更新
      ss.setActiveSheet(dashboardSheet);
      updateDashboardStatistics();
    }
  } catch (error) {
    Logger.log("ダッシュボード表示エラー: " + error.toString());
  }
}

/**
 * 統合ダッシュボードを作成
 */
function createIntegratedDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    // 既存のダッシュボードがあれば削除
    var existingDashboard = ss.getSheetByName("📊 統合ダッシュボード");
    if (existingDashboard) {
      ss.deleteSheet(existingDashboard);
    }
    
    // 新しいダッシュボードシートを作成（最初に配置）
    var dashboard = ss.insertSheet("📊 統合ダッシュボード", 0);
    
    // ========== ヘッダーセクション ==========
    dashboard.getRange("A1").setValue("YouTube チャンネル分析統合ダッシュボード");
    dashboard.getRange("A2").setValue("最終更新: " + new Date().toLocaleString());
    dashboard.getRange("A3").setValue("総合分析プラットフォーム - 全機能統括管理");
    
    // ========== クイックアクセスメニュー ==========
    dashboard.getRange("A5").setValue("🚀 クイックアクセス");
    
    var quickActions = [
      ["⚡", "API設定", "YouTube Data API キーの設定とテスト", "setApiKey"],
      ["📈", "チャンネル情報取得", "ハンドル名からチャンネルデータを一括取得", "processHandles"],
      ["📊", "ベンチマークレポート", "競合分析レポートを自動生成", "createBenchmarkReport"]
    ];
    
    var quickStartRow = 7;
    for (var i = 0; i < quickActions.length; i++) {
      var row = quickStartRow + (i * 2);
      dashboard.getRange(row, 1).setValue(quickActions[i][0]);
      dashboard.getRange(row, 2).setValue(quickActions[i][1]);
      dashboard.getRange(row + 1, 2).setValue(quickActions[i][2]);
      dashboard.getRange(row, 4).setValue("▶ 実行");
      dashboard.getRange(row, 4).setBackground("#34A853").setFontColor("white");
      dashboard.getRange(row, 5).setValue(quickActions[i][3]); // 関数名を保存
    }
    
    // ========== メイン機能メニュー ==========
    dashboard.getRange("A14").setValue("📋 主要機能");
    
    var mainFeatures = [
      ["📊", "既存チャンネル分析", "個別チャンネルの詳細パフォーマンス分析", "analyzeExistingChannel"],
      ["🔍", "ベンチマーク分析", "競合チャンネルとの比較分析ダッシュボード", "showBenchmarkDashboard"],
      ["🗺️", "ロードマップ策定", "チャンネル成長戦略とマイルストーン設定", "createRoadmap"],
      ["🔬", "市場リサーチ", "トレンド分析とニッチ市場調査", "conductMarketResearch"]
    ];
    
    var mainStartRow = 16;
    for (var i = 0; i < mainFeatures.length; i++) {
      var row = mainStartRow + (i * 3);
      dashboard.getRange(row, 1).setValue(mainFeatures[i][0]);
      dashboard.getRange(row, 2).setValue(mainFeatures[i][1]);
      dashboard.getRange(row + 1, 2).setValue(mainFeatures[i][2]);
      dashboard.getRange(row, 4).setValue("▶ 開始");
      dashboard.getRange(row, 4).setBackground("#4285F4").setFontColor("white");
      dashboard.getRange(row, 5).setValue(mainFeatures[i][3]); // 関数名を保存
    }
    
    // ========== リアルタイム統計セクション ==========
    dashboard.getRange("G5").setValue("📈 リアルタイム統計");
    
    // 主要指標
    var mainMetrics = [
      ["分析済みチャンネル数", "0", "", "チャンネル"],
      ["平均登録者数", "-", "", "人"],
      ["平均総視聴回数", "-", "", "回"],
      ["平均エンゲージメント率", "-", "", "%"]
    ];
    
    var metricsStartRow = 7;
    for (var i = 0; i < mainMetrics.length; i++) {
      dashboard.getRange(metricsStartRow + i, 7).setValue(mainMetrics[i][0]);
      dashboard.getRange(metricsStartRow + i, 9).setValue(mainMetrics[i][1]);
      dashboard.getRange(metricsStartRow + i, 10).setValue(mainMetrics[i][2]); // ステータス
      dashboard.getRange(metricsStartRow + i, 11).setValue(mainMetrics[i][3]); // 単位
    }
    
    // ========== 標準指標との比較 ==========
    dashboard.getRange("G12").setValue("📊 業界標準指標（2024年基準）");
    
    var benchmarkInfo = [
      ["エンゲージメント率", "小規模: 5%+, 中規模: 3%+, 大規模: 1%+"],
      ["月間投稿頻度", "最適: 4本, 最低: 1本"],
      ["視聴回数/登録者比", "良好: 10%+, 優秀: 20%+（1動画あたり）"],
      ["年間成長率", "健全: 5%+（登録者数）"]
    ];
    
    var benchmarkStartRow = 14;
    for (var i = 0; i < benchmarkInfo.length; i++) {
      dashboard.getRange(benchmarkStartRow + i, 7).setValue(benchmarkInfo[i][0]);
      dashboard.getRange(benchmarkStartRow + i, 9, 1, 3).merge();
      dashboard.getRange(benchmarkStartRow + i, 9).setValue(benchmarkInfo[i][1]);
    }
    
    // ========== パフォーマンスサマリー ==========
    dashboard.getRange("G19").setValue("🎯 パフォーマンスサマリー");
    
    dashboard.getRange("G21").setValue("総合スコア:");
    dashboard.getRange("I21").setValue("-");
    dashboard.getRange("J21").setValue("/ 100");
    
    dashboard.getRange("G22").setValue("改善ポイント:");
    dashboard.getRange("I22").setValue("データ収集中...");
    
    dashboard.getRange("G23").setValue("推奨アクション:");
    dashboard.getRange("I23").setValue("チャンネル分析を開始してください");
    
    // ========== データ品質指標 ==========
    dashboard.getRange("G25").setValue("📋 データ品質");
    
    dashboard.getRange("G27").setValue("API接続状況:");
    dashboard.getRange("I27").setValue("未確認");
    
    dashboard.getRange("G28").setValue("データ更新頻度:");
    dashboard.getRange("I28").setValue("-");
    
    dashboard.getRange("G29").setValue("分析対象期間:");
    dashboard.getRange("I29").setValue("-");
    
    // フォーマット設定
    formatIntegratedDashboard(dashboard);
    
    // 統計を初期化
    updateDashboardStatistics();
    
    // アクティブシートに設定
    ss.setActiveSheet(dashboard);
    
    // 初回作成時のガイド
    var props = PropertiesService.getDocumentProperties();
    if (!props.getProperty("integratedDashboardCreated")) {
      props.setProperty("integratedDashboardCreated", "true");
      ui.alert(
        "統合ダッシュボードへようこそ！",
        "YouTube チャンネル分析の統合ダッシュボードが作成されました。\n\n" +
        "🚀 クイックアクセス: よく使う機能への高速アクセス\n" +
        "📋 主要機能: 全分析機能への統括アクセス\n" +
        "📈 リアルタイム統計: 現在の分析状況を一目で確認\n\n" +
        "まずは「API設定」から始めて、その後チャンネル分析をお試しください。",
        ui.ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log("統合ダッシュボード作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "エラー",
      "ダッシュボードの作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 統合ダッシュボードのフォーマット設定
 */
function formatIntegratedDashboard(sheet) {
  // ========== ヘッダーフォーマット ==========
  sheet.getRange("A1:K1").merge();
  sheet.getRange("A1").setFontSize(20).setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#1a73e8").setFontColor("white");
  
  sheet.getRange("A2:K2").merge();
  sheet.getRange("A2").setFontSize(10).setFontStyle("italic")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A3:K3").merge();
  sheet.getRange("A3").setFontSize(12)
    .setHorizontalAlignment("center")
    .setFontColor("#5f6368");
  
  // ========== セクションヘッダー ==========
  var sectionHeaders = ["A5", "A14", "G5", "G12", "G19", "G25"];
  for (var i = 0; i < sectionHeaders.length; i++) {
    sheet.getRange(sectionHeaders[i]).setFontSize(14).setFontWeight("bold")
      .setBackground("#f8f9fa");
  }
  
  // ========== クイックアクセスフォーマット ==========
  var quickRows = [7, 9, 11];
  for (var i = 0; i < quickRows.length; i++) {
    var row = quickRows[i];
    sheet.getRange(row, 1).setFontSize(18).setHorizontalAlignment("center");
    sheet.setRowHeight(row, 35);
    sheet.getRange(row, 2, 1, 2).merge();
    sheet.getRange(row, 2).setFontSize(12).setFontWeight("bold");
    sheet.getRange(row + 1, 2, 1, 2).merge();
    sheet.getRange(row + 1, 2).setFontSize(10).setFontColor("#5f6368");
    sheet.getRange(row, 4).setHorizontalAlignment("center")
      .setVerticalAlignment("middle").setFontWeight("bold");
  }
  
  // ========== メイン機能フォーマット ==========
  var mainRows = [16, 19, 22, 25];
  for (var i = 0; i < mainRows.length; i++) {
    var row = mainRows[i];
    sheet.getRange(row, 1).setFontSize(20).setHorizontalAlignment("center");
    sheet.setRowHeight(row, 40);
    sheet.getRange(row, 2, 1, 2).merge();
    sheet.getRange(row, 2).setFontSize(13).setFontWeight("bold");
    sheet.getRange(row + 1, 2, 1, 2).merge();
    sheet.getRange(row + 1, 2).setFontSize(11).setFontColor("#5f6368");
    sheet.getRange(row, 4).setHorizontalAlignment("center")
      .setVerticalAlignment("middle").setFontWeight("bold");
  }
  
  // ========== 統計セクションフォーマット ==========
  for (var i = 0; i < 4; i++) {
    sheet.getRange(7 + i, 7, 1, 2).merge();
    sheet.getRange(7 + i, 7).setFontWeight("bold");
    sheet.getRange(7 + i, 9).setHorizontalAlignment("center").setFontSize(14);
    sheet.getRange(7 + i, 10).setHorizontalAlignment("center").setFontSize(12);
  }
  
  // ========== 列幅調整 ==========
  sheet.setColumnWidth(1, 50);   // アイコン列
  sheet.setColumnWidth(2, 180);  // タイトル列
  sheet.setColumnWidth(3, 180);  // 説明列
  sheet.setColumnWidth(4, 80);   // ボタン列
  sheet.setColumnWidth(5, 1);    // 隠し列（関数名）
  sheet.setColumnWidth(6, 30);   // スペーサー
  sheet.setColumnWidth(7, 160);  // 統計ラベル
  sheet.setColumnWidth(8, 20);   // スペーサー
  sheet.setColumnWidth(9, 100);  // 統計値
  sheet.setColumnWidth(10, 80);  // ステータス
  sheet.setColumnWidth(11, 60);  // 単位
  
  // ========== 枠線設定 ==========
  sheet.getRange("A5:E13").setBorder(true, true, true, true, true, true);
  sheet.getRange("A14:E27").setBorder(true, true, true, true, true, true);
  sheet.getRange("G5:K11").setBorder(true, true, true, true, true, true);
  sheet.getRange("G12:K18").setBorder(true, true, true, true, true, true);
  sheet.getRange("G19:K24").setBorder(true, true, true, true, true, true);
  sheet.getRange("G25:K30").setBorder(true, true, true, true, true, true);
  
  // ========== 背景色設定 ==========
  sheet.getRange("A5:E5").setBackground("#e8f0fe");
  sheet.getRange("A14:E14").setBackground("#e8f0fe");
  sheet.getRange("G5:K5").setBackground("#e8f0fe");
  sheet.getRange("G12:K12").setBackground("#e8f0fe");
  sheet.getRange("G19:K19").setBackground("#e8f0fe");
  sheet.getRange("G25:K25").setBackground("#e8f0fe");
}

/**
 * ダッシュボード統計を更新
 */
function updateDashboardStatistics() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getSheetByName("📊 統合ダッシュボード");
    
    if (!dashboardSheet) {
      return; // ダッシュボードがなければ何もしない
    }
    
    // API接続状況チェック
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    dashboardSheet.getRange("I27").setValue(apiKey ? "✅ 接続済み" : "❌ 未設定");
    
    // 全てのシートから分析済みチャンネルを集計
    var sheets = ss.getSheets();
    var analyzedChannels = 0;
    var totalSubscribers = 0;
    var totalViews = 0;
    var validChannelCount = 0;
    var dataSheet = null;
    
    // データシートを特定
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      
      // 分析シートをカウント
      if (sheetName.startsWith("分析_")) {
        analyzedChannels++;
        
        // 統計データを収集
        try {
          var subscriberValue = sheets[i].getRange("C14").getValue();
          if (subscriberValue && subscriberValue !== "非公開") {
            var subscribers = parseInt(subscriberValue.toString().replace(/,/g, "") || 0);
            if (!isNaN(subscribers)) {
              totalSubscribers += subscribers;
              validChannelCount++;
            }
          }
          
          var viewValue = sheets[i].getRange("C15").getValue();
          if (viewValue) {
            var views = parseInt(viewValue.toString().replace(/,/g, "") || 0);
            if (!isNaN(views)) {
              totalViews += views;
            }
          }
        } catch (e) {
          // 個別シートのエラーは無視
        }
      }
      
      // メインデータシートを特定
      if (sheetName === "Sheet1" || sheetName === "シート1" || sheetName.indexOf("ベンチマーク") !== -1) {
        dataSheet = sheets[i];
      }
    }
    
    // ベンチマークデータからも統計を取得
    if (dataSheet) {
      try {
        var benchmarkCount = countValidBenchmarkChannels(dataSheet);
        if (benchmarkCount > 0) {
          analyzedChannels += benchmarkCount;
          
          // ベンチマークデータの統計も加算
          var benchmarkStats = calculateBenchmarkStats(dataSheet);
          if (benchmarkStats.validCount > 0) {
            totalSubscribers += benchmarkStats.totalSubscribers;
            totalViews += benchmarkStats.totalViews;
            validChannelCount += benchmarkStats.validCount;
          }
        }
      } catch (e) {
        // ベンチマークデータエラーは無視
      }
    }
    
    // ========== 統計値の更新 ==========
    dashboardSheet.getRange("I7").setValue(analyzedChannels);
    dashboardSheet.getRange("J7").setValue(""); // チャンネル数にはステータス不要
    
    if (validChannelCount > 0) {
      var avgSubscribers = Math.round(totalSubscribers / validChannelCount);
      var avgViews = Math.round(totalViews / validChannelCount);
      var avgEngagement = ((totalViews / validChannelCount) / (totalSubscribers / validChannelCount) * 100);
      
      dashboardSheet.getRange("I8").setValue(avgSubscribers.toLocaleString());
      dashboardSheet.getRange("I9").setValue(avgViews.toLocaleString());
      dashboardSheet.getRange("I10").setValue(avgEngagement.toFixed(2));
      
      // ========== ステータス判定 ==========
      // 登録者数ステータス
      var subscriberStatus = avgSubscribers >= 100000 ? "✅ 優秀" : 
                            avgSubscribers >= 10000 ? "✅ 良好" : 
                            avgSubscribers >= 1000 ? "⚠️ 成長中" : "❌ 要改善";
      dashboardSheet.getRange("J8").setValue(subscriberStatus);
      
      // 視聴回数ステータス
      var viewRatio = avgViews / avgSubscribers;
      var viewStatus = viewRatio >= 20 ? "✅ 優秀" : 
                      viewRatio >= 10 ? "✅ 良好" : 
                      viewRatio >= 5 ? "⚠️ 平均的" : "❌ 要改善";
      dashboardSheet.getRange("J9").setValue(viewStatus);
      
      // エンゲージメント率ステータス
      var engagementStatus = avgEngagement >= 5 ? "✅ 優秀" : 
                            avgEngagement >= 2 ? "✅ 良好" : 
                            avgEngagement >= 1 ? "⚠️ 平均的" : "❌ 要改善";
      dashboardSheet.getRange("J10").setValue(engagementStatus);
      
      // 色設定
      setStatusColors(dashboardSheet, "J8", subscriberStatus);
      setStatusColors(dashboardSheet, "J9", viewStatus);
      setStatusColors(dashboardSheet, "J10", engagementStatus);
      
      // ========== 総合スコア計算 ==========
      var totalScore = calculateOverallScore(subscriberStatus, viewStatus, engagementStatus);
      dashboardSheet.getRange("I21").setValue(totalScore);
      
      // 改善ポイントとアクション提案
      var recommendations = generateRecommendations(subscriberStatus, viewStatus, engagementStatus, analyzedChannels);
      dashboardSheet.getRange("I22").setValue(recommendations.improvement);
      dashboardSheet.getRange("I23").setValue(recommendations.action);
      
    } else {
      // データがない場合
      dashboardSheet.getRange("I8").setValue("-");
      dashboardSheet.getRange("I9").setValue("-");
      dashboardSheet.getRange("I10").setValue("-");
      dashboardSheet.getRange("J8:J10").clearContent();
      dashboardSheet.getRange("I21").setValue("-");
      dashboardSheet.getRange("I22").setValue("分析データなし");
      dashboardSheet.getRange("I23").setValue("チャンネル分析を開始してください");
    }
    
    // ========== データ品質情報更新 ==========
    dashboardSheet.getRange("I28").setValue(analyzedChannels > 0 ? "リアルタイム" : "-");
    dashboardSheet.getRange("I29").setValue(analyzedChannels > 0 ? "過去30日間" : "-");
    
    // 最終更新時刻を更新
    dashboardSheet.getRange("A2").setValue("最終更新: " + new Date().toLocaleString());
    
  } catch (error) {
    Logger.log("ダッシュボード統計更新エラー: " + error.toString());
  }
}

/**
 * ベンチマークシートの有効チャンネル数をカウント
 */
function countValidBenchmarkChannels(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    var count = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2] !== "チャンネルが見つかりません" && data[i][2].toString().trim() !== "") {
        count++;
      }
    }
    
    return count;
  } catch (error) {
    return 0;
  }
}

/**
 * ベンチマークデータの統計を計算
 */
function calculateBenchmarkStats(sheet) {
  try {
    var data = sheet.getDataRange().getValues();
    var totalSubscribers = 0;
    var totalViews = 0;
    var validCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2] !== "チャンネルが見つかりません") {
        // 登録者数
        if (data[i][4] && data[i][4] !== "非公開") {
          var subscribers = parseInt(data[i][4].toString().replace(/,/g, "") || 0);
          if (!isNaN(subscribers)) {
            totalSubscribers += subscribers;
            validCount++;
          }
        }
        
        // 視聴回数
        if (data[i][5]) {
          var views = parseInt(data[i][5].toString().replace(/,/g, "") || 0);
          if (!isNaN(views)) {
            totalViews += views;
          }
        }
      }
    }
    
    return {
      totalSubscribers: totalSubscribers,
      totalViews: totalViews,
      validCount: validCount
    };
  } catch (error) {
    return { totalSubscribers: 0, totalViews: 0, validCount: 0 };
  }
}

/**
 * ステータスに応じて色を設定
 */
function setStatusColors(sheet, cellAddress, status) {
  var cell = sheet.getRange(cellAddress);
  
  if (status.includes("✅")) {
    cell.setFontColor("#0F9D58"); // 緑
  } else if (status.includes("⚠️")) {
    cell.setFontColor("#F4B400"); // 黄
  } else if (status.includes("❌")) {
    cell.setFontColor("#DB4437"); // 赤
  }
}

/**
 * 総合スコアを計算
 */
function calculateOverallScore(subscriberStatus, viewStatus, engagementStatus) {
  var score = 0;
  
  // 登録者数スコア（30点満点）
  if (subscriberStatus.includes("優秀")) score += 30;
  else if (subscriberStatus.includes("良好")) score += 25;
  else if (subscriberStatus.includes("成長中")) score += 15;
  else score += 5;
  
  // 視聴回数スコア（35点満点）
  if (viewStatus.includes("優秀")) score += 35;
  else if (viewStatus.includes("良好")) score += 30;
  else if (viewStatus.includes("平均的")) score += 20;
  else score += 10;
  
  // エンゲージメント率スコア（35点満点）
  if (engagementStatus.includes("優秀")) score += 35;
  else if (engagementStatus.includes("良好")) score += 30;
  else if (engagementStatus.includes("平均的")) score += 20;
  else score += 10;
  
  return score;
}

/**
 * 改善提案を生成
 */
function generateRecommendations(subscriberStatus, viewStatus, engagementStatus, channelCount) {
  var improvements = [];
  var actions = [];
  
  if (subscriberStatus.includes("❌") || subscriberStatus.includes("⚠️")) {
    improvements.push("登録者数");
    actions.push("コンテンツ品質向上");
  }
  
  if (viewStatus.includes("❌") || viewStatus.includes("⚠️")) {
    improvements.push("視聴回数");
    actions.push("サムネイル・タイトル最適化");
  }
  
  if (engagementStatus.includes("❌") || engagementStatus.includes("⚠️")) {
    improvements.push("エンゲージメント");
    actions.push("視聴者とのコミュニケーション強化");
  }
  
  if (channelCount < 5) {
    improvements.push("分析対象数");
    actions.push("より多くのチャンネルを分析");
  }
  
  return {
    improvement: improvements.length > 0 ? improvements.join(", ") : "全体的に良好",
    action: actions.length > 0 ? actions[0] : "現状維持"
  };
}

/**
 * ダッシュボードのボタンクリック処理
 */
function onDashboardButtonClick(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // 統合ダッシュボードでのクリックのみ処理
  if (sheet.getName() === "📊 統合ダッシュボード") {
    // ボタン列（D列）のクリックを検出
    if (range.getColumn() === 4) {
      var row = range.getRow();
      var buttonValue = range.getValue();
      
      if (buttonValue === "▶ 実行" || buttonValue === "▶ 開始") {
        // 対応する関数名を取得（E列）
        var functionName = sheet.getRange(row, 5).getValue();
        
        if (functionName) {
          try {
            // 関数を実行
            if (typeof this[functionName] === 'function') {
              this[functionName]();
            } else {
              // 他のファイルの関数を呼び出す場合
              eval(functionName + '()');
            }
            
            // 実行後に統計を更新
            Utilities.sleep(1000); // 1秒待機
            updateDashboardStatistics();
            
          } catch (error) {
            Logger.log("関数実行エラー: " + functionName + " - " + error.toString());
            SpreadsheetApp.getUi().alert(
              "実行エラー",
              "機能の実行中にエラーが発生しました: " + error.toString(),
              SpreadsheetApp.getUi().ButtonSet.OK
            );
          }
        }
      }
    }
  }
} 