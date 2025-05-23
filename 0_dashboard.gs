/* eslint-disable */
/**
 * YouTube チャンネル分析ダッシュボード（統合版）
 * 確実に動作するメニューベース・セル変更ベースの操作
 *
 * 作成者: Claude AI
 * バージョン: 4.0 (統合・実用版)
 * 最終更新: 2025-01-22
 */
/* eslint-enable */

/**
 * 統合ダッシュボードのメイン起動
 */
function createOrShowMainDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getSheetByName("📊 YouTube チャンネル分析");
    
    if (!dashboardSheet) {
      createUnifiedDashboard();
    } else {
      ss.setActiveSheet(dashboardSheet);
      updateDashboardDisplay();
    }
  } catch (error) {
    Logger.log("ダッシュボード起動エラー: " + error.toString());
  }
}

/**
 * 統合ダッシュボードを作成
 */
function createUnifiedDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    // 既存の混乱するダッシュボードを削除
    var oldSheets = ["📊 チャンネル分析", "📊 統合ダッシュボード", "🎬 YouTube事業管理"];
    for (var i = 0; i < oldSheets.length; i++) {
      var oldSheet = ss.getSheetByName(oldSheets[i]);
      if (oldSheet) {
        ss.deleteSheet(oldSheet);
      }
    }
    
    // 新しい統合ダッシュボードシートを作成
    var dashboard = ss.insertSheet("📊 YouTube チャンネル分析", 0);
    
    // ========== ヘッダー ==========
    dashboard.getRange("A1:J1").merge();
    dashboard.getRange("A1").setValue("📊 YouTube チャンネル分析ダッシュボード（統合版）");
    
    dashboard.getRange("A2:J2").merge();
    dashboard.getRange("A2").setValue("メニューベース操作 - 確実に動作します");
    
    // ========== 操作方法説明 ==========
    dashboard.getRange("A4").setValue("🚀 操作方法:");
    dashboard.getRange("A5").setValue("1. メニューから「YouTube ツール」を選択");
    dashboard.getRange("A6").setValue("2. または下記のセルに入力して操作");
    
    // ========== チャンネル入力エリア ==========
    dashboard.getRange("A8").setValue("🔗 チャンネル入力:");
    dashboard.getRange("B8").setValue("チャンネルURL or @ハンドル");
    dashboard.getRange("B8").setBackground("#f0f8ff");
    
    dashboard.getRange("A9").setValue("📝 操作セル:");
    dashboard.getRange("B9").setValue("ここに「分析」と入力してEnter");
    dashboard.getRange("B9").setBackground("#fff0f0");
    
    // ========== API設定状況 ==========
    dashboard.getRange("A11").setValue("🔑 API設定状況:");
    dashboard.getRange("B11").setValue("確認中...");
    
    // ========== 分析結果エリア ==========
    dashboard.getRange("A13").setValue("📈 最新分析結果");
    
    var resultLabels = [
      "チャンネル名:", "登録者数:", "総視聴回数:", "動画数:", 
      "チャンネル開設日:", "平均視聴回数:", "エンゲージメント率:", "総合評価:"
    ];
    
    for (var i = 0; i < resultLabels.length; i++) {
      dashboard.getRange(15 + i, 1).setValue(resultLabels[i]);
      dashboard.getRange(15 + i, 2).setValue("未分析");
    }
    
    // ========== 利用可能な機能一覧 ==========
    dashboard.getRange("D13").setValue("⚡ 利用可能な機能");
    
    var features = [
      "① API設定・テスト",
      "② チャンネル情報取得", 
      "③ ベンチマークレポート作成",
      "④ 個別チャンネル分析",
      "⑤ ベンチマーク分析",
      "⑥ 使い方ガイド"
    ];
    
    for (var i = 0; i < features.length; i++) {
      dashboard.getRange(15 + i, 4).setValue(features[i]);
      dashboard.getRange(15 + i, 5).setValue("メニューから実行");
    }
    
    // ========== 改善提案エリア ==========
    dashboard.getRange("A24").setValue("💡 改善提案・次のアクション");
    dashboard.getRange("A25:J27").merge();
    dashboard.getRange("A25").setValue("まだ分析が実行されていません。\n\n📌 はじめ方:\n1. メニュー「YouTube ツール > ① API設定・テスト」\n2. B8セルにチャンネルURL入力\n3. B9セルに「分析」と入力してEnter");
    
    // ========== クイックアクション（セル変更ベース） ==========
    dashboard.getRange("A29").setValue("⚡ クイックアクション");
    dashboard.getRange("A30").setValue("操作");
    dashboard.getRange("B30").setValue("実行内容");
    dashboard.getRange("C30").setValue("入力値");
    
    var quickActions = [
      ["分析", "チャンネル分析を実行", "B9セルに「分析」"],
      ["API", "API設定を実行", "B9セルに「API」"],
      ["レポート", "ベンチマークレポート作成", "B9セルに「レポート」"],
      ["更新", "ダッシュボード更新", "B9セルに「更新」"]
    ];
    
    for (var i = 0; i < quickActions.length; i++) {
      dashboard.getRange(31 + i, 1).setValue(quickActions[i][0]);
      dashboard.getRange(31 + i, 2).setValue(quickActions[i][1]);
      dashboard.getRange(31 + i, 3).setValue(quickActions[i][2]);
    }
    
    // フォーマットを適用
    formatUnifiedDashboard(dashboard);
    
    // 初期データを読み込み
    updateDashboardDisplay();
    
    // アクティブにする
    ss.setActiveSheet(dashboard);
    
    // 使い方ガイド
    ui.alert(
      "📊 統合ダッシュボード完成！",
      "確実に動作する統合ダッシュボードが作成されました。\n\n" +
      "🚀 操作方法:\n" +
      "1. メニュー「YouTube ツール」から各機能を実行\n" +
      "2. B8セルにチャンネルURL入力\n" +
      "3. B9セルに「分析」と入力してEnterで即座に分析実行\n\n" +
      "これで確実に動作します！",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("統合ダッシュボード作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "エラー",
      "ダッシュボード作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 統合ダッシュボードのフォーマット設定
 */
function formatUnifiedDashboard(sheet) {
  // ========== ヘッダー ==========
  sheet.getRange("A1").setFontSize(20).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("white")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A2").setFontSize(12)
    .setBackground("#e8f0fe").setFontColor("#1a73e8")
    .setHorizontalAlignment("center");
  
  // ========== 操作方法説明 ==========
  sheet.getRange("A4").setFontSize(14).setFontWeight("bold")
    .setBackground("#f8f9fa");
  sheet.getRange("A5:A6").setFontSize(11).setFontColor("#5f6368");
  
  // ========== 入力エリア ==========
  sheet.getRange("A8").setFontWeight("bold").setFontSize(12);
  sheet.getRange("B8").setBorder(true, true, true, true, false, false, "#4285F4", SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.getRange("A9").setFontWeight("bold").setFontSize(12);
  sheet.getRange("B9").setBorder(true, true, true, true, false, false, "#DB4437", SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== API設定 ==========
  sheet.getRange("A11").setFontWeight("bold");
  
  // ========== 分析結果エリア ==========
  sheet.getRange("A13").setFontSize(14).setFontWeight("bold")
    .setBackground("#f8f9fa");
  sheet.getRange("D13").setFontSize(14).setFontWeight("bold")
    .setBackground("#f8f9fa");
  
  // 分析結果の枠線
  sheet.getRange("A15:B22").setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A15:A22").setBackground("#e3f2fd").setFontWeight("bold");
  
  // 機能一覧の枠線
  sheet.getRange("D15:E20").setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("D15:D20").setBackground("#fff8e1").setFontWeight("bold");
  
  // ========== 改善提案エリア ==========
  sheet.getRange("A24").setFontSize(14).setFontWeight("bold")
    .setBackground("#f8f9fa");
  sheet.getRange("A25").setBorder(true, true, true, true, false, false, "#dddddd", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground("#fff3e0").setVerticalAlignment("top");
  
  // ========== クイックアクション ==========
  sheet.getRange("A29").setFontSize(14).setFontWeight("bold")
    .setBackground("#f8f9fa");
  
  sheet.getRange("A30:C30").setFontWeight("bold").setBackground("#e0e0e0");
  sheet.getRange("A31:C34").setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  
  // ========== 列幅設定 ==========
  sheet.setColumnWidth(1, 150);  // ラベル
  sheet.setColumnWidth(2, 200);  // 値・入力
  sheet.setColumnWidth(3, 120);  // スペーサー
  sheet.setColumnWidth(4, 150);  // 機能名
  sheet.setColumnWidth(5, 150);  // 説明
  sheet.setColumnWidth(6, 50);   // 余白
  
  // ========== 行高設定 ==========
  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 25);
  sheet.setRowHeight(8, 30);
  sheet.setRowHeight(9, 30);
  sheet.setRowHeight(25, 80);
}

/**
 * ダッシュボード表示を更新
 */
function updateDashboardDisplay() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    
    if (!dashboard) return;
    
    // API設定状況確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    var apiStatus = apiKey ? "✅ 設定済み" : "❌ 未設定（メニューから設定）";
    dashboard.getRange("B11").setValue(apiStatus);
    dashboard.getRange("B11").setFontColor(apiKey ? "#0F9D58" : "#DB4437");
    
    // 最新の分析データがあるかチェック
    var sheets = ss.getSheets();
    var latestAnalysis = null;
    
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName.startsWith("分析_")) {
        latestAnalysis = sheets[i];
        break;
      }
    }
    
    if (latestAnalysis) {
      try {
        // 複数のセル位置をチェックして柔軟にデータを取得
        var channelName = getValueFromMultipleCells(latestAnalysis, ["C6", "C4", "B6", "B4"]);
        var subscribers = getValueFromMultipleCells(latestAnalysis, ["C14", "C15", "B14", "B15"]);
        var totalViews = getValueFromMultipleCells(latestAnalysis, ["C15", "C16", "B15", "B16"]);
        var videoCount = getValueFromMultipleCells(latestAnalysis, ["C16", "C17", "B16", "B17"]);
        var createdDate = getValueFromMultipleCells(latestAnalysis, ["C8", "C13", "B8", "B13"]);
        
        // データを表示
        dashboard.getRange("B15").setValue(channelName || "取得中...");
        dashboard.getRange("B16").setValue(subscribers || "取得中...");
        dashboard.getRange("B17").setValue(totalViews || "取得中...");
        dashboard.getRange("B18").setValue(videoCount || "取得中...");
        dashboard.getRange("B19").setValue(createdDate || "取得中...");
        
        // 数値計算（エラー処理強化）
        var subscriberNum = extractNumberSafely(subscribers);
        var viewsNum = extractNumberSafely(totalViews);
        var videosNum = extractNumberSafely(videoCount);
        
        if (subscriberNum > 0 && viewsNum > 0 && videosNum > 0) {
          var avgViews = Math.round(viewsNum / videosNum);
          var engagementRate = (avgViews / subscriberNum * 100);
          var subscriberRate = viewsNum > 0 ? (subscriberNum / viewsNum * 100) : 0;
          
          dashboard.getRange("B20").setValue(avgViews.toLocaleString() + " 回/動画");
          dashboard.getRange("B21").setValue(engagementRate.toFixed(2) + "%");
          
          // 総合評価を計算
          var overallRating = calculateOverallRating(subscriberNum, engagementRate, videosNum);
          dashboard.getRange("B22").setValue(overallRating.score + "/100 (" + overallRating.grade + ")");
          
          // チャンネル登録率を新しい行に追加
          dashboard.getRange("A23").setValue("チャンネル登録率:");
          dashboard.getRange("B23").setValue(subscriberRate.toFixed(4) + "%");
          
          // 改善提案を生成
          var suggestions = generateImprovementSuggestions(subscriberNum, engagementRate, videosNum);
          dashboard.getRange("A25").setValue(suggestions);
        }
        
      } catch (e) {
        Logger.log("分析データ表示エラー: " + e.toString());
        dashboard.getRange("A25").setValue(
          "分析データの表示中にエラーが発生しました。\n" +
          "メニューから「④ 個別チャンネル分析」を再実行してください。"
        );
      }
    } else {
      // 分析データがない場合の初期状態
      var labels = ["未分析", "未分析", "未分析", "未分析", "未分析", "未分析", "未分析", "未分析"];
      for (var i = 0; i < labels.length; i++) {
        dashboard.getRange(15 + i, 2).setValue(labels[i]);
      }
      
      dashboard.getRange("A25").setValue(
        "まだ分析が実行されていません。\n\n" +
        "📌 はじめ方:\n" +
        "1. メニュー「YouTube ツール > ① API設定・テスト」\n" +
        "2. B8セルにチャンネルURL入力\n" +
        "3. B9セルに「分析」と入力してEnter"
      );
    }
    
    // 最終更新時間を記録
    dashboard.getRange("A2").setValue(
      "メニューベース操作 - 確実に動作します | 最終更新: " + 
      new Date().toLocaleTimeString()
    );
    
  } catch (error) {
    Logger.log("ダッシュボード表示更新エラー: " + error.toString());
  }
}

/**
 * 複数のセルから値を取得するヘルパー関数
 */
function getValueFromMultipleCells(sheet, cellAddresses) {
  for (var i = 0; i < cellAddresses.length; i++) {
    try {
      var value = sheet.getRange(cellAddresses[i]).getValue();
      if (value && value.toString().trim() !== "") {
        return value;
      }
    } catch (e) {
      // セルが存在しない場合は次を試す
      continue;
    }
  }
  return null;
}

/**
 * 安全な数値抽出
 */
function extractNumberSafely(value) {
  try {
    if (!value) return 0;
    var numStr = value.toString().replace(/[,\s]/g, "");
    var num = parseInt(numStr);
    return isNaN(num) ? 0 : num;
  } catch (e) {
    return 0;
  }
}

/**
 * 総合評価を計算
 */
function calculateOverallRating(subscribers, engagementRate, videoCount) {
  var score = 0;
  var grade = "";
  
  // 登録者数スコア (30点満点)
  if (subscribers >= 100000) score += 30;
  else if (subscribers >= 10000) score += 25;
  else if (subscribers >= 1000) score += 20;
  else if (subscribers >= 100) score += 15;
  else score += 10;
  
  // エンゲージメント率スコア (40点満点)
  if (engagementRate >= 10) score += 40;
  else if (engagementRate >= 5) score += 35;
  else if (engagementRate >= 2) score += 25;
  else if (engagementRate >= 1) score += 15;
  else score += 5;
  
  // 動画数スコア (30点満点)
  if (videoCount >= 100) score += 30;
  else if (videoCount >= 50) score += 25;
  else if (videoCount >= 20) score += 20;
  else if (videoCount >= 10) score += 15;
  else score += 10;
  
  // グレード判定
  if (score >= 90) grade = "S級";
  else if (score >= 80) grade = "A級";
  else if (score >= 70) grade = "B級";
  else if (score >= 60) grade = "C級";
  else grade = "成長段階";
  
  return { score: score, grade: grade };
}

/**
 * 改善提案を生成
 */
function generateImprovementSuggestions(subscribers, engagementRate, videoCount) {
  var suggestions = ["💡 改善提案・次のアクション\n"];
  
  if (subscribers < 1000) {
    suggestions.push("🎯 最優先: 収益化条件の1000人達成に向けた施策");
  } else if (subscribers < 10000) {
    suggestions.push("📈 目標: 1万人達成で中規模チャンネルへ");
  } else if (subscribers < 100000) {
    suggestions.push("🌟 目標: 10万人達成で大規模チャンネルへ");
  }
  
  if (engagementRate < 1) {
    suggestions.push("⚡ 緊急: エンゲージメント率向上が必要（現在" + engagementRate.toFixed(2) + "%）");
  } else if (engagementRate < 3) {
    suggestions.push("📊 推奨: エンゲージメント率をさらに向上（目標3%以上）");
  }
  
  if (videoCount < 10) {
    suggestions.push("🎬 コンテンツ: 動画数を増やして認知度向上");
  } else if (videoCount > 200) {
    suggestions.push("📚 整理: プレイリスト整理で視聴しやすさ向上");
  }
  
  suggestions.push("\n📋 次のアクション:");
  suggestions.push("• メニューから「③ ベンチマークレポート作成」で詳細分析");
  suggestions.push("• 定期的に分析を実行してトレンド確認");
  
  return suggestions.join("\n");
}

/**
 * セル変更イベントハンドラ（B9セル専用）
 */
function handleQuickAction(command) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    
    if (!dashboard) return;
    
    command = command.toString().toLowerCase().trim();
    
    switch (command) {
      case "分析":
        executeChannelAnalysis();
        break;
      case "api":
        setApiKey();
        break;
      case "レポート":
        createBenchmarkReport();
        break;
      case "更新":
        updateDashboardDisplay();
        SpreadsheetApp.getUi().alert("ダッシュボードを更新しました");
        break;
      default:
        SpreadsheetApp.getUi().alert(
          "不明なコマンド",
          "利用可能なコマンド: 分析, API, レポート, 更新",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
    
    // コマンド実行後、セルをクリア
    dashboard.getRange("B9").setValue("ここに「分析」と入力してEnter");
    dashboard.getRange("B9").setBackground("#fff0f0");
    
  } catch (error) {
    Logger.log("クイックアクション実行エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "実行エラー",
      "コマンド実行中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * チャンネル分析を実行
 */
function executeChannelAnalysis() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    
    if (!dashboard) return;
    
    // チャンネル入力を取得
    var channelInput = dashboard.getRange("B8").getValue();
    
    if (!channelInput || channelInput.toString().trim() === "" || channelInput === "チャンネルURL or @ハンドル") {
      SpreadsheetApp.getUi().alert(
        "入力エラー",
        "B8セルにチャンネルURLまたは@ハンドル名を入力してください。\n\n例: https://www.youtube.com/@YouTube\nまたは: @YouTube",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // API設定確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      SpreadsheetApp.getUi().alert(
        "API設定が必要",
        "先にメニューから「① API設定・テスト」を実行してください。",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 分析実行表示
    dashboard.getRange("B15").setValue("分析実行中...");
    SpreadsheetApp.flush();
    
    // ハンドル名を抽出・正規化
    var handle = normalizeChannelInput(channelInput.toString());
    
    if (!handle) {
      SpreadsheetApp.getUi().alert(
        "入力形式エラー",
        "正しいYouTubeチャンネルURL または @ハンドル名を入力してください。",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // チャンネルIDを解決してB2セルに保存（4_channelCheck.gsとの互換性）
    try {
      // 統合ダッシュボード用の簡易チャンネルID解決
      var resolvedChannelId = resolveChannelIdForDashboard(handle, apiKey);
      if (resolvedChannelId) {
        dashboard.getRange("B2").setValue(resolvedChannelId); // CHANNEL_ID_CELL互換
        Logger.log("統合ダッシュボード: チャンネルID保存 " + resolvedChannelId);
      }
    } catch (idError) {
      Logger.log("チャンネルID解決エラー: " + idError.toString());
    }
    
    // 統合ダッシュボード専用の分析関数を実行
    try {
      var result = executeUnifiedChannelAnalysis(handle, apiKey);
      
      if (result.success) {
        // ダッシュボードに結果を反映
        dashboard.getRange("B15").setValue(result.channelName);
        dashboard.getRange("B16").setValue(result.subscribers.toLocaleString() + " 人");
        dashboard.getRange("B17").setValue(result.totalViews.toLocaleString() + " 回");
        dashboard.getRange("B18").setValue(result.videoCount.toLocaleString() + " 本");
        
        var avgViews = result.videoCount > 0 ? Math.round(result.totalViews / result.videoCount) : 0;
        var engagementRate = result.subscribers > 0 ? (avgViews / result.subscribers * 100) : 0;
        var subscriberRate = result.totalViews > 0 ? (result.subscribers / result.totalViews * 100) : 0;
        
        dashboard.getRange("B20").setValue(avgViews.toLocaleString() + " 回/動画");
        dashboard.getRange("B21").setValue(engagementRate.toFixed(2) + "%");
        dashboard.getRange("B22").setValue(result.score + "/100 (" + result.grade + ")");
        
        // チャンネル登録率を新しい行に追加
        dashboard.getRange("A23").setValue("チャンネル登録率:");
        dashboard.getRange("B23").setValue(subscriberRate.toFixed(4) + "%");
        
        // 改善提案を更新
        var suggestions = generateImprovementSuggestions(result.subscribers, engagementRate, result.videoCount);
        dashboard.getRange("A25").setValue(suggestions);
        
        SpreadsheetApp.getUi().alert(
          "✅ 分析完了",
          "チャンネル分析が完了しました！\n\n• 分析シート: " + result.sheetName + "\n• チャンネル名: " + result.channelName + "\n• 総合スコア: " + result.score + "/100 (" + result.grade + ")",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } catch (analysisError) {
      dashboard.getRange("B15").setValue("分析エラー");
      dashboard.getRange("A25").setValue("分析中にエラーが発生しました:\n" + analysisError.toString());
      
      SpreadsheetApp.getUi().alert(
        "分析エラー",
        "チャンネル分析中にエラーが発生しました:\n\n" + analysisError.toString() + "\n\nAPIキーやネットワーク接続を確認してください。",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log("チャンネル分析実行エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "実行エラー",
      "分析実行中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 統合ダッシュボード用のチャンネルID解決関数
 */
function resolveChannelIdForDashboard(handle, apiKey) {
  try {
    // @ハンドル形式をチャンネルIDに変換
    if (handle.startsWith("@")) {
      var searchUrl = "https://www.googleapis.com/youtube/v3/search?part=snippet&q=" + 
                     encodeURIComponent(handle) + "&type=channel&maxResults=1&key=" + apiKey;
      
      var response = UrlFetchApp.fetch(searchUrl);
      var data = JSON.parse(response.getContentText());
      
      if (data.items && data.items.length > 0) {
        return data.items[0].snippet.channelId;
      }
    }
    
    // 既にチャンネルID形式の場合
    if (handle.match(/^UC[\w-]{22}$/)) {
      return handle;
    }
    
    return null;
  } catch (e) {
    Logger.log("チャンネルID解決エラー: " + e.toString());
    return null;
  }
}

/**
 * チャンネル入力を正規化
 */
function normalizeChannelInput(input) {
  try {
    input = input.trim();
    
    // YouTubeのURL形式の場合
    if (input.includes("youtube.com")) {
      if (input.includes("/@")) {
        return "@" + input.split("/@")[1].split("/")[0];
      } else if (input.includes("/c/")) {
        return "@" + input.split("/c/")[1].split("/")[0];
      } else if (input.includes("/channel/")) {
        return input.split("/channel/")[1].split("/")[0];
      }
    }
    
    // @ハンドル形式の場合
    if (input.startsWith("@")) {
      return input;
    }
    
    // その他の場合は@を付加
    if (!input.startsWith("UC") && input.length > 2) {
      return "@" + input;
    }
    
    return input;
  } catch (e) {
    return null;
  }
}

/**
 * チャンネル分析ダッシュボードを作成
 */
function createChannelAnalysisDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    // 既存ダッシュボードがあれば削除
    var existingDashboard = ss.getSheetByName("📊 チャンネル分析");
    if (existingDashboard) {
      ss.deleteSheet(existingDashboard);
    }
    
    // 新しいダッシュボードシートを作成
    var dashboard = ss.insertSheet("📊 チャンネル分析", 0);
    
    // ========== ヘッダー ==========
    dashboard.getRange("A1:K1").merge();
    dashboard.getRange("A1").setValue("📊 YouTubeチャンネル分析ダッシュボード");
    
    dashboard.getRange("A2:K2").merge();
    dashboard.getRange("A2").setValue("チャンネルURLを入力して、詳細分析を実行できます");
    
    // ========== チャンネル入力エリア ==========
    dashboard.getRange("A4").setValue("🔗 チャンネルURL:");
    dashboard.getRange("B4:H4").merge();
    dashboard.getRange("B4").setValue("https://www.youtube.com/@channel_handle");
    
    dashboard.getRange("I4").setValue("🔍 基本分析");
    dashboard.getRange("I4").setBackground("#4285F4").setFontColor("white");
    
    // ========== API設定エリア ==========
    dashboard.getRange("A6").setValue("🔑 API設定:");
    dashboard.getRange("B6").setValue("設定");
    dashboard.getRange("B6").setBackground("#34A853").setFontColor("white");
    
    dashboard.getRange("D6").setValue("ステータス:");
    dashboard.getRange("E6").setValue("未確認");
    
    // ========== 分析結果エリア ==========
    dashboard.getRange("A8").setValue("📈 分析結果");
    
    // 基本情報
    var basicLabels = [
      "チャンネル名", "登録者数", "総視聴回数", "動画数", 
      "チャンネル開設日", "最新動画", "平均視聴回数", "エンゲージメント率"
    ];
    
    for (var i = 0; i < basicLabels.length; i++) {
      var row = 10 + i;
      dashboard.getRange(row, 1).setValue(basicLabels[i] + ":");
      dashboard.getRange(row, 3).setValue("未分析");
      dashboard.getRange(row, 5).setValue(""); // ステータス表示
    }
    
    // ========== パフォーマンス評価 ==========
    dashboard.getRange("G8").setValue("🏆 パフォーマンス評価");
    
    var performanceLabels = [
      "総合スコア", "登録者成長率", "視聴率", "投稿頻度", "エンゲージメント"
    ];
    
    for (var i = 0; i < performanceLabels.length; i++) {
      var row = 10 + i;
      dashboard.getRange(row, 7).setValue(performanceLabels[i] + ":");
      dashboard.getRange(row, 9).setValue("-");
      dashboard.getRange(row, 10).setValue(""); // 評価
    }
    
    // ========== 改善提案エリア ==========
    dashboard.getRange("A19").setValue("💡 改善提案・アクション");
    dashboard.getRange("A20:K22").merge();
    dashboard.getRange("A20").setValue("チャンネル分析を実行すると、AIによる具体的な改善提案が表示されます。");
    
    // ========== クイックアクション ==========
    dashboard.getRange("A24").setValue("⚡ クイックアクション");
    
    var quickActions = [
      ["📊", "詳細レポート", "createDetailedReport"],
      ["🔄", "データ更新", "refreshDashboard"], 
      ["📋", "レポート出力", "exportReport"],
      ["❓", "使い方ガイド", "showHelp"]
    ];
    
    for (var i = 0; i < quickActions.length; i++) {
      var col = 1 + (i * 2);
      dashboard.getRange(25, col).setValue(quickActions[i][0]);
      dashboard.getRange(25, col + 1).setValue(quickActions[i][1]);
      dashboard.getRange(25, col + 1).setBackground("#f8f9fa");
      dashboard.getRange(26, col + 1).setValue(quickActions[i][2]); // 隠し関数名
    }
    
    // フォーマットを適用
    formatChannelDashboard(dashboard);
    
    // 初期状態を設定
    refreshDashboard();
    
    // アクティブにする
    ss.setActiveSheet(dashboard);
    
    // 初回ガイド
    ui.alert(
      "📊 チャンネル分析ダッシュボード",
      "YouTubeチャンネル分析ダッシュボードが作成されました！\n\n" +
      "📌 使い方:\n" +
      "1. 「API設定」でYouTube Data API キーを設定\n" +
      "2. チャンネルURLを入力\n" +
      "3. 「基本分析」ボタンをクリック\n\n" +
      "シンプルで使いやすい設計になっています。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("チャンネル分析ダッシュボード作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "エラー",
      "ダッシュボード作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ダッシュボードのフォーマット設定
 */
function formatChannelDashboard(sheet) {
  // ========== ヘッダー ==========
  sheet.getRange("A1").setFontSize(22).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("white")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A2").setFontSize(12)
    .setBackground("#e8f0fe").setFontColor("#1a73e8")
    .setHorizontalAlignment("center");
  
  // ========== チャンネル入力エリア ==========
  sheet.getRange("A4").setFontWeight("bold").setFontSize(14);
  sheet.getRange("B4:H4").setBackground("#f8f9fa")
    .setBorder(true, true, true, true, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  
  sheet.getRange("I4").setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  
  // ========== API設定エリア ==========
  sheet.getRange("A6").setFontWeight("bold");
  sheet.getRange("B6").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange("D6").setFontWeight("bold");
  
  // ========== 分析結果エリア ==========
  sheet.getRange("A8").setFontSize(16).setFontWeight("bold")
    .setBackground("#f8f9fa");
  sheet.getRange("G8").setFontSize(16).setFontWeight("bold")
    .setBackground("#f8f9fa");
  
  // 基本情報エリアの枠線
  sheet.getRange("A10:E17")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("A10:A17").setBackground("#e3f2fd").setFontWeight("bold");
  
  // パフォーマンス評価エリアの枠線
  sheet.getRange("G10:K14")
    .setBorder(true, true, true, true, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("G10:G14").setBackground("#fff8e1").setFontWeight("bold");
  
  // ========== 改善提案エリア ==========
  sheet.getRange("A19").setFontSize(16).setFontWeight("bold")
    .setBackground("#f8f9fa");
  sheet.getRange("A20:K22")
    .setBorder(true, true, true, true, false, false, "#dddddd", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground("#fff3e0").setVerticalAlignment("top");
  
  // ========== クイックアクション ==========
  sheet.getRange("A24").setFontSize(16).setFontWeight("bold")
    .setBackground("#f8f9fa");
  
  // クイックアクションボタン
  for (var i = 0; i < 4; i++) {
    var col = 2 + (i * 2);
    sheet.getRange(25, col).setFontWeight("bold")
      .setBorder(true, true, true, true, false, false, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  // ========== 列幅設定 ==========
  sheet.setColumnWidth(1, 140);  // ラベル
  sheet.setColumnWidth(2, 120);  // 値
  sheet.setColumnWidth(3, 120);  // 値
  sheet.setColumnWidth(4, 80);   // スペーサー
  sheet.setColumnWidth(5, 100);  // ステータス
  sheet.setColumnWidth(6, 20);   // スペーサー
  sheet.setColumnWidth(7, 140);  // パフォーマンスラベル
  sheet.setColumnWidth(8, 20);   // スペーサー
  sheet.setColumnWidth(9, 100);  // パフォーマンス値
  sheet.setColumnWidth(10, 100); // 評価
  sheet.setColumnWidth(11, 50);  // 余白
  
  // ========== 行高設定 ==========
  sheet.setRowHeight(1, 50);
  sheet.setRowHeight(2, 25);
  sheet.setRowHeight(4, 40);
  sheet.setRowHeight(6, 35);
  sheet.setRowHeight(8, 35);
  sheet.setRowHeight(19, 35);
  sheet.setRowHeight(20, 60);
  sheet.setRowHeight(24, 35);
  sheet.setRowHeight(25, 35);
}

/**
 * ダッシュボードを更新
 */
function refreshDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 チャンネル分析");
    
    if (!dashboard) return;
    
    // API設定状況確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    dashboard.getRange("E6").setValue(apiKey ? "✅ 設定済み" : "❌ 未設定");
    dashboard.getRange("E6").setFontColor(apiKey ? "#0F9D58" : "#DB4437");
    
    // 最新の分析データがあるかチェック
    var sheets = ss.getSheets();
    var latestAnalysis = null;
    
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName.startsWith("分析_")) {
        latestAnalysis = sheets[i];
        break;
      }
    }
    
    if (latestAnalysis) {
      // 最新の分析データを表示
      try {
        var channelName = latestAnalysis.getRange("C6").getValue() || latestAnalysis.getRange("C4").getValue();
        var subscribers = latestAnalysis.getRange("C14").getValue();
        var totalViews = latestAnalysis.getRange("C15").getValue();
        var videoCount = latestAnalysis.getRange("C16").getValue();
        var createdDate = latestAnalysis.getRange("C8").getValue() || latestAnalysis.getRange("C13").getValue();
        
        dashboard.getRange("C10").setValue(channelName || "取得中...");
        dashboard.getRange("C11").setValue(subscribers || "取得中...");
        dashboard.getRange("C12").setValue(totalViews || "取得中...");
        dashboard.getRange("C13").setValue(videoCount || "取得中...");
        dashboard.getRange("C14").setValue(createdDate || "取得中...");
        
        // 簡単な評価を追加
        var subscriberNum = extractNumber(subscribers);
        var viewsNum = extractNumber(totalViews);
        var videosNum = extractNumber(videoCount);
        
        if (subscriberNum > 0 && viewsNum > 0 && videosNum > 0) {
          var avgViews = Math.round(viewsNum / videosNum);
          var engagementRate = (avgViews / subscriberNum * 100);
          var subscriberRate = viewsNum > 0 ? (subscriberNum / viewsNum * 100) : 0;
          
          dashboard.getRange("C15").setValue(avgViews.toLocaleString() + " 回/動画");
          dashboard.getRange("C16").setValue(engagementRate.toFixed(2) + "%");
          
          // 総合スコア算出（改善版）
          var subscriberScore = Math.min(30, Math.log10(subscriberNum) * 10);
          var viewScore = Math.min(40, Math.log10(avgViews) * 10);
          var engagementScore = Math.min(30, engagementRate * 5);
          var totalScore = Math.round(subscriberScore + viewScore + engagementScore);
          
          dashboard.getRange("G10").setValue(totalScore + " / 100");
          dashboard.getRange("J10").setValue(
            totalScore >= 80 ? "🟢 優秀" :
            totalScore >= 60 ? "🟡 良好" :
            totalScore >= 40 ? "🟠 普通" : "🔴 要改善"
          );
          
          // チャンネル登録率を新しい行に追加
          dashboard.getRange("A23").setValue("チャンネル登録率:");
          dashboard.getRange("B23").setValue(subscriberRate.toFixed(4) + "%");
          
          // 改善提案を生成
          var suggestions = generateSimpleSuggestions(subscriberNum, engagementRate, videosNum);
          dashboard.getRange("A20").setValue(suggestions);
        }
        
      } catch (e) {
        Logger.log("分析データ表示エラー: " + e.toString());
        dashboard.getRange("A20").setValue(
          "分析データの読み込み中にエラーが発生しました。\n" +
          "再度「基本分析」を実行してください。"
        );
      }
    } else {
      // 分析データがない場合の初期化
      var defaultLabels = ["未分析", "未分析", "未分析", "未分析", "未分析", "未分析", "未分析"];
      for (var i = 0; i < defaultLabels.length; i++) {
        dashboard.getRange(10 + i, 3).setValue(defaultLabels[i]);
      }
      
      dashboard.getRange("A20").setValue(
        "まだチャンネル分析が実行されていません。\n\n" +
        "📌 分析を開始するには:\n" +
        "1. 上記のチャンネルURLを入力\n" +
        "2. 「基本分析」ボタンをクリック"
      );
    }
    
    // 最終更新時間を記録
    dashboard.getRange("A2").setValue(
      "チャンネルURLを入力して、詳細分析を実行できます | 最終更新: " + 
      new Date().toLocaleTimeString()
    );
    
  } catch (error) {
    Logger.log("ダッシュボード更新エラー: " + error.toString());
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
 * 簡単な改善提案を生成
 */
function generateSimpleSuggestions(subscribers, engagementRate, videoCount) {
  var suggestions = [];
  
  if (subscribers < 1000) {
    suggestions.push("🎯 収益化条件達成に向けて登録者1000人を目指しましょう");
  }
  
  if (engagementRate < 2) {
    suggestions.push("📈 エンゲージメント率向上のためコミュニケーションを活発化");
  }
  
  if (videoCount < 10) {
    suggestions.push("🎬 定期的な投稿でコンテンツ数を増やしましょう");
  }
  
  if (engagementRate > 5) {
    suggestions.push("🌟 高いエンゲージメント率です！この調子でコンテンツ品質を維持しましょう");
  }
  
  return suggestions.length > 0 ? 
    "💡 改善提案:\n" + suggestions.join("\n") :
    "🎉 バランスの取れた良いチャンネルです！継続的な改善を心がけましょう。";
}

/**
 * 基本分析を実行（修正版）
 */
function runBasicAnalysis() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 チャンネル分析");
    
    if (!dashboard) {
      SpreadsheetApp.getUi().alert(
        "エラー", 
        "ダッシュボードが見つかりません", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // チャンネルURLを取得
    var channelUrl = dashboard.getRange("B4").getValue();
    
    if (!channelUrl || !channelUrl.toString().includes("youtube.com")) {
      SpreadsheetApp.getUi().alert(
        "入力エラー", 
        "正しいYouTubeチャンネルURLを入力してください。\n\n例: https://www.youtube.com/@channel_handle", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // API設定確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      SpreadsheetApp.getUi().alert(
        "API設定が必要", 
        "先に「API設定」ボタンでYouTube Data API キーを設定してください。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 分析実行前の表示更新
    dashboard.getRange("C10").setValue("分析中...");
    dashboard.getRange("A20").setValue("チャンネルデータを取得中です。しばらくお待ちください...");
    SpreadsheetApp.flush();
    
    // ハンドル名を抽出
    var handle = extractChannelHandle(channelUrl);
    if (!handle) {
      SpreadsheetApp.getUi().alert(
        "URL解析エラー", 
        "チャンネルURLからハンドル名を抽出できませんでした。\n正しい形式で入力してください。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // 統合ダッシュボード専用の分析関数を使用
    try {
      var result = executeUnifiedChannelAnalysis(handle, apiKey);
      
      if (result.success) {
        // ダッシュボードに結果を反映
        dashboard.getRange("C10").setValue(result.channelName);
        dashboard.getRange("C11").setValue(result.subscribers.toLocaleString() + " 人");
        dashboard.getRange("C12").setValue(result.totalViews.toLocaleString() + " 回");
        dashboard.getRange("C13").setValue(result.videoCount.toLocaleString() + " 本");
        
        var avgViews = result.videoCount > 0 ? Math.round(result.totalViews / result.videoCount) : 0;
        var engagementRate = result.subscribers > 0 ? (avgViews / result.subscribers * 100) : 0;
        var subscriberRate = result.totalViews > 0 ? (result.subscribers / result.totalViews * 100) : 0;
        
        dashboard.getRange("C15").setValue(avgViews.toLocaleString() + " 回/動画");
        dashboard.getRange("C16").setValue(engagementRate.toFixed(2) + "%");
        
        // パフォーマンス評価を更新
        dashboard.getRange("I10").setValue(result.score + "/100");
        dashboard.getRange("J10").setValue(
          result.score >= 80 ? "🟢 優秀" :
          result.score >= 60 ? "🟡 良好" :
          result.score >= 40 ? "🟠 普通" : "🔴 要改善"
        );
        
        // チャンネル登録率を新しい行に追加
        dashboard.getRange("A23").setValue("チャンネル登録率:");
        dashboard.getRange("B23").setValue(subscriberRate.toFixed(4) + "%");
        
        // 改善提案を生成
        var suggestions = generateSimpleSuggestions(result.subscribers, engagementRate, result.videoCount);
        dashboard.getRange("A20").setValue(suggestions);
        
        SpreadsheetApp.getUi().alert(
          "✅ 分析完了", 
          "チャンネル分析が完了しました！\n\n• チャンネル名: " + result.channelName + "\n• 総合スコア: " + result.score + "/100 (" + result.grade + ")\n• 分析シート: " + result.sheetName, 
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    } catch (analysisError) {
      dashboard.getRange("C10").setValue("分析失敗");
      dashboard.getRange("A20").setValue("分析中にエラーが発生しました: " + analysisError.toString());
      
      SpreadsheetApp.getUi().alert(
        "❌ 分析エラー", 
        "分析中にエラーが発生しました:\n\n" + analysisError.toString() + "\n\nAPIキーやネットワーク接続を確認してください。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    Logger.log("基本分析実行エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "❌ 分析エラー", 
      "分析中にエラーが発生しました:\n\n" + error.toString() + "\n\nAPIキー設定やネットワーク接続を確認してください。", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * チャンネルURLからハンドル名を抽出
 */
function extractChannelHandle(url) {
  try {
    if (url.includes("/@")) {
      var handle = url.split("/@")[1].split("/")[0];
      return "@" + handle;
    } else if (url.includes("/c/")) {
      return url.split("/c/")[1].split("/")[0];
    } else if (url.includes("/channel/")) {
      return url.split("/channel/")[1].split("/")[0];
    }
    return null;
  } catch (e) {
    return null;
  }
}

/**
 * 詳細レポート作成
 */
function createDetailedReport() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "詳細レポート機能は近日実装予定です。\n\n現在は基本分析の結果をご活用ください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * レポート出力
 */
function exportReport() {
  SpreadsheetApp.getUi().alert(
    "🚧 開発中機能",
    "レポート出力機能は近日実装予定です。\n\n現在はスプレッドシート内でデータをご確認ください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 使い方ガイド
 */
function showHelp() {
  SpreadsheetApp.getUi().alert(
    "📖 使い方ガイド",
    "YouTubeチャンネル分析ダッシュボードの使い方:\n\n" +
    "1️⃣ API設定\n" +
    "・「API設定」ボタンでYouTube Data API キーを設定\n\n" +
    "2️⃣ チャンネル分析\n" +
    "・チャンネルURLを入力欄に貼り付け\n" +
    "・「基本分析」ボタンをクリック\n\n" +
    "3️⃣ 結果確認\n" +
    "・基本情報とパフォーマンス評価を確認\n" +
    "・改善提案を参考にチャンネル運営\n\n" +
    "シンプルで使いやすい設計です！",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * スプレッドシート開始時に統合ダッシュボードをセットアップ
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // メインメニュー
  var menu = ui.createMenu("YouTube ツール");
  menu.addItem("🏠 統合ダッシュボード", "createOrShowMainDashboard");
  menu.addSeparator();
  menu.addItem("① API設定・テスト", "setApiKey");
  menu.addItem("② チャンネル情報取得", "processHandles");
  menu.addItem("③ ベンチマークレポート作成", "createBenchmarkReport");
  menu.addSeparator();
  menu.addItem("📊 個別チャンネル分析", "executeChannelAnalysis");
  menu.addItem("🔍 ベンチマーク分析", "showBenchmarkDashboard");
  menu.addSeparator();
  menu.addItem("シートテンプレート作成", "setupBasicSheet");
  menu.addItem("使い方ガイドを表示", "showHelpSheet");
  menu.addSeparator();
  menu.addItem("🔧 ダッシュボード再作成", "recreateDashboard");
  menu.addItem("🧪 入力検証テスト", "testInputValidation");
  menu.addToUi();
  
  // 統合ダッシュボードを作成または表示
  createOrShowMainDashboard();
}

/**
 * セル編集イベントハンドラー - 統合ダッシュボード専用
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var sheetName = sheet.getName();
    var value = range.getValue();
    var row = range.getRow();
    var col = range.getColumn();
    
    // 統合ダッシュボードでの処理のみ
    if (sheetName === "📊 YouTube チャンネル分析") {
      
      // B9セル（操作セル）でのコマンド入力処理
      if (row === 9 && col === 2) {
        if (value && value.toString().trim() !== "" && 
            value.toString().trim() !== "ここに「分析」と入力してEnter") {
          
          // 一旦セルをクリアして処理中表示
          range.setValue("処理中...");
          SpreadsheetApp.flush();
          
          // コマンドを実行
          handleQuickAction(value);
        }
        return;
      }
      
      // B8セル（チャンネル入力）のプレースホルダー処理
      if (row === 8 && col === 2) {
        if (value && value.toString().trim() !== "" && 
            value.toString().trim() !== "チャンネルURL or @ハンドル") {
          // 入力検証
          var normalizedInput = normalizeChannelInput(value.toString());
          if (normalizedInput) {
            // 有効な入力の場合、背景色を変更
            range.setBackground("#e8f5e8");
            range.setFontColor("#2e7d32");
          } else {
            // 無効な入力の場合、エラー色
            range.setBackground("#ffebee");
            range.setFontColor("#c62828");
            SpreadsheetApp.getUi().alert(
              "入力形式エラー",
              "以下のいずれかの形式で入力してください：\n\n" +
              "• @ハンドル（例: @YouTube）\n" +
              "• チャンネルURL（例: https://www.youtube.com/@YouTube）\n" +
              "• チャンネルID（例: UC-9-kyTW8ZkZNDHQJ6FgpwQ）",
              SpreadsheetApp.getUi().ButtonSet.OK
            );
          }
        }
        return;
      }
    }
    
  } catch (error) {
    Logger.log("onEdit エラー: " + error.toString());
  }
}

/**
 * チャンネル入力検証の改善版
 */
function normalizeChannelInput(input) {
  try {
    if (!input || typeof input !== 'string') return null;
    
    input = input.trim();
    if (input === "" || input === "チャンネルURL or @ハンドル") return null;
    
    // YouTube URL形式の場合
    if (input.includes("youtube.com")) {
      // @ハンドル形式のURL
      if (input.includes("/@")) {
        var handle = input.split("/@")[1].split("/")[0].split("?")[0];
        return "@" + handle;
      }
      // /c/ 形式のURL
      else if (input.includes("/c/")) {
        var handle = input.split("/c/")[1].split("/")[0].split("?")[0];
        return "@" + handle;
      }
      // /channel/ 形式のURL (チャンネルID)
      else if (input.includes("/channel/")) {
        var channelId = input.split("/channel/")[1].split("/")[0].split("?")[0];
        if (channelId.startsWith("UC") && channelId.length === 24) {
          return channelId;
        }
      }
    }
    
    // @ハンドル形式の場合
    if (input.startsWith("@")) {
      var handle = input.substring(1);
      if (handle.length > 0 && /^[a-zA-Z0-9._-]+$/.test(handle)) {
        return input;
      }
    }
    
    // チャンネルIDの場合
    if (input.startsWith("UC") && input.length === 24 && /^[a-zA-Z0-9_-]+$/.test(input)) {
      return input;
    }
    
    // その他の場合は@を付加してハンドルとして扱う
    if (input.length > 0 && /^[a-zA-Z0-9._-]+$/.test(input)) {
      return "@" + input;
    }
    
    return null;
  } catch (e) {
    Logger.log("入力正規化エラー: " + e.toString());
    return null;
  }
}

/**
 * 統合ダッシュボード専用のチャンネル分析関数
 */
function executeUnifiedChannelAnalysis(handle, apiKey) {
  try {
    Logger.log("統合チャンネル分析開始: " + handle);
    
    // チャンネル情報を取得
    var channelInfo = getChannelByHandleUnified(handle, apiKey);
    
    if (!channelInfo) {
      return {
        success: false,
        error: "チャンネルが見つかりませんでした: " + handle
      };
    }
    
    var snippet = channelInfo.snippet;
    var statistics = channelInfo.statistics;
    
    // 基本データを取得
    var channelName = snippet.title;
    var subscribers = parseInt(statistics.subscriberCount || 0);
    var totalViews = parseInt(statistics.viewCount || 0);
    var videoCount = parseInt(statistics.videoCount || 0);
    var createdDate = snippet.publishedAt;
    
    // 分析シートを作成
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "分析_" + channelName.substring(0, 20).replace(/[/\\?*\[\]]/g, "");
    
    // 既存のシートがあれば削除
    var existingSheet = ss.getSheetByName(sheetName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // 新しい分析シートを作成
    var analysisSheet = ss.insertSheet(sheetName);
    
    // ヘッダー
    analysisSheet.getRange("A1").setValue("📊 チャンネル分析レポート（統合版）");
    analysisSheet.getRange("A2").setValue("分析日時: " + new Date().toLocaleString());
    analysisSheet.getRange("A3").setValue("対象チャンネル: " + channelName);
    
    // 基本情報セクション
    analysisSheet.getRange("A5").setValue("📋 基本情報");
    
    var basicInfo = [
      ["チャンネル名", channelName],
      ["ハンドル名", handle],
      ["チャンネルID", channelInfo.id],
      ["開設日", new Date(createdDate).toLocaleDateString()],
      ["説明", snippet.description ? snippet.description.substring(0, 200) + "..." : ""],
      ["国", snippet.country || "不明"]
    ];
    
    for (var i = 0; i < basicInfo.length; i++) {
      analysisSheet.getRange(6 + i, 1).setValue(basicInfo[i][0] + ":");
      analysisSheet.getRange(6 + i, 3).setValue(basicInfo[i][1]);
    }
    
    // パフォーマンス指標セクション
    analysisSheet.getRange("A13").setValue("📈 パフォーマンス指標");
    
    var avgViews = videoCount > 0 ? Math.round(totalViews / videoCount) : 0;
    var engagementRate = subscribers > 0 ? (avgViews / subscribers * 100) : 0;
    var subscriberRate = totalViews > 0 ? (subscribers / totalViews * 100) : 0;
    
    var metrics = [
      ["登録者数", subscribers.toLocaleString() + " 人"],
      ["総視聴回数", totalViews.toLocaleString() + " 回"],
      ["動画数", videoCount.toLocaleString() + " 本"],
      ["平均視聴回数/動画", avgViews.toLocaleString() + " 回"],
      ["エンゲージメント率", engagementRate.toFixed(2) + "%"],
      ["チャンネル登録率", subscriberRate.toFixed(4) + "%"]
    ];
    
    for (var i = 0; i < metrics.length; i++) {
      analysisSheet.getRange(14 + i, 1).setValue(metrics[i][0] + ":");
      analysisSheet.getRange(14 + i, 3).setValue(metrics[i][1]);
    }
    
    // 総合評価を計算
    var overallRating = calculateOverallRating(subscribers, engagementRate, videoCount);
    
    analysisSheet.getRange("A21").setValue("🏆 総合評価");
    analysisSheet.getRange("B22").setValue("スコア:");
    analysisSheet.getRange("C22").setValue(overallRating.score + "/100 (" + overallRating.grade + ")");
    
    // サムネイル画像
    if (snippet.thumbnails && snippet.thumbnails.high) {
      var imageFormula = '=IMAGE("' + snippet.thumbnails.high.url + '", 1)';
      analysisSheet.getRange("F1").setValue(imageFormula);
      analysisSheet.setRowHeight(1, 80);
    }
    
    // 改善提案
    analysisSheet.getRange("A24").setValue("💡 改善提案");
    var suggestions = generateImprovementSuggestions(subscribers, engagementRate, videoCount);
    analysisSheet.getRange("A25:F27").merge();
    analysisSheet.getRange("A25").setValue(suggestions);
    
    // フォーマット適用
    formatUnifiedAnalysisSheet(analysisSheet);
    
    // 結果を返す
    return {
      success: true,
      channelName: channelName,
      subscribers: subscribers,
      totalViews: totalViews,
      videoCount: videoCount,
      avgViews: avgViews,
      engagementRate: engagementRate,
      subscriberRate: subscriberRate,
      score: overallRating.score,
      grade: overallRating.grade,
      sheetName: sheetName
    };
    
  } catch (error) {
    Logger.log("統合チャンネル分析エラー: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 統合ダッシュボード用のチャンネル取得関数
 */
function getChannelByHandleUnified(handle, apiKey) {
  try {
    var username = handle.replace("@", "");
    var options = {
      method: "get",
      muteHttpExceptions: true,
    };

    // 検索APIを使用
    var searchUrl = "https://www.googleapis.com/youtube/v3/search?part=snippet&q=" + 
                   encodeURIComponent(handle) + "&type=channel&maxResults=5&key=" + apiKey;

    var searchResponse = UrlFetchApp.fetch(searchUrl, options);
    var searchData = JSON.parse(searchResponse.getContentText());

    if (searchData && searchData.items && searchData.items.length > 0) {
      var channelId = searchData.items[0].id.channelId;

      // チャンネルの詳細情報を取得
      var channelUrl = "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=" + 
                      channelId + "&key=" + apiKey;

      var channelResponse = UrlFetchApp.fetch(channelUrl, options);
      var channelData = JSON.parse(channelResponse.getContentText());

      if (channelData && channelData.items && channelData.items.length > 0) {
        return channelData.items[0];
      }
    }

    return null;
  } catch (error) {
    Logger.log("統合チャンネル取得エラー: " + error.toString());
    return null;
  }
}

/**
 * 統合分析シートのフォーマット
 */
function formatUnifiedAnalysisSheet(sheet) {
  // ヘッダー
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").setFontSize(18).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("white")
    .setHorizontalAlignment("center");
  
  sheet.getRange("A2").setFontStyle("italic");
  sheet.getRange("A3").setFontWeight("bold");
  
  // セクションヘッダー
  sheet.getRange("A5").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  sheet.getRange("A13").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  sheet.getRange("A21").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  sheet.getRange("A24").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  // 枠線
  sheet.getRange("A6:D11").setBorder(true, true, true, true, true, true);
  sheet.getRange("A14:D19").setBorder(true, true, true, true, true, true);
  sheet.getRange("A22:D22").setBorder(true, true, true, true, true, true);
  sheet.getRange("A25:F27").setBorder(true, true, true, true, false, false);
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 20);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 20);
  sheet.setColumnWidth(6, 150);
}

/**
 * 包括的一括分析システム - 既存機能統合版
 */
function executeComprehensiveAnalysis() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    var ui = SpreadsheetApp.getUi();
    
    if (!dashboard) {
      ui.alert("エラー", "統合ダッシュボードが見つかりません。", ui.ButtonSet.OK);
      return;
    }
    
    // チャンネル入力を確認
    var channelInput = dashboard.getRange("B8").getValue();
    
    if (!channelInput || channelInput.toString().trim() === "" || 
        channelInput === "チャンネルURL or @ハンドル" ||
        channelInput.toString().startsWith("例:")) {
      ui.alert(
        "入力エラー",
        "B8セルに分析対象のチャンネルを入力してください。\n\n" +
        "対応形式:\n" +
        "• @ハンドル名（例: @YouTube）\n" +
        "• チャンネルURL\n" +
        "• チャンネルID",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // API設定確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      ui.alert(
        "API設定が必要",
        "先にメニューから「① API設定・テスト」を実行してください。",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 詳細分析の確認
    var response = ui.alert(
      "包括的チャンネル分析",
      "以下の分析を実行します：\n\n" +
      "📊 基本分析: 登録者数、視聴回数、エンゲージメント率\n" +
      "📈 動画別分析: パフォーマンス、トレンド分析\n" +
      "👥 視聴者分析: 年齢層、地域、デバイス別\n" +
      "💡 AI改善提案: データに基づく具体的施策\n" +
      "📋 ギャップ分析: 業界基準との比較\n" +
      "🎯 アクションプラン: 優先度付き改善計画\n\n" +
      "この処理には2-5分かかります。実行しますか？",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // プログレス表示
    dashboard.getRange("B15").setValue("包括的分析実行中...");
    dashboard.getRange("A25").setValue("分析を開始しています。しばらくお待ちください...");
    SpreadsheetApp.flush();
    
    // チャンネル情報を解決してB8セルに正規化
    var normalizedInput = normalizeChannelInput(channelInput.toString());
    if (normalizedInput) {
      dashboard.getRange("B8").setValue(normalizedInput);
      SpreadsheetApp.flush();
    }
    
    // ステップ1: 基本統合分析
    dashboard.getRange("A25").setValue("ステップ 1/6: 基本チャンネル分析実行中...");
    SpreadsheetApp.flush();
    
    var basicResult = executeUnifiedChannelAnalysis(normalizedInput, apiKey);
    if (!basicResult.success) {
      throw new Error("基本分析に失敗: " + basicResult.error);
    }
    
    // ダッシュボードに基本結果を反映
    updateDashboardWithResults(dashboard, basicResult);
    
    // ステップ2: 既存の高度分析機能を活用
    dashboard.getRange("A25").setValue("ステップ 2/6: 詳細分析モジュール実行中...");
    SpreadsheetApp.flush();
    
    // 4_channelCheck.gsの高度機能を呼び出し
    var advancedResults = executeAdvancedAnalysisModules(normalizedInput, apiKey);
    
    // ステップ3: ギャップ分析の実行
    dashboard.getRange("A25").setValue("ステップ 3/6: 業界基準ギャップ分析中...");
    SpreadsheetApp.flush();
    
    var gapAnalysis = performGapAnalysis(basicResult, advancedResults);
    
    // ステップ4: 競合比較分析
    dashboard.getRange("A25").setValue("ステップ 4/6: 競合比較分析中...");
    SpreadsheetApp.flush();
    
    var competitorAnalysis = performCompetitorAnalysis(basicResult);
    
    // ステップ5: AI改善提案の強化
    dashboard.getRange("A25").setValue("ステップ 5/6: AI改善提案生成中...");
    SpreadsheetApp.flush();
    
    var aiRecommendations = generateEnhancedAIRecommendations(basicResult, gapAnalysis, competitorAnalysis);
    
    // ステップ6: 包括的レポート作成
    dashboard.getRange("A25").setValue("ステップ 6/6: 包括的レポート作成中...");
    SpreadsheetApp.flush();
    
    var comprehensiveReport = createComprehensiveAnalysisReport(
      basicResult, advancedResults, gapAnalysis, competitorAnalysis, aiRecommendations
    );
    
    // 最終更新
    updateDashboardWithComprehensiveResults(dashboard, comprehensiveReport);
    
    // 完了通知
    ui.alert(
      "✅ 包括的分析完了",
      "全ての分析が完了しました！\n\n" +
      "📊 作成されたレポート:\n" +
      "• " + comprehensiveReport.summarySheetName + "\n" +
      "• " + comprehensiveReport.gapAnalysisSheetName + "\n" +
      "• " + comprehensiveReport.actionPlanSheetName + "\n\n" +
      "🎯 重要な発見:\n" +
      "• 総合スコア: " + comprehensiveReport.overallScore + "/100\n" +
      "• 最優先改善項目: " + comprehensiveReport.topPriority + "\n" +
      "• 成長ポテンシャル: " + comprehensiveReport.growthPotential + "\n\n" +
      "詳細は各シートをご確認ください。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("包括的分析エラー: " + error.toString());
    var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("📊 YouTube チャンネル分析");
    if (dashboard) {
      dashboard.getRange("A25").setValue("分析中にエラーが発生しました: " + error.toString());
    }
    
    SpreadsheetApp.getUi().alert(
      "分析エラー",
      "包括的分析中にエラーが発生しました:\n\n" + error.toString() + "\n\nAPIキーやネットワーク接続を確認してください。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 既存の高度分析機能を統合実行
 */
function executeAdvancedAnalysisModules(channelInput, apiKey) {
  try {
    // 4_channelCheck.gsの高度機能を活用
    var results = {
      videoPerformance: null,
      audienceAnalysis: null,
      engagementAnalysis: null,
      trafficSources: null,
      hasOAuthAccess: false
    };
    
    // OAuth認証状態確認
    try {
      if (typeof getYouTubeOAuthService === 'function') {
        var service = getYouTubeOAuthService();
        results.hasOAuthAccess = service.hasAccess();
      }
    } catch (e) {
      Logger.log("OAuth確認エラー: " + e.toString());
    }
    
    // 基本的なチャンネル情報を取得
    var channelInfo = getChannelByHandleUnified(channelInput, apiKey);
    if (!channelInfo) {
      throw new Error("チャンネル情報の取得に失敗");
    }
    
    // 動画リストを取得して分析
    try {
      var videoList = getChannelVideos(channelInfo.id, apiKey, 50);
      results.videoPerformance = analyzeVideoPerformanceData(videoList);
    } catch (e) {
      Logger.log("動画分析エラー: " + e.toString());
      results.videoPerformance = { error: e.toString() };
    }
    
    // OAuth認証が利用可能な場合のみ高度分析実行
    if (results.hasOAuthAccess) {
      try {
        // 既存の分析関数を呼び出し（サイレントモード）
        if (typeof analyzeAudience === 'function') {
          results.audienceAnalysis = analyzeAudience(true);
        }
        if (typeof analyzeEngagement === 'function') {
          results.engagementAnalysis = analyzeEngagement(true);
        }
        if (typeof analyzeTrafficSources === 'function') {
          results.trafficSources = analyzeTrafficSources(true);
        }
      } catch (e) {
        Logger.log("OAuth分析エラー: " + e.toString());
      }
    }
    
    return results;
  } catch (error) {
    Logger.log("高度分析モジュールエラー: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * 業界基準とのギャップ分析
 */
function performGapAnalysis(basicResult, advancedResults) {
  try {
    var analysis = {
      subscriberGap: 0,
      engagementGap: 0,
      viewsGap: 0,
      contentGap: 0,
      overallGap: 0,
      benchmarks: {},
      recommendations: []
    };
    
    // 業界基準（チャンネルサイズ別）
    var benchmarks = getBenchmarksByChannelSize(basicResult.subscribers);
    analysis.benchmarks = benchmarks;
    
    // ギャップ計算
    analysis.subscriberGap = calculatePercentageGap(basicResult.subscribers, benchmarks.subscribers);
    analysis.engagementGap = calculatePercentageGap(basicResult.engagementRate, benchmarks.engagementRate);
    analysis.viewsGap = calculatePercentageGap(basicResult.avgViews, benchmarks.avgViews);
    analysis.contentGap = calculatePercentageGap(basicResult.videoCount, benchmarks.videoCount);
    
    // 総合ギャップ
    analysis.overallGap = Math.round((analysis.subscriberGap + analysis.engagementGap + analysis.viewsGap + analysis.contentGap) / 4);
    
    // ギャップに基づく推奨事項
    if (analysis.subscriberGap < -20) {
      analysis.recommendations.push({
        category: "登録者増加",
        priority: "高",
        action: "コンテンツ戦略の見直しと視聴者ターゲティング強化",
        gap: analysis.subscriberGap
      });
    }
    
    if (analysis.engagementGap < -15) {
      analysis.recommendations.push({
        category: "エンゲージメント向上", 
        priority: "高",
        action: "視聴者とのコミュニケーション活性化と質問・投票機能活用",
        gap: analysis.engagementGap
      });
    }
    
    if (analysis.viewsGap < -25) {
      analysis.recommendations.push({
        category: "視聴回数向上",
        priority: "中",
        action: "サムネイル・タイトル最適化とSEO強化",
        gap: analysis.viewsGap
      });
    }
    
    if (analysis.contentGap < -30) {
      analysis.recommendations.push({
        category: "コンテンツ量",
        priority: "中", 
        action: "投稿頻度の増加と一貫性のあるスケジュール",
        gap: analysis.contentGap
      });
    }
    
    return analysis;
  } catch (error) {
    Logger.log("ギャップ分析エラー: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * チャンネルサイズ別ベンチマーク取得
 */
function getBenchmarksByChannelSize(subscribers) {
  if (subscribers < 1000) {
    return {
      subscribers: 1000,
      engagementRate: 8.0,
      avgViews: 150,
      videoCount: 20,
      category: "新興チャンネル"
    };
  } else if (subscribers < 10000) {
    return {
      subscribers: 10000,
      engagementRate: 6.0,
      avgViews: 800,
      videoCount: 50,
      category: "成長チャンネル"
    };
  } else if (subscribers < 100000) {
    return {
      subscribers: 100000,
      engagementRate: 4.5,
      avgViews: 5000,
      videoCount: 100,
      category: "中規模チャンネル"
    };
  } else if (subscribers < 1000000) {
    return {
      subscribers: 1000000,
      engagementRate: 3.5,
      avgViews: 50000,
      videoCount: 200,
      category: "大規模チャンネル"
    };
  } else {
    return {
      subscribers: 5000000,
      engagementRate: 2.8,
      avgViews: 200000,
      videoCount: 500,
      category: "トップチャンネル"
    };
  }
}

/**
 * パーセンテージギャップ計算
 */
function calculatePercentageGap(current, benchmark) {
  if (benchmark === 0) return 0;
  return Math.round(((current - benchmark) / benchmark) * 100);
}

/**
 * 競合比較分析
 */
function performCompetitorAnalysis(basicResult) {
  try {
    var analysis = {
      competitorLevel: "",
      marketPosition: "",
      growthPotential: "",
      competitiveAdvantages: [],
      improvements: []
    };
    
    // 市場ポジション判定
    if (basicResult.subscribers < 10000) {
      analysis.competitorLevel = "新興・小規模";
      analysis.marketPosition = "ニッチ市場での成長段階";
      analysis.growthPotential = "高い（適切な戦略で急成長可能）";
    } else if (basicResult.subscribers < 100000) {
      analysis.competitorLevel = "中小規模";
      analysis.marketPosition = "地域・特定分野での認知度構築段階";
      analysis.growthPotential = "中程度（一貫した成長戦略が重要）";
    } else if (basicResult.subscribers < 1000000) {
      analysis.competitorLevel = "中規模";
      analysis.marketPosition = "業界内での地位確立段階";
      analysis.growthPotential = "安定（差別化と品質向上が鍵）";
    } else {
      analysis.competitorLevel = "大規模";
      analysis.marketPosition = "業界リーダー・影響力者";
      analysis.growthPotential = "維持重視（新規開拓と多角化）";
    }
    
    // 競合優位性分析
    if (basicResult.engagementRate > 5.0) {
      analysis.competitiveAdvantages.push("高いエンゲージメント率");
    }
    if (basicResult.subscriberRate > 0.01) {
      analysis.competitiveAdvantages.push("効率的な登録者獲得");
    }
    if (basicResult.videoCount > 100) {
      analysis.competitiveAdvantages.push("豊富なコンテンツライブラリ");
    }
    
    // 改善が必要な領域
    if (basicResult.engagementRate < 2.0) {
      analysis.improvements.push("エンゲージメント率向上が急務");
    }
    if (basicResult.avgViews < basicResult.subscribers * 0.1) {
      analysis.improvements.push("視聴率の改善が必要");
    }
    if (basicResult.videoCount < 20) {
      analysis.improvements.push("コンテンツ量の増加が必要");
    }
    
    return analysis;
  } catch (error) {
    Logger.log("競合分析エラー: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * 強化されたAI改善提案生成
 */
function generateEnhancedAIRecommendations(basicResult, gapAnalysis, competitorAnalysis) {
  try {
    var recommendations = {
      immediate: [],    // 即座に実行可能
      shortTerm: [],    // 1-3ヶ月
      longTerm: [],     // 3-12ヶ月
      strategic: []     // 戦略的・長期的
    };
    
    // 即座に実行可能な改善
    recommendations.immediate.push({
      title: "メタデータ最適化",
      description: "タイトル、説明文、タグの一括見直し",
      impact: "中",
      effort: "低",
      kpi: "検索流入+15-30%",
      action: "上位10動画のSEO最適化"
    });
    
    if (basicResult.engagementRate < 3.0) {
      recommendations.immediate.push({
        title: "エンゲージメント促進",
        description: "CTA（行動喚起）の強化と視聴者参加企画",
        impact: "高",
        effort: "低",
        kpi: "エンゲージメント率+50%",
        action: "動画終了時の具体的なアクション要請"
      });
    }
    
    // 短期的改善（1-3ヶ月）
    if (gapAnalysis.contentGap < -20) {
      recommendations.shortTerm.push({
        title: "コンテンツ制作スケジュール確立",
        description: "一貫した投稿スケジュールと品質基準の設定",
        impact: "高",
        effort: "中",
        kpi: "動画数+100%, 視聴維持率+20%",
        action: "週2-3本の定期投稿体制構築"
      });
    }
    
    if (basicResult.subscribers < 10000) {
      recommendations.shortTerm.push({
        title: "コミュニティ構築",
        description: "視聴者との双方向コミュニケーション強化",
        impact: "高",
        effort: "中",
        kpi: "登録率+200%, コメント数+150%",
        action: "ライブ配信とQ&Aセッションの定期開催"
      });
    }
    
    // 長期的改善（3-12ヶ月）
    recommendations.longTerm.push({
      title: "ブランディング強化",
      description: "一貫したビジュアルアイデンティティとメッセージング",
      impact: "高",
      effort: "高",
      kpi: "ブランド認知度+300%, 直接流入+100%",
      action: "チャンネルアート、サムネイル、ロゴの統一"
    });
    
    if (competitorAnalysis.marketPosition.includes("成長")) {
      recommendations.longTerm.push({
        title: "市場拡大戦略",
        description: "新しい視聴者層の開拓と多角化",
        impact: "高",
        effort: "高", 
        kpi: "新規視聴者+500%, 総視聴時間+300%",
        action: "関連分野への進出とコラボレーション"
      });
    }
    
    // 戦略的改善
    recommendations.strategic.push({
      title: "データドリブン運営",
      description: "Analytics活用による科学的チャンネル成長",
      impact: "最高",
      effort: "中",
      kpi: "全指標の継続的改善",
      action: "月次分析レポートと改善PDCAサイクル確立"
    });
    
    if (basicResult.subscribers > 50000) {
      recommendations.strategic.push({
        title: "収益化最適化",
        description: "多角的収益源の開発と最適化",
        impact: "最高",
        effort: "高",
        kpi: "収益性+200%, 持続可能性向上",
        action: "グッズ、メンバーシップ、スポンサーシップ戦略"
      });
    }
    
    return recommendations;
  } catch (error) {
    Logger.log("AI推奨事項生成エラー: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * 包括的分析レポート作成
 */
function createComprehensiveAnalysisReport(basicResult, advancedResults, gapAnalysis, competitorAnalysis, aiRecommendations) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date();
    var channelName = basicResult.channelName.substring(0, 15).replace(/[/\\?*\[\]]/g, "");
    
    // 1. 分析サマリーシート作成
    var summarySheetName = "📊 分析サマリー_" + channelName;
    var summarySheet = createOrReplaceSheet(ss, summarySheetName);
    
    // サマリーシートのコンテンツ
    createAnalysisSummarySheet(summarySheet, basicResult, gapAnalysis, competitorAnalysis, timestamp);
    
    // 2. ギャップ分析詳細シート作成
    var gapSheetName = "📋 ギャップ分析_" + channelName;
    var gapSheet = createOrReplaceSheet(ss, gapSheetName);
    
    createGapAnalysisSheet(gapSheet, basicResult, gapAnalysis, timestamp);
    
    // 3. アクションプランシート作成
    var actionSheetName = "🎯 アクションプラン_" + channelName;
    var actionSheet = createOrReplaceSheet(ss, actionSheetName);
    
    createActionPlanSheet(actionSheet, aiRecommendations, gapAnalysis, timestamp);
    
    // 4. 総合評価計算
    var overallScore = calculateOverallScore(basicResult, gapAnalysis, competitorAnalysis);
    var topPriority = determineTopPriority(gapAnalysis, aiRecommendations);
    var growthPotential = competitorAnalysis.growthPotential || "中程度";
    
    return {
      summarySheetName: summarySheetName,
      gapAnalysisSheetName: gapSheetName,
      actionPlanSheetName: actionSheetName,
      overallScore: overallScore,
      topPriority: topPriority,
      growthPotential: growthPotential,
      timestamp: timestamp
    };
    
  } catch (error) {
    Logger.log("包括的レポート作成エラー: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * シートを作成または置換
 */
function createOrReplaceSheet(ss, sheetName) {
  var existingSheet = ss.getSheetByName(sheetName);
  if (existingSheet) {
    ss.deleteSheet(existingSheet);
  }
  return ss.insertSheet(sheetName);
}

/**
 * 分析サマリーシート作成
 */
function createAnalysisSummarySheet(sheet, basicResult, gapAnalysis, competitorAnalysis, timestamp) {
  // ヘッダー
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").setValue("📊 YouTube チャンネル包括分析レポート")
    .setFontSize(18).setFontWeight("bold")
    .setBackground("#1a73e8").setFontColor("white")
    .setHorizontalAlignment("center");
    
  // 基本情報セクション
  sheet.getRange("A3").setValue("📋 基本チャンネル情報").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  var basicInfo = [
    ["チャンネル名", basicResult.channelName],
    ["分析日時", timestamp.toLocaleString()],
    ["登録者数", basicResult.subscribers.toLocaleString() + " 人"],
    ["総視聴回数", basicResult.totalViews.toLocaleString() + " 回"],
    ["動画数", basicResult.videoCount.toLocaleString() + " 本"],
    ["平均視聴回数", basicResult.avgViews.toLocaleString() + " 回/動画"],
    ["エンゲージメント率", basicResult.engagementRate.toFixed(2) + "%"],
    ["チャンネル登録率", basicResult.subscriberRate.toFixed(4) + "%"]
  ];
  
  for (var i = 0; i < basicInfo.length; i++) {
    sheet.getRange(4 + i, 1).setValue(basicInfo[i][0] + ":");
    sheet.getRange(4 + i, 3).setValue(basicInfo[i][1]);
  }
  
  // パフォーマンス評価セクション
  sheet.getRange("E3").setValue("🏆 パフォーマンス評価").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  var performanceData = [
    ["総合スコア", basicResult.score + "/100 (" + basicResult.grade + ")"],
    ["市場ポジション", competitorAnalysis.marketPosition],
    ["成長ポテンシャル", competitorAnalysis.growthPotential],
    ["競合レベル", competitorAnalysis.competitorLevel],
    ["業界ギャップ", gapAnalysis.overallGap + "%"],
    ["ベンチマーク", gapAnalysis.benchmarks.category],
    ["改善優先度", gapAnalysis.recommendations.length > 0 ? gapAnalysis.recommendations[0].priority : "低"]
  ];
  
  for (var i = 0; i < performanceData.length; i++) {
    sheet.getRange(4 + i, 5).setValue(performanceData[i][0] + ":");
    sheet.getRange(4 + i, 7).setValue(performanceData[i][1]);
  }
  
  // ギャップ分析サマリー
  sheet.getRange("A13").setValue("📊 主要指標ギャップ分析").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  var gapData = [
    ["指標", "現在値", "業界基準", "ギャップ", "評価"],
    ["登録者数", basicResult.subscribers.toLocaleString(), gapAnalysis.benchmarks.subscribers.toLocaleString(), gapAnalysis.subscriberGap + "%", getGapEvaluation(gapAnalysis.subscriberGap)],
    ["エンゲージメント率", basicResult.engagementRate.toFixed(2) + "%", gapAnalysis.benchmarks.engagementRate + "%", gapAnalysis.engagementGap + "%", getGapEvaluation(gapAnalysis.engagementGap)],
    ["平均視聴回数", basicResult.avgViews.toLocaleString(), gapAnalysis.benchmarks.avgViews.toLocaleString(), gapAnalysis.viewsGap + "%", getGapEvaluation(gapAnalysis.viewsGap)],
    ["動画数", basicResult.videoCount.toLocaleString(), gapAnalysis.benchmarks.videoCount.toLocaleString(), gapAnalysis.contentGap + "%", getGapEvaluation(gapAnalysis.contentGap)]
  ];
  
  for (var i = 0; i < gapData.length; i++) {
    for (var j = 0; j < gapData[i].length; j++) {
      var cell = sheet.getRange(14 + i, 1 + j);
      cell.setValue(gapData[i][j]);
      if (i === 0) {
        cell.setFontWeight("bold").setBackground("#e8f0fe");
      } else if (j === 4) {
        // 評価列に色付け
        var evaluation = gapData[i][j];
        if (evaluation === "優秀") cell.setBackground("#d4edda").setFontColor("#155724");
        else if (evaluation === "良好") cell.setBackground("#fff3cd").setFontColor("#856404");
        else if (evaluation === "要改善") cell.setBackground("#f8d7da").setFontColor("#721c24");
      }
    }
  }
  
  // 競合優位性・改善点
  sheet.getRange("A20").setValue("💪 競合優位性").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  sheet.getRange("E20").setValue("⚠️ 改善が必要な領域").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  // 優位性の表示
  for (var i = 0; i < Math.min(competitorAnalysis.competitiveAdvantages.length, 5); i++) {
    sheet.getRange(21 + i, 1).setValue("✅ " + competitorAnalysis.competitiveAdvantages[i]);
  }
  
  // 改善点の表示
  for (var i = 0; i < Math.min(competitorAnalysis.improvements.length, 5); i++) {
    sheet.getRange(21 + i, 5).setValue("🔸 " + competitorAnalysis.improvements[i]);
  }
  
  // フォーマット設定
  formatAnalysisSummarySheet(sheet);
}

/**
 * ギャップ評価を文字列で返す
 */
function getGapEvaluation(gap) {
  if (gap >= 20) return "優秀";
  if (gap >= 0) return "良好";
  if (gap >= -20) return "平均";
  return "要改善";
}

/**
 * ギャップ分析シート作成
 */
function createGapAnalysisSheet(sheet, basicResult, gapAnalysis, timestamp) {
  // ヘッダー
  sheet.getRange("A1:F1").merge();
  sheet.getRange("A1").setValue("📋 詳細ギャップ分析レポート")
    .setFontSize(18).setFontWeight("bold")
    .setBackground("#ff9800").setFontColor("white")
    .setHorizontalAlignment("center");
    
  sheet.getRange("A2").setValue("分析日時: " + timestamp.toLocaleString());
  sheet.getRange("A3").setValue("対象チャンネル: " + basicResult.channelName);
  
  // ベンチマーク情報
  sheet.getRange("A5").setValue("🎯 適用ベンチマーク").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  var benchmarkInfo = [
    ["カテゴリ", gapAnalysis.benchmarks.category],
    ["登録者基準", gapAnalysis.benchmarks.subscribers.toLocaleString() + " 人"],
    ["エンゲージメント基準", gapAnalysis.benchmarks.engagementRate + "%"],
    ["視聴回数基準", gapAnalysis.benchmarks.avgViews.toLocaleString() + " 回"],
    ["動画数基準", gapAnalysis.benchmarks.videoCount + " 本"]
  ];
  
  for (var i = 0; i < benchmarkInfo.length; i++) {
    sheet.getRange(6 + i, 1).setValue(benchmarkInfo[i][0] + ":");
    sheet.getRange(6 + i, 3).setValue(benchmarkInfo[i][1]);
  }
  
  // 詳細ギャップ分析
  sheet.getRange("A12").setValue("📊 詳細ギャップ分析").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  var detailedGaps = [
    ["指標", "現在値", "目標値", "差分", "達成率", "優先度", "推奨アクション"],
    ["登録者数", basicResult.subscribers.toLocaleString(), gapAnalysis.benchmarks.subscribers.toLocaleString(), (gapAnalysis.benchmarks.subscribers - basicResult.subscribers).toLocaleString(), Math.min(100, Math.round((basicResult.subscribers / gapAnalysis.benchmarks.subscribers) * 100)) + "%", getPriority(gapAnalysis.subscriberGap), getActionForGap("subscribers", gapAnalysis.subscriberGap)],
    ["エンゲージメント率", basicResult.engagementRate.toFixed(2) + "%", gapAnalysis.benchmarks.engagementRate + "%", (gapAnalysis.benchmarks.engagementRate - basicResult.engagementRate).toFixed(2) + "%", Math.min(100, Math.round((basicResult.engagementRate / gapAnalysis.benchmarks.engagementRate) * 100)) + "%", getPriority(gapAnalysis.engagementGap), getActionForGap("engagement", gapAnalysis.engagementGap)],
    ["平均視聴回数", basicResult.avgViews.toLocaleString(), gapAnalysis.benchmarks.avgViews.toLocaleString(), (gapAnalysis.benchmarks.avgViews - basicResult.avgViews).toLocaleString(), Math.min(100, Math.round((basicResult.avgViews / gapAnalysis.benchmarks.avgViews) * 100)) + "%", getPriority(gapAnalysis.viewsGap), getActionForGap("views", gapAnalysis.viewsGap)],
    ["動画数", basicResult.videoCount.toLocaleString(), gapAnalysis.benchmarks.videoCount.toLocaleString(), (gapAnalysis.benchmarks.videoCount - basicResult.videoCount).toLocaleString(), Math.min(100, Math.round((basicResult.videoCount / gapAnalysis.benchmarks.videoCount) * 100)) + "%", getPriority(gapAnalysis.contentGap), getActionForGap("content", gapAnalysis.contentGap)]
  ];
  
  for (var i = 0; i < detailedGaps.length; i++) {
    for (var j = 0; j < detailedGaps[i].length; j++) {
      var cell = sheet.getRange(13 + i, 1 + j);
      cell.setValue(detailedGaps[i][j]);
      if (i === 0) {
        cell.setFontWeight("bold").setBackground("#e8f0fe");
      } else if (j === 5) {
        // 優先度列に色付け
        var priority = detailedGaps[i][j];
        if (priority === "高") cell.setBackground("#ffebee").setFontColor("#c62828");
        else if (priority === "中") cell.setBackground("#fff3e0").setFontColor("#f57c00");
        else cell.setBackground("#e8f5e8").setFontColor("#2e7d32");
      }
    }
  }
  
  // 改善提案セクション
  sheet.getRange("A19").setValue("💡 優先改善提案").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  for (var i = 0; i < Math.min(gapAnalysis.recommendations.length, 5); i++) {
    var rec = gapAnalysis.recommendations[i];
    sheet.getRange(20 + i, 1).setValue((i + 1) + ". " + rec.category);
    sheet.getRange(20 + i, 3).setValue("優先度: " + rec.priority);
    sheet.getRange(20 + i, 5).setValue(rec.action);
    
    // 優先度に応じた色付け
    var priorityCell = sheet.getRange(20 + i, 3);
    if (rec.priority === "高") priorityCell.setBackground("#ffebee").setFontColor("#c62828");
    else if (rec.priority === "中") priorityCell.setBackground("#fff3e0").setFontColor("#f57c00");
    else priorityCell.setBackground("#e8f5e8").setFontColor("#2e7d32");
  }
  
  // フォーマット設定
  formatGapAnalysisSheet(sheet);
}

/**
 * 優先度を返す
 */
function getPriority(gap) {
  if (gap < -30) return "高";
  if (gap < -10) return "中";
  return "低";
}

/**
 * ギャップに基づくアクション推奨
 */
function getActionForGap(category, gap) {
  if (gap >= 0) return "現状維持・さらなる向上";
  
  switch (category) {
    case "subscribers":
      return gap < -30 ? "コンテンツ戦略全面見直し" : "ターゲティング最適化";
    case "engagement":
      return gap < -20 ? "コミュニティ構築強化" : "インタラクション改善";
    case "views":
      return gap < -25 ? "SEO・サムネイル最適化" : "推奨アルゴリズム対策";
    case "content":
      return gap < -30 ? "投稿頻度大幅増加" : "一貫性のあるスケジュール";
    default:
      return "データ分析による最適化";
  }
}

/**
 * アクションプランシート作成
 */
function createActionPlanSheet(sheet, aiRecommendations, gapAnalysis, timestamp) {
  // ヘッダー
  sheet.getRange("A1:H1").merge();
  sheet.getRange("A1").setValue("🎯 チャンネル成長アクションプラン")
    .setFontSize(18).setFontWeight("bold")
    .setBackground("#4caf50").setFontColor("white")
    .setHorizontalAlignment("center");
    
  sheet.getRange("A2").setValue("作成日時: " + timestamp.toLocaleString());
  
  // 即座に実行可能な施策
  var currentRow = 4;
  currentRow = createRecommendationSection(sheet, "⚡ 即座に実行可能（今すぐ～1週間）", aiRecommendations.immediate, currentRow);
  
  // 短期施策
  currentRow += 2;
  currentRow = createRecommendationSection(sheet, "📈 短期施策（1-3ヶ月）", aiRecommendations.shortTerm, currentRow);
  
  // 長期施策
  currentRow += 2;
  currentRow = createRecommendationSection(sheet, "🚀 長期施策（3-12ヶ月）", aiRecommendations.longTerm, currentRow);
  
  // 戦略的施策
  currentRow += 2;
  currentRow = createRecommendationSection(sheet, "🎯 戦略的施策（継続的）", aiRecommendations.strategic, currentRow);
  
  // 優先度マトリックス
  currentRow += 3;
  createPriorityMatrix(sheet, aiRecommendations, currentRow);
  
  // フォーマット設定
  formatActionPlanSheet(sheet);
}

/**
 * 推奨事項セクション作成
 */
function createRecommendationSection(sheet, title, recommendations, startRow) {
  sheet.getRange(startRow, 1).setValue(title).setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  if (recommendations.length === 0) {
    sheet.getRange(startRow + 1, 1).setValue("該当する推奨事項がありません。");
    return startRow + 1;
  }
  
  // ヘッダー
  var headers = ["施策", "説明", "影響度", "工数", "KPI目標", "具体的アクション"];
  for (var j = 0; j < headers.length; j++) {
    sheet.getRange(startRow + 1, 1 + j).setValue(headers[j]).setFontWeight("bold").setBackground("#e8f0fe");
  }
  
  // 推奨事項データ
  for (var i = 0; i < recommendations.length; i++) {
    var rec = recommendations[i];
    var row = startRow + 2 + i;
    
    sheet.getRange(row, 1).setValue(rec.title);
    sheet.getRange(row, 2).setValue(rec.description);
    sheet.getRange(row, 3).setValue(rec.impact);
    sheet.getRange(row, 4).setValue(rec.effort);
    sheet.getRange(row, 5).setValue(rec.kpi);
    sheet.getRange(row, 6).setValue(rec.action);
    
    // 影響度による色付け
    var impactCell = sheet.getRange(row, 3);
    if (rec.impact === "最高" || rec.impact === "高") impactCell.setBackground("#e8f5e8").setFontColor("#2e7d32");
    else if (rec.impact === "中") impactCell.setBackground("#fff3e0").setFontColor("#f57c00");
    else impactCell.setBackground("#fafafa").setFontColor("#757575");
    
    // 工数による色付け
    var effortCell = sheet.getRange(row, 4);
    if (rec.effort === "低") effortCell.setBackground("#e8f5e8").setFontColor("#2e7d32");
    else if (rec.effort === "中") effortCell.setBackground("#fff3e0").setFontColor("#f57c00");
    else effortCell.setBackground("#ffebee").setFontColor("#c62828");
  }
  
  return startRow + 2 + recommendations.length;
}

/**
 * 優先度マトリックス作成
 */
function createPriorityMatrix(sheet, aiRecommendations, startRow) {
  sheet.getRange(startRow, 1).setValue("📊 影響度×工数マトリックス（実行優先度参考）").setFontSize(14).setFontWeight("bold").setBackground("#f8f9fa");
  
  var matrixData = [
    ["", "低工数", "中工数", "高工数"],
    ["高影響", "最優先", "優先", "検討"],
    ["中影響", "優先", "検討", "後回し"],
    ["低影響", "検討", "後回し", "不要"]
  ];
  
  for (var i = 0; i < matrixData.length; i++) {
    for (var j = 0; j < matrixData[i].length; j++) {
      var cell = sheet.getRange(startRow + 1 + i, 1 + j);
      cell.setValue(matrixData[i][j]);
      
      if (i === 0 || j === 0) {
        cell.setFontWeight("bold").setBackground("#e8f0fe");
      } else {
        // 優先度による色付け
        var value = matrixData[i][j];
        if (value === "最優先") cell.setBackground("#c8e6c9").setFontColor("#1b5e20");
        else if (value === "優先") cell.setBackground("#fff9c4").setFontColor("#f57f17");
        else if (value === "検討") cell.setBackground("#ffecb3").setFontColor("#ff8f00");
        else cell.setBackground("#ffcdd2").setFontColor("#c62828");
      }
    }
  }
}

/**
 * ダッシュボードに結果を反映
 */
function updateDashboardWithResults(dashboard, basicResult) {
  try {
    dashboard.getRange("B15").setValue(basicResult.channelName);
    dashboard.getRange("B16").setValue(basicResult.subscribers.toLocaleString() + " 人");
    dashboard.getRange("B17").setValue(basicResult.totalViews.toLocaleString() + " 回");
    dashboard.getRange("B18").setValue(basicResult.videoCount.toLocaleString() + " 本");
    dashboard.getRange("B20").setValue(basicResult.avgViews.toLocaleString() + " 回/動画");
    dashboard.getRange("B21").setValue(basicResult.engagementRate.toFixed(2) + "%");
    dashboard.getRange("B22").setValue(basicResult.score + "/100 (" + basicResult.grade + ")");
    dashboard.getRange("B23").setValue(basicResult.subscriberRate.toFixed(4) + "%");
  } catch (error) {
    Logger.log("ダッシュボード基本更新エラー: " + error.toString());
  }
}

/**
 * ダッシュボードに包括的結果を反映
 */
function updateDashboardWithComprehensiveResults(dashboard, comprehensiveReport) {
  try {
    var improvedSuggestions = 
      "🎯 包括的分析完了！\n\n" +
      "📊 作成レポート:\n" +
      "• " + comprehensiveReport.summarySheetName + "\n" +
      "• " + comprehensiveReport.gapAnalysisSheetName + "\n" +
      "• " + comprehensiveReport.actionPlanSheetName + "\n\n" +
      "🔍 重要指標:\n" +
      "• 総合スコア: " + comprehensiveReport.overallScore + "/100\n" +
      "• 最優先改善: " + comprehensiveReport.topPriority + "\n" +
      "• 成長ポテンシャル: " + comprehensiveReport.growthPotential + "\n\n" +
      "📋 次のステップ:\n" +
      "1. アクションプランシートで具体的施策を確認\n" +
      "2. 最優先項目から実行開始\n" +
      "3. 月次で進捗を測定・分析";
    
    dashboard.getRange("A25").setValue(improvedSuggestions);
  } catch (error) {
    Logger.log("ダッシュボード包括更新エラー: " + error.toString());
  }
}

/**
 * 総合スコア計算
 */
function calculateOverallScore(basicResult, gapAnalysis, competitorAnalysis) {
  try {
    var score = 0;
    
    // 基本スコア（既存の算出）
    score += basicResult.score * 0.4; // 40%
    
    // ギャップ分析スコア（30%）
    var gapScore = 0;
    if (gapAnalysis.overallGap >= 0) gapScore = 100;
    else if (gapAnalysis.overallGap >= -20) gapScore = 80;
    else if (gapAnalysis.overallGap >= -40) gapScore = 60;
    else gapScore = 40;
    score += gapScore * 0.3;
    
    // 競合分析スコア（30%）
    var competitorScore = 0;
    if (competitorAnalysis.competitiveAdvantages.length >= 3) competitorScore = 90;
    else if (competitorAnalysis.competitiveAdvantages.length >= 2) competitorScore = 75;
    else if (competitorAnalysis.competitiveAdvantages.length >= 1) competitorScore = 60;
    else competitorScore = 45;
    
    if (competitorAnalysis.improvements.length <= 1) competitorScore += 10;
    else if (competitorAnalysis.improvements.length >= 4) competitorScore -= 15;
    
    score += competitorScore * 0.3;
    
    return Math.round(score);
  } catch (error) {
    Logger.log("総合スコア計算エラー: " + error.toString());
    return basicResult.score || 50;
  }
}

/**
 * 最優先項目決定
 */
function determineTopPriority(gapAnalysis, aiRecommendations) {
  try {
    // ギャップ分析から最大のギャップを特定
    var gaps = [
      { name: "登録者数", gap: gapAnalysis.subscriberGap },
      { name: "エンゲージメント率", gap: gapAnalysis.engagementGap },
      { name: "視聴回数", gap: gapAnalysis.viewsGap },
      { name: "コンテンツ量", gap: gapAnalysis.contentGap }
    ];
    
    gaps.sort(function(a, b) { return a.gap - b.gap; });
    
    var worstGap = gaps[0];
    
    // 即座に実行可能な項目があれば優先
    if (aiRecommendations.immediate.length > 0) {
      return aiRecommendations.immediate[0].title + "（即座実行可能）";
    }
    
    // そうでなければギャップが最大の項目
    if (worstGap.gap < -30) {
      return worstGap.name + "の大幅改善";
    } else if (worstGap.gap < -15) {
      return worstGap.name + "の改善";
    }
    
    return "現状維持・継続改善";
  } catch (error) {
    Logger.log("最優先項目決定エラー: " + error.toString());
    return "データ分析による最適化";
  }
}

/**
 * チャンネル動画リスト取得
 */
function getChannelVideos(channelId, apiKey, maxResults) {
  try {
    maxResults = maxResults || 50;
    
    // チャンネルのアップロードプレイリストIDを取得
    var channelUrl = "https://www.googleapis.com/youtube/v3/channels?part=contentDetails&id=" + channelId + "&key=" + apiKey;
    var channelResponse = UrlFetchApp.fetch(channelUrl);
    var channelData = JSON.parse(channelResponse.getContentText());
    
    if (!channelData.items || channelData.items.length === 0) {
      throw new Error("チャンネル情報が見つかりません");
    }
    
    var uploadsPlaylistId = channelData.items[0].contentDetails.relatedPlaylists.uploads;
    
    // プレイリストから動画リストを取得
    var playlistUrl = "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet&playlistId=" + 
                     uploadsPlaylistId + "&maxResults=" + maxResults + "&key=" + apiKey;
    var playlistResponse = UrlFetchApp.fetch(playlistUrl);
    var playlistData = JSON.parse(playlistResponse.getContentText());
    
    if (!playlistData.items) {
      return [];
    }
    
    // 各動画の詳細統計を取得
    var videoIds = playlistData.items.map(function(item) {
      return item.snippet.resourceId.videoId;
    }).join(',');
    
    var videosUrl = "https://www.googleapis.com/youtube/v3/videos?part=statistics,snippet&id=" + videoIds + "&key=" + apiKey;
    var videosResponse = UrlFetchApp.fetch(videosUrl);
    var videosData = JSON.parse(videosResponse.getContentText());
    
    return videosData.items || [];
  } catch (error) {
    Logger.log("動画リスト取得エラー: " + error.toString());
    return [];
  }
}

/**
 * 動画パフォーマンス分析
 */
function analyzeVideoPerformanceData(videoList) {
  try {
    if (!videoList || videoList.length === 0) {
      return { error: "分析する動画がありません" };
    }
    
    var analysis = {
      totalVideos: videoList.length,
      topPerformers: [],
      averageViews: 0,
      averageLikes: 0,
      averageComments: 0,
      engagementTrends: [],
      performanceDistribution: {}
    };
    
    var totalViews = 0;
    var totalLikes = 0;
    var totalComments = 0;
    
    // 各動画の分析
    var videoPerformance = [];
    for (var i = 0; i < videoList.length; i++) {
      var video = videoList[i];
      var stats = video.statistics;
      
      var views = parseInt(stats.viewCount || 0);
      var likes = parseInt(stats.likeCount || 0);
      var comments = parseInt(stats.commentCount || 0);
      
      totalViews += views;
      totalLikes += likes;
      totalComments += comments;
      
      var engagementRate = views > 0 ? ((likes + comments) / views * 100) : 0;
      
      videoPerformance.push({
        title: video.snippet.title,
        views: views,
        likes: likes,
        comments: comments,
        engagementRate: engagementRate,
        publishedAt: video.snippet.publishedAt
      });
    }
    
    // 平均値計算
    analysis.averageViews = Math.round(totalViews / videoList.length);
    analysis.averageLikes = Math.round(totalLikes / videoList.length);
    analysis.averageComments = Math.round(totalComments / videoList.length);
    
    // トップパフォーマー（視聴回数順）
    videoPerformance.sort(function(a, b) { return b.views - a.views; });
    analysis.topPerformers = videoPerformance.slice(0, 5);
    
    // パフォーマンス分布
    var highPerformers = videoPerformance.filter(function(v) { return v.views > analysis.averageViews * 1.5; }).length;
    var mediumPerformers = videoPerformance.filter(function(v) { return v.views >= analysis.averageViews * 0.5 && v.views <= analysis.averageViews * 1.5; }).length;
    var lowPerformers = videoPerformance.filter(function(v) { return v.views < analysis.averageViews * 0.5; }).length;
    
    analysis.performanceDistribution = {
      high: highPerformers,
      medium: mediumPerformers,
      low: lowPerformers
    };
    
    return analysis;
  } catch (error) {
    Logger.log("動画パフォーマンス分析エラー: " + error.toString());
    return { error: error.toString() };
  }
}

/**
 * 分析サマリーシートフォーマット
 */
function formatAnalysisSummarySheet(sheet) {
  // 列幅設定
  sheet.setColumnWidth(1, 120);  // ラベル
  sheet.setColumnWidth(2, 20);   // スペーサー
  sheet.setColumnWidth(3, 200);  // 値
  sheet.setColumnWidth(4, 20);   // スペーサー
  sheet.setColumnWidth(5, 120);  // ラベル
  sheet.setColumnWidth(6, 20);   // スペーサー
  sheet.setColumnWidth(7, 200);  // 値
  sheet.setColumnWidth(8, 50);   // 余白
  
  // 行高設定
  sheet.setRowHeight(1, 40);
  
  // 境界線設定
  sheet.getRange("A4:C11").setBorder(true, true, true, true, true, true);
  sheet.getRange("E4:G10").setBorder(true, true, true, true, true, true);
  sheet.getRange("A14:E18").setBorder(true, true, true, true, true, true);
}

/**
 * ギャップ分析シートフォーマット
 */
function formatGapAnalysisSheet(sheet) {
  // 列幅設定
  sheet.setColumnWidth(1, 100);  // 指標
  sheet.setColumnWidth(2, 20);   // スペーサー
  sheet.setColumnWidth(3, 150);  // 値
  sheet.setColumnWidth(4, 100);  // 差分
  sheet.setColumnWidth(5, 80);   // 達成率
  sheet.setColumnWidth(6, 60);   // 優先度
  sheet.setColumnWidth(7, 300);  // アクション
  
  // 行高設定
  sheet.setRowHeight(1, 40);
  
  // 境界線設定
  sheet.getRange("A6:C10").setBorder(true, true, true, true, true, true);
  sheet.getRange("A13:G17").setBorder(true, true, true, true, true, true);
}

/**
 * アクションプランシートフォーマット
 */
function formatActionPlanSheet(sheet) {
  // 列幅設定
  sheet.setColumnWidth(1, 150);  // 施策
  sheet.setColumnWidth(2, 200);  // 説明
  sheet.setColumnWidth(3, 80);   // 影響度
  sheet.setColumnWidth(4, 80);   // 工数
  sheet.setColumnWidth(5, 150);  // KPI
  sheet.setColumnWidth(6, 250);  // アクション
  sheet.setColumnWidth(7, 50);   // 余白
  sheet.setColumnWidth(8, 50);   // 余白
  
  // 行高設定
  sheet.setRowHeight(1, 40);
}

/**
 * メニューに包括的分析を追加
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // メインメニュー
  var menu = ui.createMenu("YouTube ツール");
  menu.addItem("🏠 統合ダッシュボード", "createOrShowMainDashboard");
  menu.addSeparator();
  menu.addItem("① API設定・テスト", "setApiKey");
  menu.addItem("② チャンネル情報取得", "processHandles");
  menu.addItem("③ ベンチマークレポート作成", "createBenchmarkReport");
  menu.addSeparator();
  menu.addItem("🚀 包括的チャンネル分析", "executeComprehensiveAnalysis");
  menu.addItem("📊 個別チャンネル分析", "executeChannelAnalysis");
  menu.addItem("🔍 ベンチマーク分析", "showBenchmarkDashboard");
  menu.addSeparator();
  
  // 既存の4_channelCheck.gs機能も統合
  if (typeof generateCompleteReport === 'function') {
    menu.addItem("🎯 ワンクリック完全分析", "generateCompleteReport");
    menu.addSubMenu(
      ui.createMenu("📈 個別分析モジュール")
        .addItem("📊 動画別パフォーマンス分析", "analyzeVideoPerformance")
        .addItem("👥 視聴者層分析", "analyzeAudience") 
        .addItem("❤️ エンゲージメント分析", "analyzeEngagement")
        .addItem("🔀 トラフィックソース分析", "analyzeTrafficSources")
        .addItem("🤖 AI改善提案", "generateAIRecommendations")
    );
    menu.addSeparator();
  }
  
  menu.addItem("シートテンプレート作成", "setupBasicSheet");
  menu.addItem("使い方ガイドを表示", "showHelpSheet");
  menu.addSeparator();
  menu.addItem("🔧 ダッシュボード再作成", "recreateDashboard");
  menu.addItem("🧪 入力検証テスト", "testInputValidation");
  menu.addToUi();
  
  // 統合ダッシュボードを作成または表示
  createOrShowMainDashboard();
}

/**
 * ダッシュボード再作成
 */
function recreateDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    
    if (dashboard) {
      ss.deleteSheet(dashboard);
    }
    
    createUnifiedDashboard();
    
    SpreadsheetApp.getUi().alert(
      "再作成完了",
      "統合ダッシュボードを再作成しました。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log("ダッシュボード再作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "エラー",
      "ダッシュボード再作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * 入力検証テスト
 */
function testInputValidation() {
  var testInputs = [
    "@YouTube",
    "https://www.youtube.com/@YouTube",
    "UC-9-kyTW8ZkZNDHQJ6FgpwQ",
    "https://www.youtube.com/channel/UC-9-kyTW8ZkZNDHQJ6FgpwQ",
    "invalid_input"
  ];
  
  var results = [];
  for (var i = 0; i < testInputs.length; i++) {
    var normalized = normalizeChannelInput(testInputs[i]);
    results.push(testInputs[i] + " → " + (normalized || "無効"));
  }
  
  SpreadsheetApp.getUi().alert(
    "入力検証テスト結果",
    "入力検証テスト結果:\n\n" + results.join("\n"),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * クイックアクションコマンドの更新（包括的分析対応）
 */
function handleQuickAction(command) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    
    if (!dashboard) return;
    
    command = command.toString().toLowerCase().trim();
    
    switch (command) {
      case "分析":
        executeChannelAnalysis();
        break;
      case "包括分析":
      case "包括的分析":
      case "完全分析":
        executeComprehensiveAnalysis();
        break;
      case "api":
        setApiKey();
        break;
      case "レポート":
        createBenchmarkReport();
        break;
      case "更新":
        updateDashboardDisplay();
        SpreadsheetApp.getUi().alert("ダッシュボードを更新しました");
        break;
      default:
        SpreadsheetApp.getUi().alert(
          "利用可能なコマンド",
          "利用可能なコマンド:\n\n• 分析: 基本チャンネル分析\n• 包括分析: 包括的チャンネル分析\n• API: API設定\n• レポート: ベンチマークレポート\n• 更新: ダッシュボード更新",
          SpreadsheetApp.getUi().ButtonSet.OK
        );
    }
    
    // コマンド実行後、セルをクリア
    dashboard.getRange("B9").setValue("ここに「分析」と入力してEnter");
    dashboard.getRange("B9").setBackground("#fff0f0");
    
  } catch (error) {
    Logger.log("クイックアクション実行エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "実行エラー",
      "コマンド実行中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}