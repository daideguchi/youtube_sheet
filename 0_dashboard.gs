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
          
          dashboard.getRange("B20").setValue(avgViews.toLocaleString());
          dashboard.getRange("B21").setValue(engagementRate.toFixed(2) + "%");
          
          // 総合評価
          var rating = calculateOverallRating(subscriberNum, engagementRate, videosNum);
          dashboard.getRange("B22").setValue(rating.score + "/100 (" + rating.grade + ")");
          
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
        
        dashboard.getRange("B20").setValue(avgViews.toLocaleString() + " 回/動画");
        dashboard.getRange("B21").setValue(engagementRate.toFixed(2) + "%");
        dashboard.getRange("B22").setValue(result.score + "/100 (" + result.grade + ")");
        
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
          
          dashboard.getRange("C15").setValue(avgViews.toLocaleString());
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
    suggestions.push("📈 エンゲージメント率向上のため、視聴者とのコミュニケーションを増やしましょう");
  }
  
  if (videoCount < 10) {
    suggestions.push("🎬 定期的な投稿でコンテンツ数を増やしましょう");
  } else if (videoCount > 100) {
    suggestions.push("✨ 豊富なコンテンツを活かし、プレイリスト整理で視聴しやすくしましょう");
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
    
    // 修正: 正しい関数呼び出し（引数なし）
    try {
      // 一時的にハンドル名をプロパティに保存
      PropertiesService.getDocumentProperties().setProperty("TEMP_HANDLE", handle);
      
      // 元の関数を呼び出し
      analyzeExistingChannel();
      
      // 分析完了後、ダッシュボードを更新
      Utilities.sleep(3000); // 3秒待機
      refreshDashboard();
      
      SpreadsheetApp.getUi().alert(
        "✅ 分析完了", 
        "チャンネルの基本分析が完了しました！\n\n結果がダッシュボードに表示されます。\n分析シートも自動作成されました。", 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } catch (analysisError) {
      dashboard.getRange("C10").setValue("分析失敗");
      dashboard.getRange("A20").setValue("分析中にエラーが発生しました: " + analysisError.toString());
      throw analysisError;
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
  menu.addItem("📊 個別チャンネル分析", "analyzeExistingChannel");
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