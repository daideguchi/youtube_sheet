/* eslint-disable */
/**
 * YouTubeチャンネル分析ダッシュボード
 * 一つのチャンネルに特化したシンプル分析ツール
 *
 * 作成者: Claude AI
 * バージョン: 3.0
 * 最終更新: 2025-01-22
 */
/* eslint-enable */

/**
 * チャンネル分析ダッシュボードのメイン起動
 */
function createOrShowMainDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardSheet = ss.getSheetByName("📊 チャンネル分析");
    
    if (!dashboardSheet) {
      createChannelAnalysisDashboard();
    } else {
      ss.setActiveSheet(dashboardSheet);
      refreshDashboard();
    }
  } catch (error) {
    Logger.log("ダッシュボード起動エラー: " + error.toString());
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
        var channelName = latestAnalysis.getRange("C4").getValue();
        var subscribers = latestAnalysis.getRange("C14").getValue();
        var totalViews = latestAnalysis.getRange("C15").getValue();
        var videoCount = latestAnalysis.getRange("C16").getValue();
        var createdDate = latestAnalysis.getRange("C13").getValue();
        
        dashboard.getRange("C10").setValue(channelName || "取得中...");
        dashboard.getRange("C11").setValue(subscribers || "取得中...");
        dashboard.getRange("C12").setValue(totalViews || "取得中...");
        dashboard.getRange("C13").setValue(videoCount || "取得中...");
        dashboard.getRange("C14").setValue(createdDate || "取得中...");
        
        // 簡単な評価を追加
        var subscriberNum = extractNumber(subscribers);
        var viewsNum = extractNumber(totalViews);
        var videosNum = extractNumber(videoCount);
        
        if (subscriberNum > 0 && viewsNum > 0) {
          var avgViews = Math.round(viewsNum / videosNum);
          var engagementRate = (avgViews / subscriberNum * 100);
          
          dashboard.getRange("C15").setValue(avgViews.toLocaleString());
          dashboard.getRange("C16").setValue(engagementRate.toFixed(2) + "%");
          
          // 総合スコア算出（簡易版）
          var totalScore = Math.min(100, Math.round(
            (subscriberNum / 1000) * 0.3 + 
            (avgViews / 1000) * 0.4 + 
            engagementRate * 10
          ));
          
          dashboard.getRange("I10").setValue(totalScore + " / 100");
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
    suggestions.push("🎯 収益化条件達成に向けて、登録者1000人を目指しましょう");
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
 * 基本分析を実行
 */
function runBasicAnalysis() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboard = ss.getSheetByName("📊 チャンネル分析");
    
    if (!dashboard) {
      SpreadsheetApp.getUi().alert("エラー", "ダッシュボードが見つかりません", SpreadsheetApp.getUi().ButtonSet.OK);
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
    
    // 既存の分析機能を呼び出し
    analyzeExistingChannel(handle);
    
    // 分析完了後、ダッシュボードを更新
    Utilities.sleep(2000); // 2秒待機
    refreshDashboard();
    
    SpreadsheetApp.getUi().alert(
      "分析完了", 
      "チャンネルの基本分析が完了しました！\n結果をダッシュボードでご確認ください。", 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("基本分析実行エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "分析エラー", 
      "分析中にエラーが発生しました: " + error.toString(), 
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
      return url.split("/@")[1].split("/")[0];
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
 * セルクリックイベントハンドラ
 */
function onEdit(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // チャンネル分析ダッシュボードでのクリックを処理
    if (sheet.getName() === "📊 チャンネル分析") {
      
      // 基本分析ボタン（I4）のクリック
      if (range.getRow() === 4 && range.getColumn() === 9) {
        runBasicAnalysis();
      }
      
      // API設定ボタン（B6）のクリック
      if (range.getRow() === 6 && range.getColumn() === 2) {
        setApiKey();
      }
      
      // クイックアクションボタン（25行目の偶数列）のクリック
      if (range.getRow() === 25) {
        var col = range.getColumn();
        if ([2, 4, 6, 8].indexOf(col) !== -1) {
          var functionName = sheet.getRange(26, col).getValue();
          
          if (functionName) {
            try {
              if (typeof eval(functionName) === 'function') {
                eval(functionName + '()');
              }
            } catch (error) {
              Logger.log("クイックアクション実行エラー: " + functionName + " - " + error.toString());
            }
          }
        }
      }
    }
    
  } catch (error) {
    Logger.log("onEditエラー: " + error.toString());
  }
} 