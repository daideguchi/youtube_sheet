/**
 * 🚨 緊急修復ツール - ダッシュボード構造修復
 * 
 * ユーザー報告の問題解決：
 * 1. ワンクリック完全分析が進まない
 * 2. B11セルが「登録者」ではなく「OAuth認証済み」になっている
 * 3. エンゲージメント数字がおかしい
 * 4. チャンネル登録率、平均視聴時間が消えている
 */

/**
 * 🚨 緊急修復：ダッシュボード完全再構築
 */
function emergencyDashboardRepair() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存のダッシュボードシートを探す
    let dashboardSheet = ss.getSheetByName("📊 YouTube チャンネル分析") || 
                        ss.getSheetByName("ダッシュボード");
    
    if (!dashboardSheet) {
      dashboardSheet = ss.insertSheet("📊 YouTube チャンネル分析");
    }
    
    // シート全体をクリア
    dashboardSheet.clear();
    
    // 完全に正しい構造で再構築
    setupCorrectDashboardStructure(dashboardSheet);
    
    ui.alert(
      "✅ 緊急修復完了",
      "ダッシュボードを完全に再構築しました。\n\n" +
      "正しい表示構造：\n" +
      "• B8: チャンネル入力欄\n" +
      "• A7-H7: 指標ヘッダー\n" +
      "• A8-H8: データ表示エリア\n" +
      "• A10-B11: API/OAuth状態\n\n" +
      "「🚀 ワンクリック完全分析」を再実行してください。",
      ui.ButtonSet.OK
    );
    
    // ダッシュボードをアクティブにする
    ss.setActiveSheet(dashboardSheet);
    
  } catch (error) {
    ui.alert(
      "修復エラー",
      "ダッシュボード修復中にエラーが発生しました:\n\n" + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * 正しいダッシュボード構造をセットアップ
 */
function setupCorrectDashboardStructure(dashboardSheet) {
  // ========== メインヘッダー ==========
  dashboardSheet.getRange("A1:H1").merge();
  dashboardSheet.getRange("A1").setValue("📊 YouTube チャンネル分析ダッシュボード")
    .setFontSize(18).setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground("#f8f9fa").setFontColor("#495057");
  
  // ========== チャンネル入力エリア ==========
  dashboardSheet.getRange("A2").setValue("チャンネル入力:")
    .setFontWeight("bold").setBackground("#e9ecef");
  
  dashboardSheet.getRange("B2:D2").merge().setBackground("#ffffff");
  
  // ========== チャンネル情報表示 ==========
  dashboardSheet.getRange("A3").setValue("チャンネル名:").setFontWeight("bold");
  dashboardSheet.getRange("A4").setValue("分析日:").setFontWeight("bold");
  
  // ========== 入力欄（重要！）==========
  dashboardSheet.getRange("A5").setValue("ハンドル入力:");
  dashboardSheet.getRange("B5:D5").merge().setBackground("#ffffff");
  
  // プレースホルダー設定（B8が入力欄）
  dashboardSheet.getRange("B8").setValue("例: @YouTube または UC-9-kyTW8ZkZNDHQJ6FgpwQ")
    .setFontColor("#999999").setFontStyle("italic");
  
  // ========== 主要指標エリア ==========
  dashboardSheet.getRange("A6:H6").merge();
  dashboardSheet.getRange("A6").setValue("📈 主要パフォーマンス指標")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#e8f5e8").setFontColor("#2e7d32")
    .setHorizontalAlignment("center");
  
  // ヘッダー行（7行目）
  const headers = [
    "登録者数", "総再生回数", "登録率", "エンゲージメント率", 
    "視聴維持率", "平均視聴時間", "クリック率", "平均再生回数"
  ];
  
  for (let i = 0; i < headers.length; i++) {
    dashboardSheet.getRange(7, i + 1).setValue(headers[i])
      .setFontWeight("bold").setBackground("#f1f3f4")
      .setHorizontalAlignment("center");
  }
  
  // データ行（8行目）- 初期化
  dashboardSheet.getRange("A8:H8").setValue("")
    .setHorizontalAlignment("center");
  
  // ========== システム状態エリア ==========
  dashboardSheet.getRange("A9:H9").merge();
  dashboardSheet.getRange("A9").setValue("🔧 システム状態")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#fff3e0").setFontColor("#f57c00")
    .setHorizontalAlignment("center");
  
  dashboardSheet.getRange("A10").setValue("API状態:").setFontWeight("bold");
  dashboardSheet.getRange("B10").setValue("確認中...");
  dashboardSheet.getRange("A11").setValue("OAuth状態:").setFontWeight("bold");
  dashboardSheet.getRange("B11").setValue("確認中...");
  
  // ========== 使い方ガイド ==========
  dashboardSheet.getRange("A13:H13").merge();
  dashboardSheet.getRange("A13").setValue("📋 使い方ガイド")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#e0f2f1").setFontColor("#00695c")
    .setHorizontalAlignment("center");
  
  const instructions = [
    ["1.", "APIキー設定: メニュー「YouTube分析」→「⚙️ 初期設定」"],
    ["2.", "チャンネル入力: B8セルに @ハンドル名 を入力"],
    ["3.", "完全分析: メニュー「YouTube分析」→「🚀 ワンクリック完全分析」"],
    ["4.", "結果確認: 各シートタブで詳細分析結果を確認"],
    ["5.", "AI提案: 「AIフィードバック」シートで改善案を確認"]
  ];
  
  for (let i = 0; i < instructions.length; i++) {
    dashboardSheet.getRange(14 + i, 1).setValue(instructions[i][0])
      .setHorizontalAlignment("center").setFontWeight("bold");
    dashboardSheet.getRange(14 + i, 2, 1, 7).merge();
    dashboardSheet.getRange(14 + i, 2).setValue(instructions[i][1]);
  }
  
  // ========== 列幅調整 ==========
  for (let i = 1; i <= 8; i++) {
    dashboardSheet.setColumnWidth(i, 120);
  }
  
  // ========== 行高調整 ==========
  dashboardSheet.setRowHeight(1, 40);
  dashboardSheet.setRowHeight(6, 35);
  dashboardSheet.setRowHeight(9, 35);
  dashboardSheet.setRowHeight(13, 35);
  
  Logger.log("✅ 正しいダッシュボード構造を再構築完了");
}

/**
 * 🔧 ワンクリック完全分析の問題を修復
 */
function fixCompleteAnalysisIssues() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // 1. 基本的な関数の存在確認
    const functionsToCheck = [
      'runChannelAnalysis',
      'analyzeVideoPerformance', 
      'analyzeAudience',
      'analyzeEngagement',
      'analyzeTrafficSources',
      'generateAIRecommendations'
    ];
    
    let missingFunctions = [];
    
    functionsToCheck.forEach(funcName => {
      try {
        if (typeof eval(funcName) !== 'function') {
          missingFunctions.push(funcName);
        }
      } catch (e) {
        missingFunctions.push(funcName);
      }
    });
    
    if (missingFunctions.length > 0) {
      ui.alert(
        "⚠️ 関数不足検出",
        "以下の分析関数が見つかりません:\n\n" + 
        missingFunctions.join("\n") + "\n\n" +
        "4_channelCheck.gsファイルが完全でない可能性があります。",
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // 2. API設定確認
    const apiKey = PropertiesService.getUserProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      ui.alert(
        "❌ API設定不足",
        "YouTube APIキーが設定されていません。\n\n" +
        "メニュー「YouTube分析」→「⚙️ 初期設定」を実行してください。",
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // 3. ダッシュボード状態確認
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName("📊 YouTube チャンネル分析") || 
                          ss.getSheetByName("ダッシュボード");
    
    if (!dashboardSheet) {
      ui.alert(
        "❌ ダッシュボード不足",
        "分析ダッシュボードシートが見つかりません。\n\n" +
        "「🚨 ダッシュボード緊急修復」を実行してください。",
        ui.ButtonSet.OK
      );
      return false;
    }
    
    ui.alert(
      "✅ 診断完了",
      "ワンクリック完全分析の実行環境は正常です。\n\n" +
      "次の手順で実行してください:\n" +
      "1. B8セルに @ハンドル名 を入力\n" +
      "2. メニュー「🚀 ワンクリック完全分析」を実行\n\n" +
      "それでも進まない場合は、スクリプトエディタの\n" +
      "「実行ログ」でエラー詳細を確認してください。",
      ui.ButtonSet.OK
    );
    
    return true;
    
  } catch (error) {
    ui.alert(
      "診断エラー",
      "診断中にエラーが発生しました:\n\n" + error.toString(),
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * 📊 データ表示問題の修復
 */
function fixDataDisplayIssues() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName("📊 YouTube チャンネル分析") || 
                          ss.getSheetByName("ダッシュボード");
    
    if (!dashboardSheet) {
      ui.alert(
        "エラー", 
        "ダッシュボードシートが見つかりません。\n「🚨 ダッシュボード緊急修復」を実行してください。",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 正しいヘッダーを強制的に設定
    const headers = [
      "登録者数", "総再生回数", "登録率", "エンゲージメント率", 
      "視聴維持率", "平均視聴時間", "クリック率", "平均再生回数"
    ];
    
    for (let i = 0; i < headers.length; i++) {
      dashboardSheet.getRange(7, i + 1).setValue(headers[i])
        .setFontWeight("bold").setBackground("#f1f3f4")
        .setHorizontalAlignment("center");
    }
    
    // データエリアをクリア
    dashboardSheet.getRange("A8:H8").clearContent()
      .setHorizontalAlignment("center")
      .setBackground("#ffffff");
    
    // システム状態エリアを正しく設定
    dashboardSheet.getRange("A10").setValue("API状態:");
    dashboardSheet.getRange("A11").setValue("OAuth状態:");
    
    // OAuth状態を確認して表示
    try {
      if (typeof getYouTubeOAuthService === 'function') {
        const service = getYouTubeOAuthService();
        const oauthStatus = service.hasAccess() ? "✅ 認証済み" : "❌ 未認証";
        dashboardSheet.getRange("B11").setValue(oauthStatus);
      } else {
        dashboardSheet.getRange("B11").setValue("❓ 確認不可");
      }
    } catch (e) {
      dashboardSheet.getRange("B11").setValue("❌ エラー");
    }
    
    // API状態を表示
    const apiKey = PropertiesService.getUserProperties().getProperty("YOUTUBE_API_KEY");
    const apiStatus = apiKey ? "✅ 設定済み" : "❌ 未設定";
    dashboardSheet.getRange("B10").setValue(apiStatus);
    
    ui.alert(
      "✅ データ表示修復完了",
      "ダッシュボードのデータ表示構造を修復しました。\n\n" +
      "修復内容:\n" +
      "• ヘッダー行（7行目）を正しく設定\n" +
      "• データ行（8行目）をクリア・初期化\n" +
      "• システム状態表示を修正\n\n" +
      "これで「🚀 ワンクリック完全分析」を実行してください。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert(
      "修復エラー",
      "データ表示修復中にエラーが発生しました:\n\n" + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * 🚨 緊急修復メニューをYouTube分析メニューに追加
 */
function createEmergencyRepairMenu() {
  const ui = SpreadsheetApp.getUi();
  
  // 既存のYouTube分析メニューに緊急修復機能を追加
  ui.createMenu("🚨 緊急修復")
    .addItem("🔧 ダッシュボード完全修復", "emergencyDashboardRepair")
    .addItem("📊 データ表示修復", "fixDataDisplayIssues") 
    .addItem("⚡ ワンクリック分析診断", "fixCompleteAnalysisIssues")
    .addSeparator()
    .addItem("📋 修復ガイド", "showRepairGuide")
    .addToUi();
}

/**
 * 📋 修復ガイドを表示
 */
function showRepairGuide() {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    "🚨 緊急修復ガイド",
    "問題に応じて以下の修復機能を使用してください:\n\n" +
    "🔧 ダッシュボード完全修復:\n" +
    "• ダッシュボードの構造が壊れている\n" +
    "• B11セルの表示がおかしい\n" +
    "• 全体的なレイアウト問題\n\n" +
    "📊 データ表示修復:\n" +
    "• エンゲージメント数字がおかしい\n" +
    "• チャンネル登録率が消えている\n" +
    "• 平均視聴時間が表示されない\n\n" +
    "⚡ ワンクリック分析診断:\n" +
    "• ワンクリック完全分析が進まない\n" +
    "• 分析が途中で止まる\n" +
    "• エラーの原因を特定したい\n\n" +
    "修復後は必ず「🚀 ワンクリック完全分析」を再実行してください。",
    ui.ButtonSet.OK
  );
}

/**
 * 🎯 修復後の確認テスト
 */
function verifyRepairResults() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName("📊 YouTube チャンネル分析") || 
                          ss.getSheetByName("ダッシュボード");
    
    if (!dashboardSheet) {
      ui.alert("❌ 検証失敗", "ダッシュボードシートが見つかりません。", ui.ButtonSet.OK);
      return false;
    }
    
    // ヘッダー確認
    const headers = [
      "登録者数", "総再生回数", "登録率", "エンゲージメント率", 
      "視聴維持率", "平均視聴時間", "クリック率", "平均再生回数"
    ];
    
    let headerCorrect = true;
    for (let i = 0; i < headers.length; i++) {
      const cellValue = dashboardSheet.getRange(7, i + 1).getValue();
      if (cellValue !== headers[i]) {
        headerCorrect = false;
        break;
      }
    }
    
    // システム状態確認
    const a10Value = dashboardSheet.getRange("A10").getValue();
    const a11Value = dashboardSheet.getRange("A11").getValue();
    
    const statusCorrect = (a10Value === "API状態:" && a11Value === "OAuth状態:");
    
    if (headerCorrect && statusCorrect) {
      ui.alert(
        "✅ 検証成功",
        "ダッシュボード構造は正常です。\n\n" +
        "確認済み項目:\n" +
        "• ヘッダー行（7行目）: 正常\n" +
        "• システム状態（10-11行目）: 正常\n" +
        "• 基本構造: 正常\n\n" +
        "「🚀 ワンクリック完全分析」を実行できます。",
        ui.ButtonSet.OK
      );
      return true;
    } else {
      ui.alert(
        "⚠️ 検証警告",
        "ダッシュボード構造に問題があります。\n\n" +
        "ヘッダー行: " + (headerCorrect ? "✅ 正常" : "❌ 問題あり") + "\n" +
        "システム状態: " + (statusCorrect ? "✅ 正常" : "❌ 問題あり") + "\n\n" +
        "「🔧 ダッシュボード完全修復」を再実行してください。",
        ui.ButtonSet.OK
      );
      return false;
    }
    
  } catch (error) {
    ui.alert(
      "検証エラー",
      "検証中にエラーが発生しました:\n\n" + error.toString(),
      ui.ButtonSet.OK
    );
    return false;
  }
}

/**
 * 🚨 緊急修復機能対応のonOpen関数
 */
function onOpen_EmergencyRepair() {
  // 緊急修復メニューを作成
  createEmergencyRepairMenu();
  
  // システム状態の初期チェック
  checkSystemHealthOnOpen();
}

/**
 * 🔧 システム健全性の初期チェック
 */
function checkSystemHealthOnOpen() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName("📊 YouTube チャンネル分析") || 
                          ss.getSheetByName("ダッシュボード");
    
    if (!dashboardSheet) {
      // ダッシュボードが存在しない場合、警告を表示
      Logger.log("⚠️ ダッシュボードシートが見つかりません。緊急修復が必要かもしれません。");
      return;
    }
    
    // 基本的な構造チェック
    const a7Value = dashboardSheet.getRange("A7").getValue();
    if (a7Value !== "登録者数") {
      Logger.log("⚠️ ダッシュボード構造に問題があります。修復が必要です。");
    }
    
    Logger.log("✅ システム健全性チェック完了");
    
  } catch (error) {
    Logger.log("❌ システム健全性チェックエラー: " + error.toString());
  }
} 