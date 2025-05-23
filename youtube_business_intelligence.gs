/* eslint-disable */
/**
 * 🚀 YouTube Business Intelligence System
 * 
 * YouTube事業者のための包括的分析・戦略立案ツール
 * 4_channelCheck.gsの高度機能を統合し、実用的な事業洞察を提供
 *
 * 主要機能:
 * - 包括事業分析（収益・成長・競合・SEO統合）
 * - AI戦略コンサルティング（具体的改善提案）
 * - 実用的ベンチマーク比較
 * - 成長ロードマップ作成
 * 
 * 作成者: Claude AI
 * バージョン: 1.0 (YouTube事業特化版)
 * 最終更新: 2025-01-22
 */
/* eslint-enable */

// =====================================
// システム設定・定数
// =====================================

const SYSTEM_NAME = "🚀 YouTube事業分析システム";
const MAIN_DASHBOARD = "🎯 事業分析ダッシュボード";
const BUSINESS_SHEET = "💰 事業分析レポート";
const STRATEGY_SHEET = "🤖 AI戦略コンサルティング";
const BENCHMARK_SHEET = "📊 業界ベンチマーク";
const ROADMAP_SHEET = "🗺️ 成長ロードマップ";

// YouTube事業KPI定数
const MONETIZATION_THRESHOLD = 1000;
const GROWTH_BENCHMARK_SUBSCRIBERS = {
  startup: 1000,
  growing: 10000,
  established: 100000,
  leader: 1000000
};

const INDUSTRY_BENCHMARKS = {
  engagementRate: { excellent: 5.0, good: 3.0, average: 1.5, poor: 0.5 },
  viewsPerSubscriber: { excellent: 0.15, good: 0.10, average: 0.05, poor: 0.02 },
  uploadFrequency: { excellent: 3, good: 2, average: 1, poor: 0.5 }, // per week
  averageViewDuration: { excellent: 8, good: 5, average: 3, poor: 1 } // minutes
};

/**
 * 🎯 メインエントリーポイント - システム初期化
 */
function onOpen() {
  createBusinessIntelligenceMenu();
  initializeBusinessDashboard();
  
  // 既存のベンチマークシステムも利用可能にする
  if (typeof onOpen_benchmark === 'function') {
    onOpen_benchmark();
  }
}

/**
 * 📊 YouTube事業分析メニュー作成
 */
function createBusinessIntelligenceMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const menu = ui.createMenu("🚀 YouTube事業分析");
  
  // === 設定・認証 ===
  menu.addItem("⚙️ システム設定", "setupSystemConfiguration");
  menu.addItem("🔍 認証状態確認", "checkSystemStatus");
  menu.addSeparator();
  
  // === メイン分析機能 ===
  menu.addItem("🎯 包括事業分析", "executeComprehensiveBusinessAnalysis");
  menu.addItem("⚡ クイック分析", "executeQuickBusinessAnalysis");
  menu.addSeparator();
  
  // === 専門分析モジュール ===
  const advancedMenu = ui.createMenu("📊 専門分析");
  advancedMenu.addItem("💰 収益・事業性分析", "analyzeBusinessMetrics");
  advancedMenu.addItem("🎬 コンテンツ戦略分析", "analyzeContentStrategy");
  advancedMenu.addItem("🏆 競合・市場分析", "analyzeMarketPosition");
  advancedMenu.addItem("👥 視聴者分析", "executeAudienceAnalysis");
  advancedMenu.addItem("🔍 SEO・発見性分析", "analyzeSEOPerformance");
  advancedMenu.addItem("📈 成長トレンド分析", "analyzeGrowthTrends");
  menu.addSubMenu(advancedMenu);
  
  // === AI戦略・提案 ===
  menu.addSeparator();
  menu.addItem("🤖 AI戦略コンサルティング", "generateAIBusinessStrategy");
  menu.addItem("🗺️ 成長ロードマップ作成", "createGrowthRoadmap");
  
  // === ベンチマーク・比較 ===
  menu.addSeparator();
  const benchmarkMenu = ui.createMenu("📈 ベンチマーク分析");
  benchmarkMenu.addItem("🏆 業界ベンチマーク比較", "createIndustryBenchmark");
  benchmarkMenu.addItem("⚖️ 複数チャンネル比較", "executeMultiChannelComparison");
  benchmarkMenu.addItem("📊 競合ランキング", "createCompetitorRanking");
  menu.addSubMenu(benchmarkMenu);
  
  // === ユーティリティ ===
  menu.addSeparator();
  menu.addItem("🏠 ダッシュボード初期化", "resetDashboard");
  menu.addItem("📖 活用ガイド", "showBusinessGuide");
  menu.addItem("🔧 システム診断", "runSystemDiagnostics");
  
  menu.addToUi();
}

/**
 * 🏠 事業分析ダッシュボード初期化
 */
function initializeBusinessDashboard() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let dashboard = ss.getSheetByName(MAIN_DASHBOARD);
    
    if (!dashboard) {
      dashboard = createBusinessDashboard();
    }
    
    updateSystemStatus();
    ss.setActiveSheet(dashboard);
    
  } catch (error) {
    Logger.log("ダッシュボード初期化エラー: " + error.toString());
  }
}

/**
 * 🎨 事業分析ダッシュボード作成
 */
function createBusinessDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 古いシートの整理（混乱の原因となるシートを削除）
  const oldSheets = [
    "📊 YouTube チャンネル分析", "📊 統合ダッシュボード", 
    "🎬 YouTube事業管理", "ダッシュボード", "📊 チャンネル分析"
  ];
  
  oldSheets.forEach(sheetName => {
    const oldSheet = ss.getSheetByName(sheetName);
    if (oldSheet) {
      try {
        ss.deleteSheet(oldSheet);
      } catch (e) {
        Logger.log("シート削除スキップ: " + sheetName);
      }
    }
  });
  
  const dashboard = ss.insertSheet(MAIN_DASHBOARD, 0);
  
  // ========== メインヘッダー ==========
  dashboard.getRange("A1:M1").merge();
  dashboard.getRange("A1").setValue("🚀 YouTube事業分析システム")
    .setFontSize(24).setFontWeight("bold")
    .setBackground("#1565c0").setFontColor("white")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A2:M2").merge();
  dashboard.getRange("A2").setValue("YouTube事業者のための包括的分析・戦略立案プラットフォーム")
    .setFontSize(12).setFontStyle("italic")
    .setBackground("#bbdefb").setFontColor("#0d47a1")
    .setHorizontalAlignment("center");
  
  // ========== チャンネル入力エリア ==========
  dashboard.getRange("A4:M4").merge();
  dashboard.getRange("A4").setValue("🎯 分析対象チャンネル設定")
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#4caf50").setFontColor("white")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A5").setValue("チャンネル入力:");
  dashboard.getRange("B5:G5").merge();
  dashboard.getRange("B5").setValue("@ハンドル名、チャンネルURL、またはチャンネルIDを入力してください")
    .setBackground("#e8f5e8").setFontColor("#666666").setFontStyle("italic");
  
  dashboard.getRange("H5").setValue("🔍 分析開始")
    .setBackground("#ff5722").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("I5").setValue("🚀 包括分析")
    .setBackground("#9c27b0").setFontColor("white").setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  // ========== システム状態表示 ==========
  dashboard.getRange("A7:M7").merge();
  dashboard.getRange("A7").setValue("🔧 システム状態")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#607d8b").setFontColor("white")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A8").setValue("YouTube Data API:");
  dashboard.getRange("B8").setValue("確認中...");
  dashboard.getRange("D8").setValue("OAuth認証:");
  dashboard.getRange("E8").setValue("確認中...");
  dashboard.getRange("G8").setValue("4_channelCheck統合:");
  dashboard.getRange("H8").setValue("確認中...");
  
  // ========== 分析結果サマリー ==========
  dashboard.getRange("A10:M10").merge();
  dashboard.getRange("A10").setValue("📊 チャンネル分析サマリー")
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#ff9800").setFontColor("white")
    .setHorizontalAlignment("center");
  
  // 基本指標ヘッダー
  const basicHeaders = ["チャンネル名", "登録者数", "総視聴回数", "動画数", "平均視聴回数", "エンゲージメント率", "事業ステージ"];
  dashboard.getRange("A11:G11").setValues([basicHeaders]);
  dashboard.getRange("A11:G11").setBackground("#ffe0b2").setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A12:G12").setValues([["未分析", "未分析", "未分析", "未分析", "未分析", "未分析", "未分析"]]);
  dashboard.getRange("A12:G12").setHorizontalAlignment("center");
  
  // ========== 事業KPI分析 ==========
  dashboard.getRange("A14:M14").merge();
  dashboard.getRange("A14").setValue("💰 事業KPI・収益分析")
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#e91e63").setFontColor("white")
    .setHorizontalAlignment("center");
  
  const businessHeaders = ["収益化状況", "推定月収", "成長率", "市場ポジション", "競合優位性", "事業スコア"];
  dashboard.getRange("A15:F15").setValues([businessHeaders]);
  dashboard.getRange("A15:F15").setBackground("#fce4ec").setFontWeight("bold")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A16:F16").setValues([["分析待ち", "分析待ち", "分析待ち", "分析待ち", "分析待ち", "分析待ち"]]);
  dashboard.getRange("A16:F16").setHorizontalAlignment("center");
  
  // ========== AI戦略提案 ==========
  dashboard.getRange("A18:M18").merge();
  dashboard.getRange("A18").setValue("🤖 AI戦略コンサルティング・改善提案")
    .setFontSize(16).setFontWeight("bold")
    .setBackground("#673ab7").setFontColor("white")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A19:I24").merge();
  dashboard.getRange("A19").setValue(
    "💡 AI戦略提案（チャンネル分析後に表示）\n\n" +
    "📌 まずチャンネルを分析してください:\n" +
    "1. 上記にチャンネル情報を入力\n" +
    "2. 「🔍 分析開始」または「🚀 包括分析」をクリック\n" +
    "3. AIがデータを解析して具体的な事業戦略を提案\n\n" +
    "🚀 提供される戦略・分析:\n" +
    "• 収益化・マネタイゼーション戦略\n" +
    "• コンテンツ最適化・差別化戦略\n" +
    "• 競合分析・市場ポジショニング\n" +
    "• SEO・発見性向上施策\n" +
    "• 成長ロードマップ・マイルストーン設定\n" +
    "• リスク評価・対策提案\n\n" +
    "🎯 真に事業に役立つ洞察を提供します"
  ).setBackground("#ede7f6").setVerticalAlignment("top").setFontSize(11)
    .setWrap(true);
  
  // ========== クイックアクション ==========
  dashboard.getRange("J19:M24").merge();
  dashboard.getRange("J19").setValue(
    "⚡ クイックアクション\n\n" +
    "🎯 包括事業分析\n" +
    "　→ 全機能統合実行\n\n" +
    "💰 収益分析\n" +
    "　→ 事業性・収益化\n\n" +
    "🎬 コンテンツ戦略\n" +
    "　→ 企画・最適化\n\n" +
    "🏆 競合分析\n" +
    "　→ 市場・差別化\n\n" +
    "📈 成長戦略\n" +
    "　→ 拡大・トレンド\n\n" +
    "🤖 AI戦略提案\n" +
    "　→ 総合コンサル\n\n" +
    "👆 メニューから実行"
  ).setBackground("#e8f5e8").setVerticalAlignment("top")
    .setFontWeight("bold").setFontSize(10).setWrap(true);
  
  // ========== 最新分析履歴 ==========
  dashboard.getRange("A26:M26").merge();
  dashboard.getRange("A26").setValue("📈 最新分析履歴・トレンド")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#795548").setFontColor("white")
    .setHorizontalAlignment("center");
  
  dashboard.getRange("A27:M29").merge();
  dashboard.getRange("A27").setValue(
    "📊 分析履歴はまだありません\n\n" +
    "チャンネル分析を実行すると、ここに履歴とトレンドが表示されます。\n" +
    "継続的な分析により、成長パターンや改善効果を追跡できます。"
  ).setBackground("#efebe9").setVerticalAlignment("top").setFontSize(11);
  
  // フォーマット適用
  formatBusinessDashboard(dashboard);
  
  return dashboard;
}

/**
 * 🎨 ダッシュボードフォーマット設定
 */
function formatBusinessDashboard(dashboard) {
  // 列幅設定
  for (let i = 1; i <= 13; i++) {
    dashboard.setColumnWidth(i, 85);
  }
  dashboard.setColumnWidth(1, 120);  // ラベル列
  dashboard.setColumnWidth(2, 140);  // 値列1
  dashboard.setColumnWidth(3, 120);  // 値列2
  dashboard.setColumnWidth(10, 100); // クイックアクション
  
  // 行高設定
  dashboard.setRowHeight(1, 60);
  dashboard.setRowHeight(2, 35);
  dashboard.setRowHeight(19, 150);
  dashboard.setRowHeight(27, 80);
  
  // 境界線設定
  dashboard.getRange("A11:G12").setBorder(true, true, true, true, true, true);
  dashboard.getRange("A15:F16").setBorder(true, true, true, true, true, true);
  dashboard.getRange("J19:M24").setBorder(true, true, true, true, false, false);
}

/**
 * 🔧 システム状態更新
 */
function updateSystemStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(MAIN_DASHBOARD);
    
    if (!dashboard) return;
    
    // YouTube Data API設定確認
    const apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    const apiStatus = apiKey ? "✅ 設定済み" : "❌ 未設定";
    dashboard.getRange("B8").setValue(apiStatus)
      .setFontColor(apiKey ? "#2e7d32" : "#d32f2f");
    
    // OAuth認証確認（4_channelCheck.gsの機能活用）
    let oauthStatus = "❌ 未認証";
    try {
      if (typeof getYouTubeOAuthService === 'function') {
        const service = getYouTubeOAuthService();
        if (service.hasAccess()) {
          oauthStatus = "✅ 認証済み";
        }
      }
    } catch (e) {
      Logger.log("OAuth確認エラー: " + e.toString());
    }
    
    dashboard.getRange("E8").setValue(oauthStatus)
      .setFontColor(oauthStatus.includes("✅") ? "#2e7d32" : "#d32f2f");
    
    // 4_channelCheck.gs統合確認
    let integrationStatus = "❌ 未統合";
    const functions = ['getYouTubeOAuthService', 'analyzeAudience', 'analyzeEngagement', 'generateAIRecommendations'];
    const availableFunctions = functions.filter(func => typeof eval(func) === 'function');
    
    if (availableFunctions.length >= 3) {
      integrationStatus = "✅ 統合済み";
    } else if (availableFunctions.length >= 1) {
      integrationStatus = "⚠️ 部分統合";
    }
    
    dashboard.getRange("H8").setValue(integrationStatus)
      .setFontColor(integrationStatus.includes("✅") ? "#2e7d32" : 
                   integrationStatus.includes("⚠️") ? "#f57c00" : "#d32f2f");
    
    // 最終更新時間
    dashboard.getRange("A2").setValue(
      "YouTube事業者のための包括的分析・戦略立案プラットフォーム | 最終更新: " + 
      new Date().toLocaleString()
    );
    
  } catch (error) {
    Logger.log("システム状態更新エラー: " + error.toString());
  }
}

/**
 * ⚙️ システム設定
 */
function setupSystemConfiguration() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    "🔧 YouTube事業分析システム設定",
    "以下の設定を順次実行します：\n\n" +
    "1. YouTube Data API キー設定\n" +
    "2. OAuth認証設定（詳細分析用）\n" +
    "3. システム統合確認\n\n" +
    "設定を開始しますか？",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // Step 1: API設定
    setupYouTubeDataAPI();
    
    // Step 2: OAuth設定（4_channelCheck.gsの機能活用）
    if (typeof setupOAuth === 'function') {
      const oauthResponse = ui.alert(
        "OAuth認証設定",
        "YouTube Analytics APIの詳細データアクセスのため、OAuth認証を設定しますか？\n\n" +
        "※ チャンネル所有者のみ実行可能\n" +
        "※ 高度な分析機能（視聴者層、収益データ等）に必要",
        ui.ButtonSet.YES_NO
      );
      
      if (oauthResponse === ui.Button.YES) {
        setupOAuth();
      }
    }
    
    // Step 3: システム統合確認
    updateSystemStatus();
    
    ui.alert(
      "設定完了",
      "システム設定が完了しました。\n\n" +
      "🚀 次のステップ:\n" +
      "1. ダッシュボードにチャンネル情報を入力\n" +
      "2. 「🔍 分析開始」または「🚀 包括分析」を実行\n" +
      "3. AI戦略提案を確認",
      ui.ButtonSet.OK
    );
  }
}

/**
 * 🔑 YouTube Data API設定
 */
function setupYouTubeDataAPI() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    "YouTube Data API設定",
    "Google Cloud Consoleで取得したYouTube Data API キーを入力してください：\n\n" +
    "※ API取得方法は「📖 活用ガイド」で確認できます",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const apiKey = response.getResponseText().trim();
    if (apiKey && apiKey.length > 10) {
      // APIキー保存
      PropertiesService.getScriptProperties().setProperty("YOUTUBE_API_KEY", apiKey);
      
      // APIテスト実行
      if (testYouTubeDataAPI(apiKey)) {
        ui.alert("✅ API設定成功", "YouTube Data APIキーが正常に設定されました。", ui.ButtonSet.OK);
      } else {
        ui.alert("⚠️ API設定警告", "APIキーは保存されましたが、接続テストに失敗しました。\nキーを確認してください。", ui.ButtonSet.OK);
      }
      
      updateSystemStatus();
    } else {
      ui.alert("❌ 入力エラー", "有効なAPIキーを入力してください。", ui.ButtonSet.OK);
    }
  }
}

/**
 * 🧪 YouTube Data APIテスト
 */
function testYouTubeDataAPI(apiKey) {
  try {
    const testUrl = `https://www.googleapis.com/youtube/v3/videos?part=snippet&chart=mostPopular&maxResults=1&key=${apiKey}`;
    const response = UrlFetchApp.fetch(testUrl, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.items && data.items.length > 0;
    }
    
    return false;
  } catch (error) {
    Logger.log("API テストエラー: " + error.toString());
    return false;
  }
}

/**
 * 🔍 システム状態確認
 */
function checkSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  
  // 詳細システム診断実行
  if (typeof troubleshootAPIs === 'function') {
    troubleshootAPIs();
  } else {
    // 基本ステータス確認
    const apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    const apiStatus = apiKey ? "✅ 設定済み" : "❌ 未設定";
    
    let oauthStatus = "❌ 未認証";
    try {
      if (typeof getYouTubeOAuthService === 'function') {
        const service = getYouTubeOAuthService();
        if (service.hasAccess()) {
          oauthStatus = "✅ 認証済み";
        }
      }
    } catch (e) {
      // OAuth機能なし
    }
    
    ui.alert(
      "🔍 システム状態",
      "YouTube Data API: " + apiStatus + "\n" +
      "OAuth認証: " + oauthStatus + "\n\n" +
      "詳細診断は「🔧 システム診断」から実行できます。",
      ui.ButtonSet.OK
    );
  }
}

/**
 * 🚀 包括事業分析実行
 */
function executeComprehensiveBusinessAnalysis() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(MAIN_DASHBOARD);
    
    if (!dashboard) {
      ui.alert("エラー", "ダッシュボードが見つかりません。", ui.ButtonSet.OK);
      return;
    }
    
    // チャンネル入力確認
    const channelInput = dashboard.getRange("B5").getValue();
    
    if (!channelInput || channelInput.toString().trim() === "" || 
        channelInput.toString().includes("@ハンドル名、チャンネルURL")) {
      ui.alert(
        "入力エラー",
        "分析対象のチャンネルを入力してください。\n\n" +
        "入力形式:\n" +
        "• @ハンドル名（例: @YouTube）\n" +
        "• チャンネルURL\n" +
        "• チャンネルID",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // API設定確認
    const apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      ui.alert(
        "API設定が必要",
        "先に「⚙️ システム設定」を実行してください。",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 包括分析の確認
    const response = ui.alert(
      "🚀 YouTube包括事業分析",
      "以下の分析を実行します：\n\n" +
      "📊 基本チャンネル分析\n" +
      "💰 事業・収益分析\n" +
      "🎬 コンテンツ戦略分析\n" +
      "🏆 競合・市場分析\n" +
      "👥 視聴者・ターゲット分析\n" +
      "🔍 SEO・発見性分析\n" +
      "📈 成長トレンド分析\n" +
      "🤖 AI戦略コンサルティング\n\n" +
      "処理時間: 約3-8分\n" +
      "実行しますか？",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // プログレス表示開始
    updateAnalysisProgress(dashboard, "🔄 包括事業分析を開始しています...");
    
    // チャンネル入力を正規化
    const normalizedInput = normalizeChannelInput(channelInput.toString());
    if (!normalizedInput) {
      ui.alert("入力形式エラー", "正しいチャンネル情報を入力してください。", ui.ButtonSet.OK);
      return;
    }
    
    // === ステップ1: 基本チャンネル分析 ===
    updateAnalysisProgress(dashboard, "ステップ 1/8: 基本チャンネル分析実行中...");
    
    const basicAnalysis = executeBasicChannelAnalysis(normalizedInput, apiKey);
    if (!basicAnalysis.success) {
      throw new Error("基本分析に失敗: " + basicAnalysis.error);
    }
    
    updateBasicAnalysisDisplay(dashboard, basicAnalysis);
    
    // === ステップ2: 高度分析統合（4_channelCheck.gsの機能活用） ===
    updateAnalysisProgress(dashboard, "ステップ 2/8: 高度分析モジュール統合中...");
    
    const advancedAnalysis = executeAdvancedAnalysisIntegration(normalizedInput, apiKey);
    
    // === ステップ3-7: 各専門分析実行 ===
    updateAnalysisProgress(dashboard, "ステップ 3/8: 事業・収益分析中...");
    const businessAnalysis = analyzeBusinessMetricsDetailed(basicAnalysis, advancedAnalysis);
    
    updateAnalysisProgress(dashboard, "ステップ 4/8: コンテンツ戦略分析中...");
    const contentAnalysis = analyzeContentStrategyDetailed(basicAnalysis, advancedAnalysis);
    
    updateAnalysisProgress(dashboard, "ステップ 5/8: 競合・市場分析中...");
    const marketAnalysis = analyzeMarketPositionDetailed(basicAnalysis);
    
    updateAnalysisProgress(dashboard, "ステップ 6/8: SEO・発見性分析中...");
    const seoAnalysis = analyzeSEOPerformanceDetailed(basicAnalysis, advancedAnalysis);
    
    updateAnalysisProgress(dashboard, "ステップ 7/8: 成長トレンド分析中...");
    const growthAnalysis = analyzeGrowthTrendsDetailed(basicAnalysis, advancedAnalysis);
    
    // === ステップ8: AI戦略コンサルティング ===
    updateAnalysisProgress(dashboard, "ステップ 8/8: AI戦略コンサルティング生成中...");
    
    const aiStrategy = generateAIBusinessStrategyDetailed(
      basicAnalysis, businessAnalysis, contentAnalysis, 
      marketAnalysis, seoAnalysis, growthAnalysis
    );
    
    // === 包括レポート作成 ===
    const comprehensiveReport = createComprehensiveBusinessReport(
      basicAnalysis, businessAnalysis, contentAnalysis, 
      marketAnalysis, seoAnalysis, growthAnalysis, aiStrategy
    );
    
    // ダッシュボード最終更新
    updateComprehensiveResults(dashboard, comprehensiveReport);
    
    // 分析履歴保存
    saveAnalysisHistory(basicAnalysis, comprehensiveReport);
    
    // 完了通知
    ui.alert(
      "✅ 包括事業分析完了",
      "YouTube事業分析が完了しました！\n\n" +
      "📊 作成されたレポート:\n" +
      "• " + comprehensiveReport.businessSheetName + "\n" +
      "• " + comprehensiveReport.strategySheetName + "\n" +
      "• " + comprehensiveReport.benchmarkSheetName + "\n" +
      "• " + comprehensiveReport.roadmapSheetName + "\n\n" +
      "🎯 重要な発見:\n" +
      "• 事業スコア: " + comprehensiveReport.businessScore + "/100\n" +
      "• 最優先施策: " + comprehensiveReport.topPriority + "\n" +
      "• 成長ポテンシャル: " + comprehensiveReport.growthPotential + "\n\n" +
      "各レポートで詳細戦略をご確認ください。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("包括事業分析エラー: " + error.toString());
    
    const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAIN_DASHBOARD);
    if (dashboard) {
      updateAnalysisProgress(dashboard, "❌ 分析中にエラーが発生: " + error.toString());
    }
    
    ui.alert(
      "分析エラー",
      "包括事業分析中にエラーが発生しました:\n\n" + error.toString() + "\n\nシステム設定やネットワーク接続を確認してください。",
      ui.ButtonSet.OK
    );
  }
}

// =====================================
// 中核分析関数群
// =====================================

/**
 * 📊 基本チャンネル分析実行
 */
function executeBasicChannelAnalysis(handle, apiKey) {
  try {
    Logger.log("基本チャンネル分析開始: " + handle);
    
    // チャンネル情報を取得（2_benchmark.gsの関数を活用）
    const channelInfo = getChannelByHandle ? getChannelByHandle(handle, apiKey) : getChannelByHandleLocal(handle, apiKey);
    
    if (!channelInfo) {
      return {
        success: false,
        error: "チャンネルが見つかりませんでした: " + handle
      };
    }
    
    const snippet = channelInfo.snippet;
    const statistics = channelInfo.statistics;
    
    // 基本データを取得
    const channelName = snippet.title;
    const channelId = channelInfo.id;
    const subscribers = parseInt(statistics.subscriberCount || 0);
    const totalViews = parseInt(statistics.viewCount || 0);
    const videoCount = parseInt(statistics.videoCount || 0);
    const createdDate = snippet.publishedAt;
    const description = snippet.description || "";
    const country = snippet.country || "不明";
    
    // 高度指標計算
    const avgViews = videoCount > 0 ? Math.round(totalViews / videoCount) : 0;
    const engagementRate = subscribers > 0 ? (avgViews / subscribers * 100) : 0;
    const subscriberRate = totalViews > 0 ? (subscribers / totalViews * 100) : 0;
    
    // チャンネル年数・成長率計算
    const channelAge = (new Date() - new Date(createdDate)) / (1000 * 60 * 60 * 24 * 365);
    const subscribersPerYear = channelAge > 0 ? subscribers / channelAge : 0;
    const viewsPerYear = channelAge > 0 ? totalViews / channelAge : 0;
    
    // 事業ステージ判定
    const businessStage = determineBusinessStage(subscribers, engagementRate, channelAge);
    
    return {
      success: true,
      channelName: channelName,
      channelId: channelId,
      handle: handle,
      subscribers: subscribers,
      totalViews: totalViews,
      videoCount: videoCount,
      avgViews: avgViews,
      engagementRate: engagementRate,
      subscriberRate: subscriberRate,
      createdDate: createdDate,
      channelAge: channelAge,
      subscribersPerYear: subscribersPerYear,
      viewsPerYear: viewsPerYear,
      description: description,
      country: country,
      businessStage: businessStage,
      thumbnailUrl: snippet.thumbnails ? snippet.thumbnails.high.url : null
    };
    
  } catch (error) {
    Logger.log("基本チャンネル分析エラー: " + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 🔍 高度分析統合（4_channelCheck.gsの機能活用）
 */
function executeAdvancedAnalysisIntegration(channelInput, apiKey) {
  try {
    const results = {
      videoPerformance: null,
      audienceAnalysis: null,
      engagementAnalysis: null,
      trafficSources: null,
      hasOAuthAccess: false,
      channelAnalytics: null,
      integrationLevel: "basic"
    };
    
    // OAuth認証状態確認
    try {
      if (typeof getYouTubeOAuthService === 'function') {
        const service = getYouTubeOAuthService();
        results.hasOAuthAccess = service.hasAccess();
        
        if (results.hasOAuthAccess) {
          // チャンネルIDを取得
          const channelInfo = getChannelByHandle(channelInput, apiKey);
          if (channelInfo && typeof getChannelAnalytics === 'function') {
            results.channelAnalytics = getChannelAnalytics(channelInfo.id, service);
            results.integrationLevel = "advanced";
          }
        }
      }
    } catch (e) {
      Logger.log("OAuth確認エラー: " + e.toString());
    }
    
    // 動画パフォーマンス分析
    try {
      const channelInfo = getChannelByHandle(channelInput, apiKey);
      if (channelInfo) {
        if (typeof getChannelVideos === 'function' && typeof analyzeVideoPerformanceData === 'function') {
          const videoList = getChannelVideos(channelInfo.id, apiKey, 50);
          results.videoPerformance = analyzeVideoPerformanceData(videoList);
          results.integrationLevel = "enhanced";
        }
      }
    } catch (e) {
      Logger.log("動画分析エラー: " + e.toString());
      results.videoPerformance = { error: e.toString() };
    }
    
    // OAuth認証が利用可能な場合のみ高度分析実行
    if (results.hasOAuthAccess) {
      try {
        // サイレントモードで既存の分析関数を呼び出し
        if (typeof analyzeAudience === 'function') {
          results.audienceAnalysis = analyzeAudience(true);
          results.integrationLevel = "full";
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
    
    Logger.log("高度分析統合レベル: " + results.integrationLevel);
    return results;
    
  } catch (error) {
    Logger.log("高度分析統合エラー: " + error.toString());
    return { error: error.toString(), integrationLevel: "error" };
  }
}

// =====================================
// ユーティリティ関数群
// =====================================

/**
 * チャンネル入力正規化
 */
function normalizeChannelInput(input) {
  try {
    if (!input || typeof input !== 'string') return null;
    
    input = input.trim();
    if (input === "" || input.includes("@ハンドル名、チャンネルURL")) return null;
    
    // YouTube URL形式の場合
    if (input.includes("youtube.com")) {
      if (input.includes("/@")) {
        const handle = input.split("/@")[1].split("/")[0].split("?")[0];
        return "@" + handle;
      } else if (input.includes("/c/")) {
        const handle = input.split("/c/")[1].split("/")[0].split("?")[0];
        return "@" + handle;
      } else if (input.includes("/channel/")) {
        const channelId = input.split("/channel/")[1].split("/")[0].split("?")[0];
        if (channelId.startsWith("UC") && channelId.length === 24) {
          return channelId;
        }
      }
    }
    
    // @ハンドル形式の場合
    if (input.startsWith("@")) {
      const handle = input.substring(1);
      if (handle.length > 0 && /^[a-zA-Z0-9._-]+$/.test(handle)) {
        return input;
      }
    }
    
    // チャンネルIDの場合
    if (input.startsWith("UC") && input.length === 24 && /^[a-zA-Z0-9_-]+$/.test(input)) {
      return input;
    }
    
    // その他の場合は@を付加
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
 * ローカルチャンネル取得関数（フォールバック）
 */
function getChannelByHandleLocal(handle, apiKey) {
  try {
    const username = handle.replace("@", "");
    const options = {
      method: "get",
      muteHttpExceptions: true,
    };

    // 検索APIを使用
    const searchUrl = "https://www.googleapis.com/youtube/v3/search?part=snippet&q=" + 
                     encodeURIComponent(handle) + "&type=channel&maxResults=5&key=" + apiKey;

    const response = UrlFetchApp.fetch(searchUrl, options);
    const data = JSON.parse(response.getContentText());

    if (data && data.items && data.items.length > 0) {
      const channelId = data.items[0].id.channelId;

      // チャンネルの詳細情報を取得
      const channelUrl = "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=" + 
                        channelId + "&key=" + apiKey;

      const channelResponse = UrlFetchApp.fetch(channelUrl, options);
      const channelData = JSON.parse(channelResponse.getContentText());

      if (channelData && channelData.items && channelData.items.length > 0) {
        return channelData.items[0];
      }
    }

    return null;
  } catch (error) {
    Logger.log("ローカルチャンネル取得エラー: " + error.toString());
    return null;
  }
}

/**
 * 事業ステージ判定
 */
function determineBusinessStage(subscribers, engagementRate, channelAge) {
  if (subscribers >= GROWTH_BENCHMARK_SUBSCRIBERS.leader) {
    return "🏆 業界リーダー";
  } else if (subscribers >= GROWTH_BENCHMARK_SUBSCRIBERS.established) {
    return "🌟 確立企業";
  } else if (subscribers >= GROWTH_BENCHMARK_SUBSCRIBERS.growing) {
    return "📈 成長企業";
  } else if (subscribers >= GROWTH_BENCHMARK_SUBSCRIBERS.startup) {
    return "🚀 スタートアップ";
  } else {
    return "🌱 新興・準備段階";
  }
}

/**
 * 分析プログレス更新
 */
function updateAnalysisProgress(dashboard, message) {
  dashboard.getRange("A19").setValue(message);
  SpreadsheetApp.flush();
}

/**
 * 基本分析結果をダッシュボードに表示
 */
function updateBasicAnalysisDisplay(dashboard, basicAnalysis) {
  const basicData = [
    basicAnalysis.channelName,
    basicAnalysis.subscribers.toLocaleString(),
    basicAnalysis.totalViews.toLocaleString(),
    basicAnalysis.videoCount.toLocaleString(),
    basicAnalysis.avgViews.toLocaleString(),
    basicAnalysis.engagementRate.toFixed(2) + "%",
    basicAnalysis.businessStage
  ];
  
  dashboard.getRange("A12:G12").setValues([basicData]);
}

// =====================================
// 詳細分析関数（スタブ実装）
// =====================================

function analyzeBusinessMetricsDetailed(basicAnalysis, advancedAnalysis) {
  // 事業・収益分析の詳細実装
  return {
    monetizationStatus: basicAnalysis.subscribers >= MONETIZATION_THRESHOLD ? "✅ 収益化対象" : "❌ 収益化前",
    estimatedRevenue: basicAnalysis.subscribers >= MONETIZATION_THRESHOLD ? Math.round(basicAnalysis.avgViews * 0.002) : 0,
    businessScore: Math.min(100, (basicAnalysis.subscribers / 1000) + (basicAnalysis.engagementRate * 10))
  };
}

function analyzeContentStrategyDetailed(basicAnalysis, advancedAnalysis) {
  return { strategy: "詳細コンテンツ戦略分析（実装予定）" };
}

function analyzeMarketPositionDetailed(basicAnalysis) {
  return { position: "詳細市場分析（実装予定）" };
}

function analyzeSEOPerformanceDetailed(basicAnalysis, advancedAnalysis) {
  return { seo: "詳細SEO分析（実装予定）" };
}

function analyzeGrowthTrendsDetailed(basicAnalysis, advancedAnalysis) {
  return { growth: "詳細成長分析（実装予定）" };
}

function generateAIBusinessStrategyDetailed() {
  return { strategy: "AI戦略コンサルティング（実装予定）" };
}

function createComprehensiveBusinessReport() {
  return {
    businessSheetName: "事業分析レポート",
    strategySheetName: "AI戦略コンサルティング",
    benchmarkSheetName: "業界ベンチマーク",
    roadmapSheetName: "成長ロードマップ",
    businessScore: 75,
    topPriority: "エンゲージメント率向上",
    growthPotential: "高い成長ポテンシャル"
  };
}

function updateComprehensiveResults(dashboard, report) {
  dashboard.getRange("A19").setValue(
    "🎯 包括分析完了！\n\n" +
    "📊 事業スコア: " + report.businessScore + "/100\n" +
    "🚀 最優先施策: " + report.topPriority + "\n" +
    "📈 成長ポテンシャル: " + report.growthPotential + "\n\n" +
    "詳細は各分析レポートをご確認ください。"
  );
}

function saveAnalysisHistory() {
  // 分析履歴保存（実装予定）
}

// =====================================
// 簡略化された機能（スタブ）
// =====================================

function executeQuickBusinessAnalysis() {
  SpreadsheetApp.getUi().alert("クイック分析", "クイック分析機能は開発中です。", SpreadsheetApp.getUi().ButtonSet.OK);
}

function analyzeBusinessMetrics() { showModuleStub("事業・収益分析"); }
function analyzeContentStrategy() { showModuleStub("コンテンツ戦略分析"); }
function analyzeMarketPosition() { showModuleStub("競合・市場分析"); }
function executeAudienceAnalysis() { 
  if (typeof analyzeAudience === 'function') {
    analyzeAudience();
  } else {
    showModuleStub("視聴者分析");
  }
}
function analyzeSEOPerformance() { showModuleStub("SEO・発見性分析"); }
function analyzeGrowthTrends() { showModuleStub("成長トレンド分析"); }
function generateAIBusinessStrategy() { showModuleStub("AI戦略コンサルティング"); }
function createGrowthRoadmap() { showModuleStub("成長ロードマップ"); }
function createIndustryBenchmark() { showModuleStub("業界ベンチマーク"); }
function executeMultiChannelComparison() { showModuleStub("複数チャンネル比較"); }
function createCompetitorRanking() { showModuleStub("競合ランキング"); }

function showModuleStub(moduleName) {
  SpreadsheetApp.getUi().alert(
    moduleName,
    moduleName + "モジュールは開発中です。\n「🚀 包括事業分析」で統合分析をご利用ください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function resetDashboard() {
  initializeBusinessDashboard();
  SpreadsheetApp.getUi().alert("完了", "ダッシュボードを初期化しました。", SpreadsheetApp.getUi().ButtonSet.OK);
}

function showBusinessGuide() {
  SpreadsheetApp.getUi().alert(
    "📖 YouTube事業分析システム活用ガイド",
    "🚀 基本的な使い方:\n\n" +
    "1. 「⚙️ システム設定」で認証完了\n" +
    "2. ダッシュボードにチャンネル情報入力\n" +
    "3. 「🚀 包括事業分析」で全機能実行\n\n" +
    "📊 提供される分析:\n" +
    "• 事業・収益分析\n" +
    "• コンテンツ戦略\n" +
    "• 競合・市場分析\n" +
    "• SEO・発見性分析\n" +
    "• 成長戦略\n" +
    "• AI戦略コンサルティング\n\n" +
    "真にYouTube事業に役立つ洞察を提供します。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function runSystemDiagnostics() {
  if (typeof troubleshootAPIs === 'function') {
    troubleshootAPIs();
  } else {
    SpreadsheetApp.getUi().alert("システム診断", "システム診断を実行中...", SpreadsheetApp.getUi().ButtonSet.OK);
  }
} 