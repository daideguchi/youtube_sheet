/* eslint-disable */
/**
 * YouTubeベンチマークチャンネルリスト
 * ハンドル名（@username形式）を入力することで、YouTubeチャンネルの情報を表形式で取得
 *
 * 作成者: Claude AI
 * バージョン: 5.3
 * 最終更新: 2025-05-22
 */
/* eslint-enable */

// スプレッドシートのUIにメニューを追加（シンプル/詳細切り替え対応）
function onOpen_benchmark() {
  var ui = SpreadsheetApp.getUi();
  var simpleMode = PropertiesService.getDocumentProperties().getProperty("BENCHMARK_SIMPLE_MODE") !== "false";

  var menu = ui.createMenu("📊 YouTube ベンチマーク");

  if (simpleMode) {
    // === シンプルモード ===
    menu.addItem("⚙️ API設定", "setApiKey");
    menu.addItem("🔍 チャンネル分析", "analyzeExistingChannel");
    menu.addItem("📊 複数チャンネル取得", "processHandles");
    menu.addItem("📈 レポート作成", "createBenchmarkReport");
    menu.addSeparator();
    menu.addItem("📖 使い方ガイド", "showHelpSheet");
    menu.addItem("⚙️ 詳細モードに切り替え", "enableBenchmarkSimpleMode");
  } else {
    // === 詳細モード（従来機能全て表示） ===
    menu.addItem("🏠 統合ダッシュボード", "createOrShowMainDashboard");
    menu.addSeparator();
    menu.addItem("① API設定・テスト", "setApiKey");
    menu.addItem("② チャンネル情報取得", "processHandles");
    menu.addItem("③ ベンチマークレポート作成", "createBenchmarkReport");
    menu.addSeparator();
    menu.addItem("📊 個別チャンネル分析", "analyzeExistingChannel");
    menu.addItem("🔍 ベンチマーク分析", "showBenchmarkDashboard");
    menu.addSeparator();
    menu.addItem("🎨 ダッシュボード色更新", "refreshDashboardColors");
    menu.addItem("シートテンプレート作成", "setupBasicSheet");
    menu.addItem("使い方ガイドを表示", "showHelpSheet");
    menu.addItem("⚙️ シンプルモードに切り替え", "enableBenchmarkSimpleMode");
  }
  
  menu.addToUi();

  // 初回実行時にダッシュボード作成
  createOrShowMainDashboard();
}

/**
 * 🎯 ベンチマークシステム シンプルモード切り替え
 */
function enableBenchmarkSimpleMode() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    "シンプルモード切り替え",
    "ベンチマークシステムのメニューをシンプルにしますか？\n\n" +
    "✅ シンプルモード:\n" +
    "・API設定\n" +
    "・チャンネル分析\n" +
    "・複数チャンネル取得\n" +
    "・レポート作成\n" +
    "・使い方ガイド\n\n" +
    "❌ 詳細モード:\n" +
    "・全ての機能とダッシュボード機能",
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    PropertiesService.getDocumentProperties().setProperty("BENCHMARK_SIMPLE_MODE", "true");
    ui.alert("設定完了", "シンプルモードが有効になりました。\nページを再読み込みしてください。", ui.ButtonSet.OK);
  } else {
    PropertiesService.getDocumentProperties().setProperty("BENCHMARK_SIMPLE_MODE", "false");
    ui.alert("設定完了", "詳細モードが有効になりました。\nページを再読み込みしてください。", ui.ButtonSet.OK);
  }
}

/**
 * 初回実行時のガイド表示
 */
function showInitialGuide() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getDocumentProperties();

    // 初回表示フラグをチェック
    if (!props.getProperty("initialGuideShown")) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        "YouTube ベンチマークツールへようこそ！",
        "基本的な使い方：\n\n" +
          "1. ハンドル名（@から始まる）を入力\n" +
          "2. 「YouTube ツール > ② チャンネル情報取得」を実行\n" +
          "3. 「YouTube ツール > ③ ベンチマークレポート作成」を実行\n\n" +
          "より詳しいガイドはメニューの「使い方ガイドを表示」から確認できます。\n\n" +
          "準備を始めますか？",
        ui.ButtonSet.YES_NO
      );

      // 初回表示フラグを設定
      props.setProperty("initialGuideShown", "true");

      // はいを選択した場合
      if (response == ui.Button.YES) {
        setupBasicSheet();
        setApiKey();
      }
    }
  } catch (error) {
    Logger.log("初期ガイド表示エラー: " + error.toString());
  }
}

/**
 * ダッシュボードを作成または表示する（統合ダッシュボードへのリダイレクト）
 */
function createOrShowDashboard() {
  // 新しい統合ダッシュボードにリダイレクト
  createOrShowMainDashboard();
}

/**
 * メインダッシュボードを表示（統合ダッシュボードへのリダイレクト）
 */
function showMainDashboard() {
  createOrShowMainDashboard();
}

/**
 * メインダッシュボードを作成
 */
function createMainDashboard() {
  // 統合ダッシュボードにリダイレクト
  createOrShowMainDashboard();
}

/**
 * 統合ダッシュボードを作成または表示
 */
function createOrShowMainDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardName = "📊 YouTube ベンチマーク管理";
    var dashboard = ss.getSheetByName(dashboardName);
    
    if (!dashboard) {
      dashboard = ss.insertSheet(dashboardName, 0);
      setupBenchmarkDashboard(dashboard);
    }
    
    ss.setActiveSheet(dashboard);
    updateBenchmarkDashboardStatus(dashboard);
    
  } catch (error) {
    Logger.log("統合ダッシュボード作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "ダッシュボード作成エラー",
      "統合ダッシュボードの作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ベンチマークダッシュボードをセットアップ
 */
function setupBenchmarkDashboard(dashboard) {
  // ヘッダー
  dashboard.getRange("A1:J1").merge();
  dashboard.getRange("A1").setValue("📊 YouTube ベンチマーク分析システム")
    .setFontSize(18).setFontWeight("bold")
    .setBackground("#e3f2fd").setFontColor("#1565c0")
    .setHorizontalAlignment("center");
    
  dashboard.getRange("A2:J2").merge();
  dashboard.getRange("A2").setValue("チャンネル情報の取得・分析・ベンチマークレポート作成")
    .setFontSize(12).setBackground("#f5f5f5")
    .setHorizontalAlignment("center");
    
  // システム状態
  dashboard.getRange("A4:J4").merge();
  dashboard.getRange("A4").setValue("🔧 システム状態")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#e8f5e8").setFontColor("#2e7d32")
    .setHorizontalAlignment("center");
    
  dashboard.getRange("A5").setValue("API状態:");
  dashboard.getRange("B5").setValue("確認中...");
  dashboard.getRange("D5").setValue("最終更新:");
  dashboard.getRange("E5").setValue(new Date().toLocaleString());
  
  // クイックアクション
  dashboard.getRange("A7:J7").merge();
  dashboard.getRange("A7").setValue("⚡ クイックアクション")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#fff3e0").setFontColor("#f57c00")
    .setHorizontalAlignment("center");
    
  var actions = [
    ["📝", "① API設定", "APIキーを設定してシステムを初期化", "setApiKey"],
    ["📊", "② チャンネル取得", "B列のハンドル名からチャンネル情報を取得", "processHandles"],
    ["📈", "③ レポート作成", "取得データからベンチマークレポートを作成", "createBenchmarkReport"],
    ["🔍", "④ 個別分析", "特定チャンネルの詳細分析を実行", "analyzeExistingChannel"],
    ["📖", "ヘルプ", "使い方ガイドを表示", "showHelpSheet"]
  ];
  
  for (var i = 0; i < actions.length; i++) {
    var row = 8 + i;
    dashboard.getRange(row, 1).setValue(actions[i][0]);
    dashboard.getRange(row, 2, 1, 2).merge();
    dashboard.getRange(row, 2).setValue(actions[i][1]);
    dashboard.getRange(row, 4, 1, 5).merge();
    dashboard.getRange(row, 4).setValue(actions[i][2]);
    dashboard.getRange(row, 9).setValue("▶ 実行")
      .setBackground("#4caf50").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center");
    dashboard.getRange(row, 10).setValue(actions[i][3]); // 関数名（非表示）
  }
  
  // データ入力エリア
  dashboard.getRange("A14:J14").merge();
  dashboard.getRange("A14").setValue("📊 チャンネル分析方法")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#f3e5f5").setFontColor("#7b1fa2")
    .setHorizontalAlignment("center");
    
  dashboard.getRange("A15").setValue("🎯 個別分析:");
  dashboard.getRange("B15:D15").merge();
  dashboard.getRange("B15").setValue("上記の「④ 個別分析」ボタンから1つのチャンネルを詳細分析")
    .setBackground("#e8f5e8").setFontStyle("italic");
    
  dashboard.getRange("A16").setValue("📊 複数比較:");
  dashboard.getRange("B16:D16").merge();
  dashboard.getRange("B16").setValue("別シートのB列に@ハンドル名リストを入力→「② チャンネル取得」")
    .setBackground("#fff3e0").setFontStyle("italic");
  
  dashboard.getRange("A17").setValue("処理状況:");
  dashboard.getRange("B17").setValue("待機中");
  
  // 統計情報エリア
  dashboard.getRange("A19:J19").merge();
  dashboard.getRange("A19").setValue("📈 統計情報")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#e0f2f1").setFontColor("#00695c")
    .setHorizontalAlignment("center");
    
  dashboard.getRange("A20").setValue("取得済みチャンネル数:");
  dashboard.getRange("B20").setValue("0件");
  dashboard.getRange("D20").setValue("平均登録者数:");
  dashboard.getRange("E20").setValue("未計算");
  
  dashboard.getRange("A21").setValue("最高登録者数:");
  dashboard.getRange("B21").setValue("未計算");
  dashboard.getRange("D21").setValue("最新レポート:");
  dashboard.getRange("E21").setValue("未作成");
  
  dashboard.getRange("A22").setValue("総視聴回数:");
  dashboard.getRange("B22").setValue("計算中...");
  dashboard.getRange("D22").setValue("処理状況:");
  dashboard.getRange("E22").setValue("待機中");
  
  dashboard.getRange("A23").setValue("動画本数:");
  dashboard.getRange("B23").setValue("計算中...");
  dashboard.getRange("D23").setValue("処理状況:");
  dashboard.getRange("E23").setValue("待機中");
  
  // フォーマット
  formatBenchmarkDashboard(dashboard);
}

/**
 * ベンチマークダッシュボードのフォーマット設定
 */
function formatBenchmarkDashboard(dashboard) {
  // 列幅設定
  for (var i = 1; i <= 10; i++) {
    dashboard.setColumnWidth(i, 80);
  }
  dashboard.setColumnWidth(4, 200); // 説明列
  dashboard.setColumnWidth(9, 80);  // ボタン列
  dashboard.setColumnWidth(10, 10); // 関数名列（非表示）
  
  // 行高設定
  dashboard.setRowHeight(1, 40);
  dashboard.setRowHeight(2, 25);
  
  // ボーダー設定
  dashboard.getRange("A8:I12").setBorder(true, true, true, true, true, true);
  dashboard.getRange("A19:D22").setBorder(true, true, true, true, true, true);
  
  // 列を非表示（関数名列）
  dashboard.hideColumns(10);
}

/**
 * ベンチマークダッシュボードの状態更新
 */
function updateBenchmarkDashboardStatus(dashboard) {
  try {
    // API状態確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    var apiStatus = apiKey ? "✅ 設定済み" : "❌ 未設定";
    dashboard.getRange("B5").setValue(apiStatus);
    
    // データ件数確認
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var dataCount = 0;
    
    // 他のシートでデータをチェック（ベンチマーク用のシート）
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i].getName();
      if (sheetName !== dashboard.getName() && !sheetName.includes("レポート") && !sheetName.includes("ガイド")) {
        var data = sheets[i].getRange("B:B").getValues();
        for (var j = 1; j < data.length; j++) {
          if (data[j][0] && data[j][0].toString().startsWith("@")) {
            dataCount++;
          }
        }
        break; // 最初のデータシートのみチェック
      }
    }
    
    dashboard.getRange("B20").setValue(dataCount + "件");
    dashboard.getRange("E5").setValue(new Date().toLocaleString());
    
  } catch (error) {
    Logger.log("ダッシュボード状態更新エラー: " + error.toString());
  }
}

/**
 * メインダッシュボードのフォーマット
 */
function formatMainDashboard(sheet) {
  // この関数は統合ダッシュボードで置き換えられました
  return;
}

/**
 * 使い方ガイドを専用シートに表示
 */
function showHelpSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var existingSheet = ss.getSheetByName("使い方ガイド");

    // 既存のガイドシートがあれば削除
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }

    // 新しいガイドシートを作成
    var helpSheet = ss.insertSheet("使い方ガイド");

    // ガイドコンテンツを設定
    helpSheet.appendRow(["YouTube チャンネル分析ツール - 使い方ガイド"]);
    helpSheet.appendRow([""]);
    helpSheet.appendRow(["基本的な使い方"]);
    helpSheet.appendRow([
      "1.",
      "API設定",
      "「YouTube ツール > ① API設定・テスト」からYouTube Data APIキーを設定します",
    ]);
    helpSheet.appendRow([
      "2.",
      "ハンドル入力",
      "B列に@から始まるYouTubeハンドル名を入力します（例: @YouTube）",
    ]);
    helpSheet.appendRow([
      "3.",
      "情報取得",
      "「YouTube ツール > ② チャンネル情報取得」を選択",
    ]);
    helpSheet.appendRow([
      "4.",
      "レポート作成",
      "「YouTube ツール > ③ ベンチマークレポート作成」を選択",
    ]);
    helpSheet.appendRow([""]);
    helpSheet.appendRow(["ヒント"]);
    helpSheet.appendRow([
      "・",
      "ジャンル分類",
      "A列にはジャンルなど任意の分類を入力できます（レポートに反映されます）",
    ]);
    helpSheet.appendRow([
      "・",
      "API取得上限",
      "1日あたりのAPI呼び出し上限があります。多数のチャンネルを分析する場合は注意してください",
    ]);
    helpSheet.appendRow([
      "・",
      "データ更新",
      "情報を再取得する場合は、再度「② チャンネル情報取得」を実行してください",
    ]);

    // 書式設定
    helpSheet.getRange("A1").setFontSize(16).setFontWeight("bold");
    helpSheet.getRange("A3").setFontSize(14).setFontWeight("bold");
    helpSheet.getRange("A4:C13").setBorder(true, true, true, true, true, true);
    helpSheet.getRange("A9").setFontSize(14).setFontWeight("bold");
    helpSheet.getRange("A4:A7").setHorizontalAlignment("center");
    helpSheet.getRange("A10:A13").setHorizontalAlignment("center");
    helpSheet.getRange("B4:B13").setFontWeight("bold");

    // 列幅の調整
    helpSheet.setColumnWidth(1, 40);
    helpSheet.setColumnWidth(2, 120);
    helpSheet.setColumnWidth(3, 500);

    // ヘッダー行の色設定
    helpSheet.getRange("A1:C1").setBackground("#e3f2fd").setFontColor("#1565c0");
    helpSheet.getRange("A3:C3").setBackground("#e8f5e8");
    helpSheet.getRange("A9:C9").setBackground("#e8f5e8");

    // シートをアクティブに
    ss.setActiveSheet(helpSheet);
  } catch (error) {
    Logger.log("ヘルプシート作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "ヘルプシートの作成中にエラーが発生しました: " + error.toString()
    );
  }
}

/**
 * APIキーを設定する関数
 */
function setApiKey() {
  var ui = SpreadsheetApp.getUi();

  try {
    // APIキーの確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      ui.alert(
        "APIキーが設定されていません",
        "「YouTube ツール > ① API設定・テスト」を実行してAPIキーを設定してください。",
        ui.ButtonSet.OK
      );
      return;
    }

    // プロンプト表示
    var response = ui.prompt(
      "YouTube API設定",
      "提供されたAPIキーを入力してください:" +
        (apiKey
          ? "\n\n現在設定されているAPIキー: " + apiKey
          : ""),
      ui.ButtonSet.OK_CANCEL
    );

    // OKボタンがクリックされた場合
    if (response.getSelectedButton() == ui.Button.OK) {
      var apiKey = response.getResponseText().trim();
      if (apiKey) {
        PropertiesService.getScriptProperties().setProperty(
          "YOUTUBE_API_KEY",
          apiKey
        );

        // APIテストを実行
        var testResponse = testYouTubeAPI(true);

        if (testResponse.success) {
          ui.alert(
            "API設定完了",
            "YouTube Data API キーが保存され、接続テストに成功しました。\n\nB列にハンドル名（@から始まる）を入力し、「YouTube ツール > ② チャンネル情報取得」を実行してください。",
            ui.ButtonSet.OK
          );
        } else {
          ui.alert(
            "APIテスト失敗",
            "APIキーは保存されましたが、接続テストに失敗しました。\n\nエラー: " +
              testResponse.error,
            ui.ButtonSet.OK
          );
        }

        // 基本的なシートの初期化（ヘッダー設定など）
        setupBasicSheet();
      } else {
        ui.alert("エラー", "APIキーが入力されていません。", ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    ui.alert(
      "エラー",
      "API設定中にエラーが発生しました: " + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * 基本的なシート初期化（ヘッダー設定など）
 */
function setupBasicSheet() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // ヘッダー行を設定（既存の値を上書きしない）
    var headers = [
      ["A1", "ジャンル"],
      ["B1", "ハンドル名"],
      ["C1", "チャンネル名"],
      ["D1", "チャンネルID"],
      ["E1", "登録者数"],
      ["F1", "総視聴回数"],
      ["G1", "動画本数"],
      ["H1", "作成日"],
      ["I1", "チャンネルURL"],
      ["J1", "サムネイル"],
    ];

    // ヘッダーを設定
    for (var i = 0; i < headers.length; i++) {
      if (!sheet.getRange(headers[i][0]).getValue()) {
        sheet.getRange(headers[i][0]).setValue(headers[i][1]);
      }
    }

    // ヘッダー行の書式設定
    sheet.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#f3f3f3");

    // 列幅の調整
    sheet.setColumnWidth(1, 120); // ジャンル
    sheet.setColumnWidth(2, 150); // ハンドル名
    sheet.setColumnWidth(3, 200); // チャンネル名
    sheet.setColumnWidth(4, 150); // チャンネルID
    sheet.setColumnWidth(5, 100); // 登録者数
    sheet.setColumnWidth(6, 120); // 総視聴回数
    sheet.setColumnWidth(7, 100); // 動画本数
    sheet.setColumnWidth(8, 120); // 作成日
    sheet.setColumnWidth(9, 200); // チャンネルURL
    sheet.setColumnWidth(10, 120); // サムネイル画像

    // ステータス表示用セルを作成
    if (!sheet.getRange("K1").getValue()) {
      sheet.getRange("K1").setValue("ステータス:");
      sheet.getRange("L1").setValue("準備完了");
    }

    // プレースホルダーを設定（説明用の注釈も追加）
    sheet.getRange("A2").setValue("例）シニア・自己啓発");
    sheet.getRange("A2").setFontColor("#999999").setFontStyle("italic");
    sheet
      .getRange("A2")
      .setNote("ジャンルの例です。A列にはジャンル名を入力してください");

    sheet.getRange("B2").setValue("@と入力してハンドル名を入力");
    sheet.getRange("B2").setFontColor("#999999").setFontStyle("italic");
    sheet
      .getRange("B2")
      .setNote("B列にはYouTubeのハンドル名（@から始まる）を入力してください");

    // 編集時に書式をリセットするためのトリガーを設定
    setEditTrigger();
  } catch (error) {
    Logger.log("シート初期化エラー: " + error.toString());
  }
}

/**
 * 編集トリガーを設定
 */
function setEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();
  var hasEditTrigger = false;

  // 既存のトリガーをチェック
  for (var i = 0; i < triggers.length; i++) {
    if (
      triggers[i].getHandlerFunction() === "onEdit" &&
      triggers[i].getEventType() === ScriptApp.EventType.ON_EDIT
    ) {
      hasEditTrigger = true;
      break;
    }
  }

  // トリガーがなければ作成
  if (!hasEditTrigger) {
    ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();
  }
}

/**
 * 編集時のイベントハンドラ - 統合ダッシュボード対応版
 */
/* function onEdit_disabled(e) {
  try {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var sheetName = sheet.getName();
    var value = range.getValue();
    
    // ========== 統合ダッシュボードでのコマンド処理 ==========
    if (sheetName === "📊 YouTube チャンネル分析") {
      
      // B9セル（操作セル）でのコマンド入力
      if (range.getRow() === 9 && range.getColumn() === 2) {
        if (value && value.toString().trim() !== "" && value.toString().trim() !== "ここに「分析」と入力してEnter") {
          // コマンドを実行
          handleQuickAction(value);
        }
        return;
      }
    }
    
    // ========== 既存のベンチマークダッシュボードでのクリック処理 ==========
    if (sheetName === "📊 統合ダッシュボード" || sheetName === "🔍 ベンチマーク分析") {
      
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
              if (typeof eval(functionName) === 'function') {
                eval(functionName + '()');
              }
              
              // 実行後に統計を更新
              Utilities.sleep(1000); // 1秒待機
              if (typeof updateDashboardStatistics === 'function') {
                updateDashboardStatistics();
              }
              
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
    
    // ========== 従来のプレースホルダー処理 ==========
    // B列でのプレースホルダー処理（従来の機能）
    if (range.getColumn() === 2 && range.getRow() >= 2) {
      var value = range.getValue();
      
      if (value && value.toString().trim() !== "") {
        // プレースホルダーの場合はクリア
        if (value.toString().includes("@と入力して") || value.toString().includes("例）")) {
          range.setValue("");
          return;
        }
        
        // 通常のテキストの場合はフォーマットをリセット
        range.setFontColor("#000000").setFontStyle("normal");
      }
    }
    
  } catch (error) {
    Logger.log("onEditエラー: " + error.toString());
  }
}
*/

/**
 * テスト関数 - YouTube APIが正しく動作しているか確認
 * @param {boolean} silent - trueの場合、ダイアログを表示せずに結果を返す
 * @return {Object} - 結果オブジェクト（silentがtrueの場合）
 */
function testYouTubeAPI(silent) {
  var ui = SpreadsheetApp.getUi();
  var result = { success: false, error: "" };

  try {
    // APIキーを取得
    var apiKey =
      PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");

    // APIキーがない場合は設定を促す
    if (!apiKey) {
      var message =
        "「YouTube ツール > ① API設定・テスト」を実行してAPIキーを設定してください。";
      if (silent) {
        result.error = "APIキーが設定されていません";
        return result;
      } else {
        ui.alert("APIキーが設定されていません", message, ui.ButtonSet.OK);
        return;
      }
    }

    // ステータス表示
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var statusCell = sheet.getRange("L1");
    statusCell.setValue("テスト中...");

    // 単純な検索テスト
    var options = {
      method: "get",
      muteHttpExceptions: true,
    };

    // チャンネル情報のテスト (YouTube公式チャンネル)
    var url =
      "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&forUsername=YouTube&key=" +
      apiKey;
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());

    if (data && data.items && data.items.length > 0) {
      statusCell.setValue("テスト成功");
      result.success = true;

      if (!silent) {
        ui.alert(
          "API接続成功！",
          "YouTube APIに正常に接続できました。\n\n" +
            "チャンネル名: " +
            data.items[0].snippet.title +
            "\n" +
            "登録者数: " +
            parseInt(
              data.items[0].statistics.subscriberCount
            ).toLocaleString() +
            "\n" +
            "動画数: " +
            data.items[0].statistics.videoCount,
          ui.ButtonSet.OK
        );
      }
    } else {
      // 検索APIをテスト
      url =
        "https://www.googleapis.com/youtube/v3/search?part=snippet&q=YouTube&type=channel&maxResults=1&key=" +
        apiKey;
      response = UrlFetchApp.fetch(url, options);
      data = JSON.parse(response.getContentText());

      if (data && data.items && data.items.length > 0) {
        statusCell.setValue("テスト成功");
        result.success = true;

        if (!silent) {
          ui.alert(
            "API検索接続成功！",
            "検索APIは正常に動作しています。\n\n" +
              "チャンネル名: " +
              data.items[0].snippet.title,
            ui.ButtonSet.OK
          );
        }
      } else {
        statusCell.setValue("テスト失敗");
        result.error =
          "APIは接続できましたが、期待した応答が得られませんでした。";

        if (!silent) {
          ui.alert(
            "API応答エラー",
            result.error + "\n\nレスポンス: " + response.getContentText(),
            ui.ButtonSet.OK
          );
        }
      }
    }

    return result;
  } catch (error) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var statusCell = sheet.getRange("L1");
    statusCell.setValue("エラー");

    result.error = error.toString();

    if (!silent) {
      ui.alert(
        "APIエラー",
        "YouTube APIに接続できませんでした: " + error.toString(),
        ui.ButtonSet.OK
      );
    }

    return result;
  }
}

/**
 * 入力されたハンドル名を処理する関数
 */
function processHandles() {
  try {
    // アクティブなスプレッドシートを取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();
    var ui = SpreadsheetApp.getUi();
    var sheetName = sheet.getName();

    // APIキーの確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      ui.alert(
        "APIキーが設定されていません",
        "「YouTube ツール > ① API設定・テスト」を実行してAPIキーを設定してください。",
        ui.ButtonSet.OK
      );
      return;
    }

    // 統合ダッシュボードの場合は専用処理
    if (sheetName === "📊 YouTube チャンネル分析" || sheetName === "🎯 事業分析ダッシュボード") {
      var channelInput = sheet.getRange("B5").getValue(); // B5セルに修正
      
      if (!channelInput || channelInput.toString().trim() === "" || 
          channelInput.toString().includes("@ハンドル名、チャンネルURL")) {
        ui.alert(
          "入力エラー",
          "🎯 事業分析ダッシュボードのB5セルにチャンネル情報を入力してください。\n\n" +
          "📌 入力形式:\n" +
          "• @YouTube（ハンドル名）\n" +
          "• https://www.youtube.com/@YouTube（URL）\n" +
          "• UCチャンネルID\n\n" +
          "💡 ヒント: B5セルのプレースホルダーテキストをクリアして入力してください",
          ui.ButtonSet.OK
        );
        return;
      }
      
      // 入力を正規化
      var handle = normalizeChannelInputForProcess(channelInput.toString());
      if (!handle) {
        ui.alert(
          "入力形式エラー",
          "正しいYouTubeチャンネルURL または @ハンドル名を入力してください。",
          ui.ButtonSet.OK
        );
        return;
      }
      
      // プレースホルダーをクリア（次回のため）
      if (channelInput.toString().includes("@ハンドル名、チャンネルURL")) {
        sheet.getRange("B5").setValue("");
      }
      
      // 事業分析ダッシュボード専用の分析実行
      executeBusinessChannelAnalysis(handle, apiKey, sheet);
      return;
    }

    // 従来のベンチマーク方式（B列リスト処理）
    // プログレスバーセルを用意
    var statusCell = sheet.getRange("L1");
    statusCell.setValue("準備中...");

    // B列のデータを取得（ハンドル名の列）
    var data = sheet.getRange("B:B").getValues();
    var colors = sheet.getRange("B:B").getFontColors(); // フォント色を取得
    var handles = [];

    // ハンドル名をリストに追加（空でないセルかつ@で始まるもの）
    for (var i = 1; i < data.length; i++) {
      // i=1 からスタート（ヘッダー行をスキップ）
      var handle = data[i][0].toString().trim();
      var color = colors[i][0];

      // 実際のハンドル名のみを処理（プレースホルダーでないもの＝黒いテキスト）
      if (
        handle &&
        handle.startsWith("@") &&
        color !== "#999999" && // グレーのプレースホルダーは除外
        handle !== "@と入力してハンドル名を入力"
      ) {
        handles.push({
          handle: handle,
          row: i + 1, // スプレッドシートの行番号（1-indexed）
        });
      }
    }

    if (handles.length === 0) {
      statusCell.setValue("エラー");
      ui.alert(
        "エラー",
        "現在のシート「" + sheetName + "」のB列に@で始まるハンドル名が見つかりませんでした。\n\n" +
        "📌 使い分け方法:\n" +
        "• 統合ダッシュボード: B8セルに1つのチャンネル入力\n" +
        "• ベンチマーク分析: B列に複数のハンドル名をリスト形式で入力\n\n" +
        "例: @YouTube",
        ui.ButtonSet.OK
      );
      return;
    }

    // 進捗状況を表示
    var successCount = 0;
    var errorCount = 0;

    for (var i = 0; i < handles.length; i++) {
      var handle = handles[i].handle;
      var row = handles[i].row;

      try {
        // YouTube APIを使ってチャンネル情報を取得
        var channelInfo = getChannelByHandle(handle, apiKey);

        if (channelInfo) {
          // 成功した場合、データをスプレッドシートに書き込む
          var snippet = channelInfo.snippet;
          var statistics = channelInfo.statistics;

          sheet.getRange(row, 3).setValue(snippet.title); // チャンネル名
          sheet
            .getRange(row, 4)
            .setValue(parseInt(statistics.subscriberCount).toLocaleString()); // 登録者数
          sheet
            .getRange(row, 5)
            .setValue(parseInt(statistics.viewCount).toLocaleString()); // 総視聴回数
          sheet
            .getRange(row, 6)
            .setValue(parseInt(statistics.videoCount).toLocaleString()); // 動画数
          sheet.getRange(row, 7).setValue(snippet.publishedAt); // 開設日
          sheet.getRange(row, 8).setValue(snippet.description.substring(0, 100)); // 説明（最初の100文字）
          sheet.getRange(row, 9).setValue(snippet.country || "不明"); // 国
          
          // サムネイル画像を挿入
          if (snippet.thumbnails && snippet.thumbnails.default) {
            try {
              var imageBlob = UrlFetchApp.fetch(snippet.thumbnails.default.url).getBlob();
              sheet.getRange(row, 10).clear(); // 既存の画像をクリア
              sheet.insertImage(imageBlob, row, 10);
            } catch (imageError) {
              sheet.getRange(row, 10).setValue("画像取得失敗");
            }
          }

          successCount++;
        } else {
          // 失敗した場合、エラーメッセージを表示
          sheet.getRange(row, 3).setValue("チャンネルが見つかりません");
          sheet.getRange(row, 4, 1, 7).setValue(""); // 他のセルをクリア
          errorCount++;
        }

        // 処理状況を更新
        var progress = Math.round(((i + 1) / handles.length) * 100);
        statusCell.setValue(progress + "%");
        SpreadsheetApp.flush(); // 画面を更新
      } catch (error) {
        Logger.log("チャンネル処理エラー(" + handle + "): " + error.toString());
        sheet.getRange(row, 3).setValue("処理エラー");
        sheet.getRange(row, 4, 1, 7).setValue(""); // 他のセルをクリア
        errorCount++;
      }
    }

    // 処理完了
    statusCell.setValue("完了");

    // 行の高さを調整（サムネイル画像用）
    for (var i = 0; i < handles.length; i++) {
      sheet.setRowHeight(handles[i].row, 60);
    }

    // 完了メッセージ
    ui.alert(
      "処理完了",
      "合計: " +
        handles.length +
        " 件\n" +
        "成功: " +
        successCount +
        " 件\n" +
        "失敗: " +
        errorCount +
        " 件\n\n" +
        "「YouTube ツール > ③ ベンチマークレポート作成」を実行するとレポートを生成できます",
      ui.ButtonSet.OK
    );
  } catch (error) {
    Logger.log("エラー: " + error.toString());
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var statusCell = sheet.getRange("L1");
    statusCell.setValue("エラー");

    SpreadsheetApp.getUi().alert(
      "エラー",
      "チャンネル情報の取得中にエラーが発生しました: " + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * チャンネル入力を正規化するヘルパー関数
 */
function normalizeChannelInputForProcess(input) {
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
 * ハンドル名からYouTubeチャンネル情報を取得する関数
 * @param {string} handle - ハンドル名（@username形式）
 * @param {string} apiKey - YouTube Data API キー
 * @return {Object|null} - チャンネル情報（取得失敗時はnull）
 */
function getChannelByHandle(handle, apiKey) {
  try {
    // ハンドル名から@を削除
    var username = handle.replace("@", "");
    var options = {
      method: "get",
      muteHttpExceptions: true,
    };

    // 1. 検索APIを使用してハンドル名で検索（そのままの形式で）
    var searchUrl =
      "https://www.googleapis.com/youtube/v3/search?part=snippet&q=" +
      encodeURIComponent(handle) +
      "&type=channel&maxResults=10&key=" +
      apiKey;

    var searchResponse = UrlFetchApp.fetch(searchUrl, options);
    var searchData = JSON.parse(searchResponse.getContentText());

    if (searchData && searchData.items && searchData.items.length > 0) {
      for (var i = 0; i < searchData.items.length; i++) {
        var channelId = searchData.items[i].id.channelId;

        // チャンネルの詳細情報を取得
        var channelUrl =
          "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=" +
          channelId +
          "&key=" +
          apiKey;

        var channelResponse = UrlFetchApp.fetch(channelUrl, options);
        var channelData = JSON.parse(channelResponse.getContentText());

        if (channelData && channelData.items && channelData.items.length > 0) {
          var channelDetails = channelData.items[0];

          // チャンネル名かcustomUrlがユーザー名を含むものを優先
          if (
            channelDetails.snippet.title
              .toLowerCase()
              .includes(username.toLowerCase()) ||
            (channelDetails.snippet.customUrl &&
              channelDetails.snippet.customUrl
                .toLowerCase()
                .includes(username.toLowerCase()))
          ) {
            return channelDetails;
          }
        }
      }

      // 名前に一致するものがなければ最初の結果を返す
      var firstChannelId = searchData.items[0].id.channelId;
      var firstChannelUrl =
        "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=" +
        firstChannelId +
        "&key=" +
        apiKey;

      var firstChannelResponse = UrlFetchApp.fetch(firstChannelUrl, options);
      var firstChannelData = JSON.parse(firstChannelResponse.getContentText());

      if (
        firstChannelData &&
        firstChannelData.items &&
        firstChannelData.items.length > 0
      ) {
        return firstChannelData.items[0];
      }
    }

    // 2. @なしのユーザー名で検索
    var usernameSearchUrl =
      "https://www.googleapis.com/youtube/v3/search?part=snippet&q=" +
      encodeURIComponent(username) +
      "&type=channel&maxResults=5&key=" +
      apiKey;

    var usernameSearchResponse = UrlFetchApp.fetch(usernameSearchUrl, options);
    var usernameSearchData = JSON.parse(
      usernameSearchResponse.getContentText()
    );

    if (
      usernameSearchData &&
      usernameSearchData.items &&
      usernameSearchData.items.length > 0
    ) {
      var channelId = usernameSearchData.items[0].id.channelId;
      var channelUrl =
        "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=" +
        channelId +
        "&key=" +
        apiKey;

      var channelResponse = UrlFetchApp.fetch(channelUrl, options);
      var channelData = JSON.parse(channelResponse.getContentText());

      if (channelData && channelData.items && channelData.items.length > 0) {
        return channelData.items[0];
      }
    }

    // 3. forUsername パラメータで検索 (旧式、バックアップとして)
    var usernameUrl =
      "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&forUsername=" +
      encodeURIComponent(username) +
      "&key=" +
      apiKey;

    var usernameResponse = UrlFetchApp.fetch(usernameUrl, options);
    var usernameData = JSON.parse(usernameResponse.getContentText());

    if (usernameData && usernameData.items && usernameData.items.length > 0) {
      return usernameData.items[0];
    }

    return null;
  } catch (error) {
    Logger.log("チャンネル取得エラー (" + handle + "): " + error.toString());
    return null;
  }
}

/**
 * ベンチマークレポートを作成する関数
 */
function createBenchmarkReport() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ui = SpreadsheetApp.getUi();

    // プログレスバーセルを用意
    var statusCell = sheet.getRange("L1");
    statusCell.setValue("レポート作成中...");

    // 既存データを取得
    var data = sheet.getDataRange().getValues();

    // ヘッダー行を除く有効なデータ行を確認
    var validRows = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2] !== "チャンネルが見つかりません") {
        validRows++;
      }
    }

    if (validRows === 0) {
      statusCell.setValue("エラー");
      ui.alert(
        "有効なデータがありません。先にチャンネル情報を取得してください。"
      );
      return;
    }

    statusCell.setValue("25%");

    // 現在日時からユニークなレポート名を生成
    var now = new Date();
    var reportName =
      "ベンチマークレポート_" +
      now.getFullYear() +
      ("0" + (now.getMonth() + 1)).slice(-2) +
      ("0" + now.getDate()).slice(-2) +
      "_" +
      ("0" + now.getHours()).slice(-2) +
      ("0" + now.getMinutes()).slice(-2);

    // 新しいシートを作成
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.insertSheet(reportName);

    // レポートヘッダーを設定
    reportSheet.appendRow(["YouTubeチャンネルベンチマークレポート"]);
    reportSheet.appendRow([
      "作成日: " + now.toLocaleDateString() + " " + now.toLocaleTimeString(),
    ]);
    reportSheet.appendRow(["チャンネル数: " + validRows + " 件"]);
    reportSheet.appendRow([""]);

    // 統計情報を計算
    var totalChannels = validRows;
    var totalSubscribers = 0;
    var totalViews = 0;
    var totalVideos = 0;

    for (var i = 1; i < data.length; i++) {
      // チャンネル名がある行のみ処理
      if (data[i][2] && data[i][2] !== "チャンネルが見つかりません") {
        var subscribers = data[i][4];
        if (subscribers !== "非公開" && subscribers) {
          // カンマを削除して数値に変換
          subscribers = subscribers.toString().replace(/,/g, "");
          if (!isNaN(parseInt(subscribers))) {
            totalSubscribers += parseInt(subscribers);
          }
        }

        var views = data[i][5] ? data[i][5].toString().replace(/,/g, "") : "0";
        if (!isNaN(parseInt(views))) {
          totalViews += parseInt(views);
        }

        var videos = data[i][6] ? data[i][6].toString().replace(/,/g, "") : "0";
        if (!isNaN(parseInt(videos))) {
          totalVideos += parseInt(videos);
        }
      }
    }

    // 進捗状況更新
    statusCell.setValue("50%");

    // 統計サマリーを追加
    reportSheet.appendRow(["統計サマリー:"]);
    reportSheet.appendRow(["指標", "合計", "平均", "最大値", "中央値"]);

    // 登録者数の統計データを計算
    var subscribersStats = calculateStats(data, 4);
    var viewsStats = calculateStats(data, 5);
    var videosStats = calculateStats(data, 6);

    // 統計データを追加
    reportSheet.appendRow([
      "登録者数",
      totalSubscribers.toLocaleString(),
      Math.round(subscribersStats.average).toLocaleString(),
      subscribersStats.max.toLocaleString(),
      subscribersStats.median.toLocaleString(),
    ]);

    reportSheet.appendRow([
      "総視聴回数",
      totalViews.toLocaleString(),
      Math.round(viewsStats.average).toLocaleString(),
      viewsStats.max.toLocaleString(),
      viewsStats.median.toLocaleString(),
    ]);

    reportSheet.appendRow([
      "動画本数",
      totalVideos.toLocaleString(),
      Math.round(videosStats.average).toLocaleString(),
      videosStats.max.toLocaleString(),
      videosStats.median.toLocaleString(),
    ]);

    reportSheet.appendRow([""]);

    // 進捗状況更新
    statusCell.setValue("70%");

    // トップ10チャンネルを追加
    reportSheet.appendRow(["登録者数トップ10チャンネル:"]);
    reportSheet.appendRow([
      "ランク",
      "ジャンル",
      "ハンドル名",
      "チャンネル名",
      "登録者数",
      "総視聴回数",
      "動画本数",
      "チャンネルURL",
      "サムネイル",
    ]);

    // データを登録者数でソート（非公開は除外）
    var sortableData = [];
    for (var i = 1; i < data.length; i++) {
      if (
        data[i][2] &&
        data[i][2] !== "チャンネルが見つかりません" &&
        data[i][4] !== "非公開"
      ) {
        var subscribers = data[i][4]
          ? data[i][4].toString().replace(/,/g, "")
          : "0";
        if (!isNaN(parseInt(subscribers))) {
          // IMAGE関数からサムネイルURLを抽出
          var thumbnailUrl = "";
          var imageFormula = data[i][9] ? data[i][9].toString() : "";
          if (imageFormula.indexOf('=IMAGE("') === 0) {
            thumbnailUrl = imageFormula.substring(
              8,
              imageFormula.indexOf('", 1')
            );
          }

          sortableData.push({
            genre: data[i][0], // ジャンル
            handle: data[i][1], // ハンドル名
            name: data[i][2], // チャンネル名
            subscribers: parseInt(subscribers),
            views: data[i][5] ? data[i][5].toString().replace(/,/g, "") : "0",
            videos: data[i][6] ? data[i][6].toString().replace(/,/g, "") : "0",
            url: data[i][8],
            thumbnail: thumbnailUrl,
          });
        }
      }
    }

    // 登録者数でソート
    sortableData.sort(function (a, b) {
      return b.subscribers - a.subscribers;
    });

    // トップ10を追加
    var maxRows = Math.min(10, sortableData.length);
    for (var i = 0; i < maxRows; i++) {
      reportSheet.appendRow([
        i + 1,
        sortableData[i].genre,
        sortableData[i].handle,
        sortableData[i].name,
        sortableData[i].subscribers.toLocaleString(),
        sortableData[i].views.toLocaleString(),
        sortableData[i].videos.toLocaleString(),
        sortableData[i].url,
        sortableData[i].thumbnail
          ? '=IMAGE("' + sortableData[i].thumbnail + '", 1)'
          : "",
      ]);
    }

    // サムネイル用に行の高さを調整
    for (var i = 0; i < maxRows; i++) {
      reportSheet.setRowHeight(13 + i, 60);
    }

    // 進捗状況更新
    statusCell.setValue("90%");

    // 平均以上の登録者数を持つチャンネル情報
    reportSheet.appendRow([""]);
    reportSheet.appendRow(["平均以上の登録者数を持つチャンネル:"]);
    reportSheet.appendRow([
      "ジャンル",
      "ハンドル名",
      "チャンネル名",
      "登録者数",
      "総視聴回数",
      "動画本数",
      "チャンネルURL",
      "サムネイル",
    ]);

    var avgSubscribers = subscribersStats.average;
    var aboveAvgChannels = sortableData.filter(function (channel) {
      return channel.subscribers >= avgSubscribers;
    });

    for (var i = 0; i < aboveAvgChannels.length; i++) {
      reportSheet.appendRow([
        aboveAvgChannels[i].genre,
        aboveAvgChannels[i].handle,
        aboveAvgChannels[i].name,
        aboveAvgChannels[i].subscribers.toLocaleString(),
        aboveAvgChannels[i].views.toLocaleString(),
        aboveAvgChannels[i].videos.toLocaleString(),
        aboveAvgChannels[i].url,
        aboveAvgChannels[i].thumbnail
          ? '=IMAGE("' + aboveAvgChannels[i].thumbnail + '", 1)'
          : "",
      ]);

      // サムネイル用に行の高さを調整
      reportSheet.setRowHeight(16 + maxRows + i, 60);
    }

    // ジャンル別の分析を追加
    var genreAnalysis = analyzeByGenre(sortableData);
    if (genreAnalysis.genres.length > 1) {
      // 複数のジャンルがある場合のみ
      reportSheet.appendRow([""]);
      reportSheet.appendRow(["ジャンル別分析:"]);
      reportSheet.appendRow([
        "ジャンル",
        "チャンネル数",
        "平均登録者数",
        "平均視聴回数",
        "平均動画数",
      ]);

      for (var i = 0; i < genreAnalysis.genres.length; i++) {
        var genre = genreAnalysis.genres[i];
        var stats = genreAnalysis.stats[genre];

        reportSheet.appendRow([
          genre,
          stats.count,
          Math.round(stats.avgSubscribers).toLocaleString(),
          Math.round(stats.avgViews).toLocaleString(),
          Math.round(stats.avgVideos).toLocaleString(),
        ]);
      }
    }

    // レポートの書式設定
    formatBenchmarkReport(
      reportSheet,
      maxRows,
      aboveAvgChannels.length,
      genreAnalysis.genres.length > 1
    );

    // レポートシートをアクティブにする
    ss.setActiveSheet(reportSheet);

    // 進捗状況更新
    statusCell.setValue("完了");

    ui.alert("ベンチマークレポートが作成されました。");
  } catch (error) {
    Logger.log("レポート作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "レポート作成中にエラーが発生しました: " + error.toString()
    );

    var progressCell = SpreadsheetApp.getActiveSpreadsheet()
      .getActiveSheet()
      .getRange("L1");
    if (progressCell) {
      progressCell.setValue("エラー");
    }
  }
}

/**
 * ジャンル別の統計を分析する関数
 */
function analyzeByGenre(sortableData) {
  var genreStats = {};
  var genres = [];

  // 各チャンネルのジャンルごとに集計
  for (var i = 0; i < sortableData.length; i++) {
    var genre = sortableData[i].genre || "未分類";

    if (!genreStats[genre]) {
      genreStats[genre] = {
        count: 0,
        totalSubscribers: 0,
        totalViews: 0,
        totalVideos: 0,
      };
      genres.push(genre);
    }

    genreStats[genre].count++;
    genreStats[genre].totalSubscribers += sortableData[i].subscribers;
    genreStats[genre].totalViews += parseInt(
      sortableData[i].views.replace(/,/g, "")
    );
    genreStats[genre].totalVideos += parseInt(
      sortableData[i].videos.replace(/,/g, "")
    );
  }

  // 平均値を計算
  for (var genre in genreStats) {
    genreStats[genre].avgSubscribers =
      genreStats[genre].totalSubscribers / genreStats[genre].count;
    genreStats[genre].avgViews =
      genreStats[genre].totalViews / genreStats[genre].count;
    genreStats[genre].avgVideos =
      genreStats[genre].totalVideos / genreStats[genre].count;
  }

  return {
    genres: genres,
    stats: genreStats,
  };
}

/**
 * 統計値を計算する関数
 * @param {Array} data - 元データ配列
 * @param {number} colIndex - 計算対象の列インデックス
 * @return {Object} 統計値を含むオブジェクト
 */
function calculateStats(data, colIndex) {
  var values = [];

  // 有効な数値を抽出
  for (var i = 1; i < data.length; i++) {
    if (
      data[i][2] &&
      data[i][2] !== "チャンネルが見つかりません" &&
      data[i][colIndex] !== "非公開"
    ) {
      var value = data[i][colIndex]
        ? data[i][colIndex].toString().replace(/,/g, "")
        : "0";
      if (!isNaN(parseInt(value))) {
        values.push(parseInt(value));
      }
    }
  }

  // 値がない場合は0を返す
  if (values.length === 0) {
    return {
      average: 0,
      max: 0,
      median: 0,
    };
  }

  // 昇順ソート
  values.sort(function (a, b) {
    return a - b;
  });

  // 平均値計算
  var sum = values.reduce(function (a, b) {
    return a + b;
  }, 0);
  var average = sum / values.length;

  // 最大値
  var max = values[values.length - 1];

  // 中央値計算
  var median;
  var middle = Math.floor(values.length / 2);

  if (values.length % 2 === 0) {
    median = (values[middle - 1] + values[middle]) / 2;
  } else {
    median = values[middle];
  }

  return {
    average: average,
    max: max,
    median: median,
  };
}

/**
 * レポートの書式設定を行う関数
 */
function formatBenchmarkReport(
  sheet,
  topChannelsCount,
  aboveAvgCount,
  hasGenreAnalysis
) {
  // タイトルと日付の書式設定
  sheet.getRange("A1:I1").merge();
  sheet.getRange("A1").setValue("📊 YouTube チャンネル ベンチマークレポート")
    .setFontSize(20).setFontWeight("bold")
    .setBackground("#f8f9fa").setFontColor("#495057");
  sheet.getRange("A2:I2").merge();
  sheet.getRange("A2").setFontStyle("italic").setHorizontalAlignment("center");

  sheet.getRange("A3:I3").merge();
  sheet.getRange("A3").setFontWeight("bold").setHorizontalAlignment("center");

  // 統計サマリーの書式設定
  sheet.getRange("A5").setFontWeight("bold");
  sheet.getRange("A6:E6").setFontWeight("bold").setBackground("#E0E0E0");
  sheet.getRange("A7:E9").setBorder(true, true, true, true, true, true);
  sheet.getRange("A6:E9").setHorizontalAlignment("center");

  // トップ10チャンネルの書式設定
  sheet.getRange("A11").setFontWeight("bold");
  sheet.getRange("A12:I12").setFontWeight("bold").setBackground("#E0E0E0");
  sheet
    .getRange("A13:I" + (12 + topChannelsCount))
    .setBorder(true, true, true, true, true, true);

  // 平均以上の書式設定
  var aboveAvgStartRow = 14 + topChannelsCount;
  sheet.getRange("A" + aboveAvgStartRow).setFontWeight("bold");
  sheet
    .getRange("A" + (aboveAvgStartRow + 1) + ":H" + (aboveAvgStartRow + 1))
    .setFontWeight("bold")
    .setBackground("#E0E0E0");
  sheet
    .getRange(
      "A" +
        (aboveAvgStartRow + 2) +
        ":H" +
        (aboveAvgStartRow + 1 + aboveAvgCount)
    )
    .setBorder(true, true, true, true, true, true);

  // ジャンル分析の書式設定（存在する場合）
  if (hasGenreAnalysis) {
    var genreStartRow = aboveAvgStartRow + aboveAvgCount + 3;
    sheet.getRange("A" + genreStartRow).setFontWeight("bold");
    sheet
      .getRange("A" + (genreStartRow + 1) + ":E" + (genreStartRow + 1))
      .setFontWeight("bold")
      .setBackground("#E0E0E0");

    // ジャンル行数を取得（ヘッダー+データ行）
    var genreCount = 0;
    for (var i = genreStartRow + 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange("A" + i).getValue()) {
        genreCount++;
      }
    }

    if (genreCount > 0) {
      sheet
        .getRange(
          "A" + (genreStartRow + 2) + ":E" + (genreStartRow + 1 + genreCount)
        )
        .setBorder(true, true, true, true, true, true);
    }
  }

  // 列幅の調整
  sheet.setColumnWidth(1, 80); // ランク/ジャンル
  sheet.setColumnWidth(2, 120); // ジャンル/ハンドル名
  sheet.setColumnWidth(3, 150); // ハンドル名/チャンネル名
  sheet.setColumnWidth(4, 200); // チャンネル名/登録者数
  sheet.setColumnWidth(5, 100); // 登録者数/視聴回数
  sheet.setColumnWidth(6, 120); // 視聴回数/動画数
  sheet.setColumnWidth(7, 100); // 動画数/URL
  sheet.setColumnWidth(8, 250); // URL
  sheet.setColumnWidth(9, 150); // サムネイル

  // チャンネルURLを青色のハイパーリンクに
  var urls = sheet.getRange("H13:H" + (12 + topChannelsCount)).getValues();
  for (var i = 0; i < urls.length; i++) {
    if (urls[i][0]) {
      var range = sheet.getRange("H" + (13 + i));
      range.setFontColor("#1155CC").setFontLine("underline");
    }
  }

  var aboveAvgUrls = sheet
    .getRange(
      "G" +
        (aboveAvgStartRow + 2) +
        ":G" +
        (aboveAvgStartRow + 1 + aboveAvgCount)
    )
    .getValues();
  for (var i = 0; i < aboveAvgUrls.length; i++) {
    if (aboveAvgUrls[i][0]) {
      var range = sheet.getRange("G" + (aboveAvgStartRow + 2 + i));
      range.setFontColor("#1155CC").setFontLine("underline");
    }
  }
}

/**
 * 既存チャンネルの個別分析を実行
 */
function analyzeExistingChannel() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // APIキーの確認
    var apiKey = PropertiesService.getScriptProperties().getProperty("YOUTUBE_API_KEY");
    if (!apiKey) {
      ui.alert(
        "API設定が必要",
        "先に「① API設定」を実行してAPIキーを設定してください。",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // チャンネル入力を求める
    var response = ui.prompt(
      "個別チャンネル分析",
      "分析したいチャンネルのハンドル名またはURLを入力してください:\n\n" +
      "例: @YouTube\n" +
      "例: https://www.youtube.com/@YouTube",
      ui.ButtonSet.OK_CANCEL
    );
    
    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }
    
    var input = response.getResponseText().trim();
    if (!input) {
      ui.alert("入力エラー", "チャンネル情報を入力してください。", ui.ButtonSet.OK);
      return;
    }
    
    // 入力を正規化
    var normalizedInput = normalizeChannelInputForProcess(input);
    if (!normalizedInput) {
      ui.alert(
        "入力形式エラー", 
        "正しいYouTubeチャンネルURL または @ハンドル名を入力してください。",
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 分析実行の確認
    var confirmResponse = ui.alert(
      "個別分析実行",
      "チャンネル「" + normalizedInput + "」の詳細分析を実行しますか？\n\n" +
      "🔍 実行内容:\n" +
      "• 基本チャンネル情報取得\n" +
      "• 詳細統計情報分析\n" +
      "• 個別分析レポート作成",
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      return;
    }
    
    // 分析実行
    ui.alert("分析開始", "チャンネル分析を開始します。完了まで少々お待ちください。", ui.ButtonSet.OK);
    
    // チャンネル情報を取得
    var channelInfo = getChannelByHandle(normalizedInput, apiKey);
    if (!channelInfo) {
      ui.alert(
        "チャンネル未発見",
        "指定されたチャンネルが見つかりませんでした: " + normalizedInput,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 分析レポートシートを作成
    createIndividualAnalysisReport(channelInfo, normalizedInput);
    
    ui.alert(
      "分析完了",
      "個別チャンネル分析が完了しました！\n\n" +
      "📊 チャンネル名: " + channelInfo.snippet.title + "\n" +
      "📈 登録者数: " + parseInt(channelInfo.statistics.subscriberCount).toLocaleString() + "\n" +
      "🎬 動画数: " + channelInfo.statistics.videoCount + "\n\n" +
      "詳細レポートシートが作成されました。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("個別チャンネル分析エラー: " + error.toString());
    ui.alert(
      "分析エラー",
      "個別チャンネル分析中にエラーが発生しました:\n\n" + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * 個別分析レポートを作成
 */
function createIndividualAnalysisReport(channelInfo, originalInput) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timestamp = new Date();
    var reportName = "🔍 " + channelInfo.snippet.title + "_" + 
                    timestamp.getFullYear() + 
                    ("0" + (timestamp.getMonth() + 1)).slice(-2) +
                    ("0" + timestamp.getDate()).slice(-2);
    
    // 既存のレポートシートがあれば削除
    var existingSheet = ss.getSheetByName(reportName);
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    var reportSheet = ss.insertSheet(reportName);
    
    // レポートヘッダー
    reportSheet.getRange("A1:H1").merge();
    reportSheet.getRange("A1").setValue("🔍 個別チャンネル分析レポート")
      .setFontSize(18).setFontWeight("bold")
      .setBackground("#1976D2").setFontColor("white")
      .setHorizontalAlignment("center");
    
    reportSheet.getRange("A2:H2").merge();
    reportSheet.getRange("A2").setValue("作成日時: " + timestamp.toLocaleString())
      .setHorizontalAlignment("center").setFontStyle("italic");
    
    // 基本情報セクション
    reportSheet.getRange("A4:H4").merge();
    reportSheet.getRange("A4").setValue("📊 基本チャンネル情報")
      .setFontSize(14).setFontWeight("bold")
      .setBackground("#4CAF50").setFontColor("white")
      .setHorizontalAlignment("center");
    
    var snippet = channelInfo.snippet;
    var statistics = channelInfo.statistics;
    
    var basicInfo = [
      ["チャンネル名", snippet.title],
      ["チャンネルID", channelInfo.id],
      ["入力値", originalInput],
      ["登録者数", parseInt(statistics.subscriberCount).toLocaleString()],
      ["総視聴回数", parseInt(statistics.viewCount).toLocaleString()],
      ["動画本数", statistics.videoCount],
      ["チャンネル開設日", snippet.publishedAt],
      ["国/地域", snippet.country || "不明"],
      ["説明", snippet.description.substring(0, 200) + (snippet.description.length > 200 ? "..." : "")]
    ];
    
    for (var i = 0; i < basicInfo.length; i++) {
      reportSheet.getRange(5 + i, 1).setValue(basicInfo[i][0]).setFontWeight("bold");
      reportSheet.getRange(5 + i, 2, 1, 7).merge();
      reportSheet.getRange(5 + i, 2).setValue(basicInfo[i][1]);
    }
    
    // 分析指標セクション
    var analyticsRow = 5 + basicInfo.length + 2;
    reportSheet.getRange("A" + analyticsRow + ":H" + analyticsRow).merge();
    reportSheet.getRange("A" + analyticsRow).setValue("📈 分析指標")
      .setFontSize(14).setFontWeight("bold")
      .setBackground("#FF9800").setFontColor("white")
      .setHorizontalAlignment("center");
    
    analyticsRow++;
    
    // 計算された指標
    var subscribers = parseInt(statistics.subscriberCount);
    var totalViews = parseInt(statistics.viewCount);
    var videoCount = parseInt(statistics.videoCount);
    var avgViewsPerVideo = videoCount > 0 ? Math.round(totalViews / videoCount) : 0;
    var viewsPerSubscriber = subscribers > 0 ? (totalViews / subscribers).toFixed(2) : 0;
    var createdDate = new Date(snippet.publishedAt);
    var daysSinceCreated = Math.floor((new Date() - createdDate) / (1000 * 60 * 60 * 24));
    var yearsActive = (daysSinceCreated / 365).toFixed(1);
    var subscribersPerYear = yearsActive > 0 ? Math.round(subscribers / parseFloat(yearsActive)) : 0;
    
    var analytics = [
      ["動画あたり平均視聴回数", avgViewsPerVideo.toLocaleString()],
      ["登録者あたり視聴回数", viewsPerSubscriber],
      ["チャンネル運営年数", yearsActive + "年"],
      ["年間平均登録者数増加", subscribersPerYear.toLocaleString()],
      ["動画投稿頻度", videoCount > 0 ? (videoCount / parseFloat(yearsActive)).toFixed(1) + "本/年" : "0本/年"]
    ];
    
    for (var i = 0; i < analytics.length; i++) {
      reportSheet.getRange(analyticsRow + i, 1).setValue(analytics[i][0]).setFontWeight("bold");
      reportSheet.getRange(analyticsRow + i, 2, 1, 7).merge();
      reportSheet.getRange(analyticsRow + i, 2).setValue(analytics[i][1]);
    }
    
    // サムネイル画像
    if (snippet.thumbnails && snippet.thumbnails.high) {
      try {
        reportSheet.getRange("F5").setValue("チャンネルサムネイル:");
        reportSheet.getRange("G5:H7").merge();
        var imageBlob = UrlFetchApp.fetch(snippet.thumbnails.high.url).getBlob();
        reportSheet.insertImage(imageBlob, 5, 7);
      } catch (imageError) {
        reportSheet.getRange("G5").setValue("画像取得失敗");
      }
    }
    
    // フォーマット設定
    formatIndividualReport(reportSheet);
    
    // レポートシートをアクティブに
    ss.setActiveSheet(reportSheet);
    
  } catch (error) {
    Logger.log("個別レポート作成エラー: " + error.toString());
    throw error;
  }
}

/**
 * 個別レポートのフォーマット設定
 */
function formatIndividualReport(sheet) {
  // 列幅設定
  sheet.setColumnWidth(1, 150);
  for (var i = 2; i <= 8; i++) {
    sheet.setColumnWidth(i, 100);
  }
  
  // 行高設定
  sheet.setRowHeight(1, 40);
  sheet.setRowHeight(2, 25);
  
  // ボーダー設定
  sheet.getRange("A5:B13").setBorder(true, true, true, true, true, true);
  sheet.getRange("A16:B20").setBorder(true, true, true, true, true, true);
  
  // 背景色設定
  sheet.getRange("A5:A13").setBackground("#E8F5E8");
  sheet.getRange("A16:A20").setBackground("#FFF3E0");
}

/**
 * ベンチマーク分析ダッシュボードを表示
 */
function showBenchmarkDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardName = "🔍 ベンチマーク分析";
    var dashboard = ss.getSheetByName(dashboardName);
    
    if (!dashboard) {
      dashboard = ss.insertSheet(dashboardName);
      setupBenchmarkAnalysisDashboard(dashboard);
    }
    
    ss.setActiveSheet(dashboard);
    updateBenchmarkAnalysisData(dashboard);
    
  } catch (error) {
    Logger.log("ベンチマーク分析ダッシュボード表示エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "ダッシュボード表示エラー",
      "ベンチマーク分析ダッシュボードの表示中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * ベンチマーク分析ダッシュボードをセットアップ
 */
function setupBenchmarkAnalysisDashboard(dashboard) {
  // ヘッダー
  dashboard.getRange("A1:J1").merge();
  dashboard.getRange("A1").setValue("🔍 YouTube ベンチマーク分析ダッシュボード")
    .setFontSize(18).setFontWeight("bold")
    .setBackground("#e3f2fd").setFontColor("#1565c0")
    .setHorizontalAlignment("center");
    
  dashboard.getRange("A2:J2").merge();
  dashboard.getRange("A2").setValue("取得済みチャンネルデータの統計分析・比較・ランキング表示")
    .setFontSize(12).setBackground("#f5f5f5")
    .setHorizontalAlignment("center");
    
  // データサマリー
  dashboard.getRange("A4:J4").merge();
  dashboard.getRange("A4").setValue("📊 データサマリー")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#e8f5e8").setFontColor("#2e7d32")
    .setHorizontalAlignment("center");
    
  var summaryLabels = [
    ["分析対象チャンネル数:", "取得中..."],
    ["平均登録者数:", "計算中..."],
    ["総視聴回数:", "計算中..."],
    ["最新更新日:", new Date().toLocaleDateString()]
  ];
  
  for (var i = 0; i < summaryLabels.length; i++) {
    dashboard.getRange(5 + i, 1, 1, 2).merge();
    dashboard.getRange(5 + i, 1).setValue(summaryLabels[i][0]).setFontWeight("bold");
    dashboard.getRange(5 + i, 3, 1, 2).merge();
    dashboard.getRange(5 + i, 3).setValue(summaryLabels[i][1]);
  }
  
  // トップチャンネルランキング
  dashboard.getRange("A10:J10").merge();
  dashboard.getRange("A10").setValue("🏆 登録者数トップランキング")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#fff3e0").setFontColor("#f57c00")
    .setHorizontalAlignment("center");
    
  var rankingHeaders = ["順位", "チャンネル名", "ハンドル", "登録者数", "視聴回数", "動画数", "カテゴリ"];
  for (var i = 0; i < rankingHeaders.length; i++) {
    dashboard.getRange(11, i + 1).setValue(rankingHeaders[i]).setFontWeight("bold")
      .setBackground("#e8f5e8").setHorizontalAlignment("center");
  }
  
  // 分析アクション
  dashboard.getRange("A18:J18").merge();
  dashboard.getRange("A18").setValue("⚡ 分析アクション")
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#f3e5f5").setFontColor("#7b1fa2")
    .setHorizontalAlignment("center");
    
  var analysisActions = [
    ["📈", "詳細レポート生成", "現在のデータから詳細ベンチマークレポートを作成", "createBenchmarkReport"],
    ["🔄", "データ更新", "全チャンネルの情報を最新データで更新", "processHandles"],
    ["➕", "チャンネル追加", "新しいチャンネルを分析対象に追加", "addNewChannelForAnalysis"],
    ["📊", "統計分析", "詳細な統計分析を実行", "runDetailedStatistics"]
  ];
  
  for (var i = 0; i < analysisActions.length; i++) {
    var row = 19 + i;
    dashboard.getRange(row, 1).setValue(analysisActions[i][0]);
    dashboard.getRange(row, 2, 1, 2).merge();
    dashboard.getRange(row, 2).setValue(analysisActions[i][1]).setFontWeight("bold");
    dashboard.getRange(row, 4, 1, 5).merge();
    dashboard.getRange(row, 4).setValue(analysisActions[i][2]);
    dashboard.getRange(row, 9).setValue("▶ 実行")
      .setBackground("#4caf50").setFontColor("white").setFontWeight("bold")
      .setHorizontalAlignment("center");
    dashboard.getRange(row, 10).setValue(analysisActions[i][3]); // 関数名
  }
  
  // フォーマット
  formatBenchmarkAnalysisDashboard(dashboard);
}

/**
 * ベンチマーク分析ダッシュボードのフォーマット設定
 */
function formatBenchmarkAnalysisDashboard(dashboard) {
  // 列幅設定
  dashboard.setColumnWidth(1, 60);  // アイコン
  dashboard.setColumnWidth(2, 120); // チャンネル名
  dashboard.setColumnWidth(3, 100); // ハンドル
  dashboard.setColumnWidth(4, 100); // 登録者数
  dashboard.setColumnWidth(5, 120); // 視聴回数
  dashboard.setColumnWidth(6, 80);  // 動画数
  dashboard.setColumnWidth(7, 100); // カテゴリ
  dashboard.setColumnWidth(8, 80);  // 予備
  dashboard.setColumnWidth(9, 80);  // ボタン
  dashboard.setColumnWidth(10, 10); // 関数名（非表示）
  
  // 行高設定
  dashboard.setRowHeight(1, 40);
  dashboard.setRowHeight(2, 25);
  
  // ボーダー設定
  dashboard.getRange("A11:G17").setBorder(true, true, true, true, true, true);
  dashboard.getRange("A19:I22").setBorder(true, true, true, true, true, true);
  
  // 列を非表示
  dashboard.hideColumns(10);
}

/**
 * ベンチマーク分析データを更新
 */
function updateBenchmarkAnalysisData(dashboard) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var channelData = [];
    
    // 全シートからチャンネルデータを収集
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      var sheetName = sheet.getName();
      
      // ダッシュボード、レポート、ガイドシートをスキップ
      if (sheetName.includes("ダッシュボード") || sheetName.includes("レポート") || 
          sheetName.includes("ガイド") || sheetName.includes("分析")) {
        continue;
      }
      
      try {
        var data = sheet.getDataRange().getValues();
        for (var j = 1; j < data.length; j++) { // ヘッダー行をスキップ
          if (data[j][1] && data[j][1].toString().startsWith("@") && 
              data[j][2] && data[j][2] !== "チャンネルが見つかりません") {
            
            var subscribers = data[j][4] ? data[j][4].toString().replace(/,/g, "") : "0";
            var views = data[j][5] ? data[j][5].toString().replace(/,/g, "") : "0";
            var videos = data[j][6] ? data[j][6].toString().replace(/,/g, "") : "0";
            
            if (!isNaN(parseInt(subscribers))) {
              channelData.push({
                category: data[j][0] || "未分類",
                handle: data[j][1],
                name: data[j][2],
                subscribers: parseInt(subscribers),
                views: parseInt(views),
                videos: parseInt(videos)
              });
            }
          }
        }
      } catch (e) {
        Logger.log("シートデータ読み込みエラー: " + sheetName + " - " + e.toString());
      }
    }
    
    // データサマリーを更新
    dashboard.getRange("C5").setValue(channelData.length + "件");
    
    if (channelData.length > 0) {
      // 平均登録者数計算
      var totalSubscribers = channelData.reduce(function(sum, channel) {
        return sum + channel.subscribers;
      }, 0);
      var avgSubscribers = Math.round(totalSubscribers / channelData.length);
      dashboard.getRange("C6").setValue(avgSubscribers.toLocaleString());
      
      // 総視聴回数計算
      var totalViews = channelData.reduce(function(sum, channel) {
        return sum + channel.views;
      }, 0);
      dashboard.getRange("C7").setValue(totalViews.toLocaleString());
      
      // 登録者数でソート
      channelData.sort(function(a, b) {
        return b.subscribers - a.subscribers;
      });
      
      // トップ5を表示
      var topCount = Math.min(5, channelData.length);
      for (var i = 0; i < topCount; i++) {
        var row = 12 + i;
        var channel = channelData[i];
        
        dashboard.getRange(row, 1).setValue(i + 1);
        dashboard.getRange(row, 2).setValue(channel.name);
        dashboard.getRange(row, 3).setValue(channel.handle);
        dashboard.getRange(row, 4).setValue(channel.subscribers.toLocaleString());
        dashboard.getRange(row, 5).setValue(channel.views.toLocaleString());
        dashboard.getRange(row, 6).setValue(channel.videos.toLocaleString());
        dashboard.getRange(row, 7).setValue(channel.category);
      }
      
      // 空行をクリア
      for (var i = topCount; i < 5; i++) {
        var row = 12 + i;
        dashboard.getRange(row, 1, 1, 7).clearContent();
      }
      
    } else {
      dashboard.getRange("C6").setValue("データなし");
      dashboard.getRange("C7").setValue("データなし");
      dashboard.getRange("A12:G16").clearContent();
    }
    
    dashboard.getRange("C8").setValue(new Date().toLocaleString());
    
  } catch (error) {
    Logger.log("ベンチマーク分析データ更新エラー: " + error.toString());
  }
}

/**
 * 新しいチャンネルを分析対象に追加
 */
function addNewChannelForAnalysis() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    "チャンネル追加",
    "追加したいチャンネルのハンドル名またはURLを入力してください:\n\n" +
    "例: @YouTube\n" +
    "例: https://www.youtube.com/@YouTube",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    var input = response.getResponseText().trim();
    if (input) {
      // メインのprocessHandles関数にリダイレクト
      ui.alert(
        "チャンネル追加",
        "チャンネル「" + input + "」を追加します。\n\n" +
        "データシートのB列に追加して「② チャンネル取得」を実行してください。",
        ui.ButtonSet.OK
      );
    }
  }
}

/**
 * 詳細統計分析を実行
 */
function runDetailedStatistics() {
  var ui = SpreadsheetApp.getUi();
  ui.alert(
    "詳細統計分析",
    "詳細統計分析機能は「③ ベンチマークレポート作成」で利用できます。\n\n" +
    "より詳細な統計情報とグラフが含まれたレポートが作成されます。",
    ui.ButtonSet.OK
  );
}

/**
 * 事業分析ダッシュボード専用のチャンネル分析実行
 */
function executeBusinessChannelAnalysis(handle, apiKey, dashboard) {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 分析開始の確認
    var response = ui.alert(
      "🚀 チャンネル分析実行",
      "チャンネル「" + handle + "」の事業分析を実行しますか？\n\n" +
      "📊 実行内容:\n" +
      "• 基本チャンネル情報取得\n" +
      "• 事業指標分析\n" +
      "• ダッシュボード更新\n\n" +
      "⚡ より詳細な分析は「🚀 包括事業分析」メニューから実行できます。",
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 進捗表示
    dashboard.getRange("E16").setValue("分析中...");
    SpreadsheetApp.flush();
    
    // チャンネル情報を取得
    var channelInfo = getChannelByHandle(handle, apiKey);
    if (!channelInfo) {
      ui.alert(
        "チャンネル未発見",
        "指定されたチャンネルが見つかりませんでした: " + handle,
        ui.ButtonSet.OK
      );
      dashboard.getRange("E16").setValue("エラー");
      return;
    }
    
    // 基本情報を抽出
    var snippet = channelInfo.snippet;
    var statistics = channelInfo.statistics;
    var channelName = snippet.title;
    var subscribers = parseInt(statistics.subscriberCount || 0);
    var totalViews = parseInt(statistics.viewCount || 0);
    var videoCount = parseInt(statistics.videoCount || 0);
    
    // 分析指標を計算
    var avgViews = videoCount > 0 ? Math.round(totalViews / videoCount) : 0;
    var engagementRate = subscribers > 0 ? ((avgViews / subscribers) * 100).toFixed(2) : 0;
    
    // 事業ステージを判定
    var businessStage = "🌱 新興";
    if (subscribers >= 1000000) businessStage = "🏆 業界リーダー";
    else if (subscribers >= 100000) businessStage = "🌟 確立企業";
    else if (subscribers >= 10000) businessStage = "📈 成長企業";
    else if (subscribers >= 1000) businessStage = "🚀 スタートアップ";
    
    // ダッシュボードを更新
    updateBusinessDashboardResults(dashboard, {
      channelName: channelName,
      subscribers: subscribers,
      totalViews: totalViews,
      videoCount: videoCount,
      avgViews: avgViews,
      engagementRate: engagementRate,
      businessStage: businessStage
    });
    
    dashboard.getRange("E16").setValue("完了");
    
    ui.alert(
      "✅ 分析完了",
      "チャンネル分析が完了しました！\n\n" +
      "📊 " + channelName + "\n" +
      "👥 登録者数: " + subscribers.toLocaleString() + "\n" +
      "📈 エンゲージメント率: " + engagementRate + "%\n" +
      "🎯 事業ステージ: " + businessStage + "\n\n" +
      "ダッシュボードが更新されました。\n" +
      "より詳細な分析は「🚀 包括事業分析」メニューをご利用ください。",
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("事業分析エラー: " + error.toString());
    dashboard.getRange("E16").setValue("エラー");
    ui.alert(
      "分析エラー",
      "チャンネル分析中にエラーが発生しました:\n\n" + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * 事業ダッシュボードの結果更新
 */
function updateBusinessDashboardResults(dashboard, results) {
  try {
    // 基本情報更新（A12行）
    var basicData = [
      results.channelName,
      results.subscribers.toLocaleString(),
      results.totalViews.toLocaleString(),
      results.videoCount.toLocaleString(),
      results.avgViews.toLocaleString(),
      results.engagementRate + "%",
      results.businessStage
    ];
    
    dashboard.getRange("A12:G12").setValues([basicData]);
    
    // 事業KPI更新（A16行）
    var monetizationStatus = results.subscribers >= 1000 ? "✅ 収益化対象" : "❌ 収益化前";
    var estimatedRevenue = results.subscribers >= 1000 ? Math.round(results.avgViews * 0.002) + "円/月" : "未収益化";
    var growthRate = "中程度"; // 簡易判定
    var marketPosition = results.subscribers >= 100000 ? "上位" : 
                        results.subscribers >= 10000 ? "中位" : "新興";
    var competitiveAdvantage = results.engagementRate > 3 ? "高エンゲージメント" : "標準";
    var businessScore = Math.min(100, Math.round(
      (results.subscribers / 1000) * 0.3 + 
      (results.engagementRate * 10) + 
      (results.avgViews / 1000) * 0.1
    ));
    
    var businessData = [
      monetizationStatus,
      estimatedRevenue,
      growthRate,
      marketPosition,
      competitiveAdvantage,
      businessScore + "/100"
    ];
    
    dashboard.getRange("A16:F16").setValues([businessData]);
    
    // 現在のデータ件数を更新
    dashboard.getRange("B16").setValue("1件");
    dashboard.getRange("E5").setValue(new Date().toLocaleString());
    
  } catch (error) {
    Logger.log("ダッシュボード更新エラー: " + error.toString());
  }
}

/**
 * ダッシュボードを強制的に再作成（新しい色設定適用）
 */
function refreshDashboardColors() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dashboardName = "📊 YouTube ベンチマーク管理";
    
    // 既存のダッシュボードを削除
    var existingDashboard = ss.getSheetByName(dashboardName);
    if (existingDashboard) {
      ss.deleteSheet(existingDashboard);
    }
    
    // 新しいダッシュボードを作成
    var newDashboard = ss.insertSheet(dashboardName, 0);
    setupBenchmarkDashboard(newDashboard);
    ss.setActiveSheet(newDashboard);
    updateBenchmarkDashboardStatus(newDashboard);
    
    SpreadsheetApp.getUi().alert(
      "ダッシュボード更新完了",
      "新しい色設定でダッシュボードを再作成しました。",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    Logger.log("ダッシュボード再作成エラー: " + error.toString());
    SpreadsheetApp.getUi().alert(
      "エラー",
      "ダッシュボードの再作成中にエラーが発生しました: " + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
