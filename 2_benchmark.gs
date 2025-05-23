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

// スプレッドシートのUIにメニューを追加
function onOpen_backup() {
  var ui = SpreadsheetApp.getUi();

  // メインメニュー
  var menu = ui.createMenu("YouTube ツール");

  // 統合ダッシュボードを最優先で表示
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
  menu.addToUi();

  // 初回実行時または統合ダッシュボードがない場合は作成
  createOrShowMainDashboard();
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
    helpSheet.getRange("A1:C1").setBackground("#4285F4").setFontColor("white");
    helpSheet.getRange("A3:C3").setBackground("#E0E0E0");
    helpSheet.getRange("A9:C9").setBackground("#E0E0E0");

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
    if (sheetName === "📊 YouTube チャンネル分析") {
      var channelInput = sheet.getRange("B8").getValue();
      
      if (!channelInput || channelInput.toString().trim() === "" || channelInput === "チャンネルURL or @ハンドル") {
        ui.alert(
          "入力エラー",
          "統合ダッシュボードのB8セルにチャンネルURLまたは@ハンドル名を入力してください。\n\n例: https://www.youtube.com/@YouTube\nまたは: @YouTube",
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
      
      // 統合ダッシュボード専用の処理を実行
      executeChannelAnalysis();
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
