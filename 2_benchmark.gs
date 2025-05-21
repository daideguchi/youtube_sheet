/**
 * YouTubeベンチマークチャンネルリスト
 * ハンドル名（@username形式）を入力することで、YouTubeチャンネルの情報を表形式で取得
 * 
 * 作成者: Claude AI
 * バージョン: 5.3
 * 最終更新: 2025-05-22
 */

// スプレッドシートのUIにメニューを追加
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // メインメニュー
  var menu = ui.createMenu('YouTube ツール');
  
  // シンプル化したメニュー構成
  menu.addItem('① API設定・テスト', 'setApiKey');
  menu.addItem('② チャンネル情報取得', 'processHandles');
  menu.addItem('③ ベンチマークレポート作成', 'createBenchmarkReport');
  menu.addSeparator();
  menu.addItem('シートテンプレート作成', 'setupBasicSheet');
  menu.addItem('使い方ガイドを表示', 'showHelpSheet');
  menu.addToUi();
  
  // 初回実行時は基本的な説明を表示
  showInitialGuide();
}

/**
 * 初回実行時のガイド表示
 */
function showInitialGuide() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var props = PropertiesService.getDocumentProperties();
    
    // 初回表示フラグをチェック
    if (!props.getProperty('initialGuideShown')) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        'YouTube ベンチマークツールへようこそ！',
        '基本的な使い方：\n\n' +
        '1. ハンドル名（@から始まる）を入力\n' +
        '2. 「YouTube ツール > ② チャンネル情報取得」を実行\n' +
        '3. 「YouTube ツール > ③ ベンチマークレポート作成」を実行\n\n' +
        'より詳しいガイドはメニューの「使い方ガイドを表示」から確認できます。\n\n' +
        '準備を始めますか？',
        ui.ButtonSet.YES_NO
      );
      
      // 初回表示フラグを設定
      props.setProperty('initialGuideShown', 'true');
      
      // はいを選択した場合
      if (response == ui.Button.YES) {
        setupBasicSheet();
        setApiKey();
      }
    }
  } catch (error) {
    Logger.log('初期ガイド表示エラー: ' + error.toString());
  }
}

/**
 * 使い方ガイドを専用シートに表示
 */
function showHelpSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var existingSheet = ss.getSheetByName('使い方ガイド');
    
    // 既存のガイドシートがあれば削除
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
    }
    
    // 新しいガイドシートを作成
    var helpSheet = ss.insertSheet('使い方ガイド');
    
    // ガイドコンテンツを設定
    helpSheet.appendRow(['YouTube チャンネル分析ツール - 使い方ガイド']);
    helpSheet.appendRow(['']);
    helpSheet.appendRow(['基本的な使い方']);
    helpSheet.appendRow(['1.', 'API設定', '「YouTube ツール > ① API設定・テスト」からYouTube Data APIキーを設定します']);
    helpSheet.appendRow(['2.', 'ハンドル入力', 'B列に@から始まるYouTubeハンドル名を入力します（例: @YouTube）']);
    helpSheet.appendRow(['3.', '情報取得', '「YouTube ツール > ② チャンネル情報取得」を選択']);
    helpSheet.appendRow(['4.', 'レポート作成', '「YouTube ツール > ③ ベンチマークレポート作成」を選択']);
    helpSheet.appendRow(['']);
    helpSheet.appendRow(['ヒント']);
    helpSheet.appendRow(['・', 'ジャンル分類', 'A列にはジャンルなど任意の分類を入力できます（レポートに反映されます）']);
    helpSheet.appendRow(['・', 'API取得上限', '1日あたりのAPI呼び出し上限があります。多数のチャンネルを分析する場合は注意してください']);
    helpSheet.appendRow(['・', 'データ更新', '情報を再取得する場合は、再度「② チャンネル情報取得」を実行してください']);
    
    // 書式設定
    helpSheet.getRange('A1').setFontSize(16).setFontWeight('bold');
    helpSheet.getRange('A3').setFontSize(14).setFontWeight('bold');
    helpSheet.getRange('A4:C13').setBorder(true, true, true, true, true, true);
    helpSheet.getRange('A9').setFontSize(14).setFontWeight('bold');
    helpSheet.getRange('A4:A7').setHorizontalAlignment('center');
    helpSheet.getRange('A10:A13').setHorizontalAlignment('center');
    helpSheet.getRange('B4:B13').setFontWeight('bold');
    
    // 列幅の調整
    helpSheet.setColumnWidth(1, 40);
    helpSheet.setColumnWidth(2, 120);
    helpSheet.setColumnWidth(3, 500);
    
    // ヘッダー行の色設定
    helpSheet.getRange('A1:C1').setBackground('#4285F4').setFontColor('white');
    helpSheet.getRange('A3:C3').setBackground('#E0E0E0');
    helpSheet.getRange('A9:C9').setBackground('#E0E0E0');
    
    // シートをアクティブに
    ss.setActiveSheet(helpSheet);
    
  } catch (error) {
    Logger.log('ヘルプシート作成エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('ヘルプシートの作成中にエラーが発生しました: ' + error.toString());
  }
}

/**
 * APIキーを設定する関数
 */
function setApiKey() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 現在のAPIキーを取得
    var currentApiKey = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY');
    
    // プロンプト表示
    var response = ui.prompt(
      'YouTube API設定',
      '提供されたAPIキーを入力してください:' + 
      (currentApiKey ? '\n\n現在設定されているAPIキー: ' + currentApiKey : ''),
      ui.ButtonSet.OK_CANCEL
    );
    
    // OKボタンがクリックされた場合
    if (response.getSelectedButton() == ui.Button.OK) {
      var apiKey = response.getResponseText().trim();
      if (apiKey) {
        PropertiesService.getScriptProperties().setProperty('YOUTUBE_API_KEY', apiKey);
        
        // APIテストを実行
        var testResponse = testYouTubeAPI(true);
        
        if (testResponse.success) {
          ui.alert('API設定完了', 'YouTube Data API キーが保存され、接続テストに成功しました。\n\nB列にハンドル名（@から始まる）を入力し、「YouTube ツール > ② チャンネル情報取得」を実行してください。', ui.ButtonSet.OK);
        } else {
          ui.alert('APIテスト失敗', 'APIキーは保存されましたが、接続テストに失敗しました。\n\nエラー: ' + testResponse.error, ui.ButtonSet.OK);
        }
        
        // 基本的なシートの初期化（ヘッダー設定など）
        setupBasicSheet();
      } else {
        ui.alert('エラー', 'APIキーが入力されていません。', ui.ButtonSet.OK);
      }
    }
  } catch (error) {
    ui.alert('エラー', 'API設定中にエラーが発生しました: ' + error.toString(), ui.ButtonSet.OK);
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
      ['A1', 'ジャンル'],
      ['B1', 'ハンドル名'],
      ['C1', 'チャンネル名'],
      ['D1', 'チャンネルID'],
      ['E1', '登録者数'],
      ['F1', '総視聴回数'],
      ['G1', '動画本数'],
      ['H1', '作成日'],
      ['I1', 'チャンネルURL'],
      ['J1', 'サムネイル']
    ];
    
    // ヘッダーを設定
    for (var i = 0; i < headers.length; i++) {
      if (!sheet.getRange(headers[i][0]).getValue()) {
        sheet.getRange(headers[i][0]).setValue(headers[i][1]);
      }
    }
    
    // ヘッダー行の書式設定
    sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#f3f3f3');
    
    // 列幅の調整
    sheet.setColumnWidth(1, 120);  // ジャンル
    sheet.setColumnWidth(2, 150);  // ハンドル名
    sheet.setColumnWidth(3, 200);  // チャンネル名
    sheet.setColumnWidth(4, 150);  // チャンネルID
    sheet.setColumnWidth(5, 100);  // 登録者数
    sheet.setColumnWidth(6, 120);  // 総視聴回数
    sheet.setColumnWidth(7, 100);  // 動画本数
    sheet.setColumnWidth(8, 120);  // 作成日
    sheet.setColumnWidth(9, 200);  // チャンネルURL
    sheet.setColumnWidth(10, 120); // サムネイル画像
    
    // ステータス表示用セルを作成
    if (!sheet.getRange('K1').getValue()) {
      sheet.getRange('K1').setValue('ステータス:');
      sheet.getRange('L1').setValue('準備完了');
    }
    
    // プレースホルダーを設定（説明用の注釈も追加）
    sheet.getRange('A2').setValue('例）シニア・自己啓発');
    sheet.getRange('A2').setFontColor('#999999').setFontStyle('italic');
    sheet.getRange('A2').setNote('ジャンルの例です。A列にはジャンル名を入力してください');
    
    sheet.getRange('B2').setValue('@と入力してハンドル名を入力');
    sheet.getRange('B2').setFontColor('#999999').setFontStyle('italic');
    sheet.getRange('B2').setNote('B列にはYouTubeのハンドル名（@から始まる）を入力してください');
    
    // 編集時に書式をリセットするためのトリガーを設定
    setEditTrigger();
    
  } catch (error) {
    Logger.log('シート初期化エラー: ' + error.toString());
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
    if (triggers[i].getHandlerFunction() === 'onEdit' && 
        triggers[i].getEventType() === ScriptApp.EventType.ON_EDIT) {
      hasEditTrigger = true;
      break;
    }
  }
  
  // トリガーがなければ作成
  if (!hasEditTrigger) {
    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

/**
 * 編集時のイベントハンドラ - プレースホルダーから通常テキストへの変換を処理
 */
function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    var value = range.getValue();
    
    // A列またはB列が編集された場合
    if ((range.getColumn() === 1 || range.getColumn() === 2) && range.getRow() > 1) {
      
      // プレースホルダーでない実際の入力値の場合、書式をリセット
      if (value && 
          value !== '例）シニア・自己啓発' && 
          value !== '@と入力してハンドル名を入力') {
        
        // 書式を通常に戻す
        range.setFontColor('black');
        range.setFontStyle('normal');
      }
      // 空白になった場合、プレースホルダーを再表示
      else if (!value) {
        if (range.getColumn() === 1) {
          range.setValue('例）シニア・自己啓発');
          range.setFontColor('#999999').setFontStyle('italic');
        } else if (range.getColumn() === 2) {
          range.setValue('@と入力してハンドル名を入力');
          range.setFontColor('#999999').setFontStyle('italic');
        }
      }
    }
  } catch (error) {
    Logger.log('onEdit エラー: ' + error.toString());
  }
}

/**
 * テスト関数 - YouTube APIが正しく動作しているか確認
 * @param {boolean} silent - trueの場合、ダイアログを表示せずに結果を返す
 * @return {Object} - 結果オブジェクト（silentがtrueの場合）
 */
function testYouTubeAPI(silent) {
  var ui = SpreadsheetApp.getUi();
  var result = { success: false, error: '' };
  
  try {
    // APIキーを取得
    var apiKey = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY');
    
    // APIキーがない場合は設定を促す
    if (!apiKey) {
      var message = '「YouTube ツール > ① API設定・テスト」を実行してAPIキーを設定してください。';
      if (silent) {
        result.error = 'APIキーが設定されていません';
        return result;
      } else {
        ui.alert('APIキーが設定されていません', message, ui.ButtonSet.OK);
        return;
      }
    }
    
    // ステータス表示
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var statusCell = sheet.getRange('L1');
    statusCell.setValue('テスト中...');
    
    // 単純な検索テスト
    var options = {
      'method': 'get',
      'muteHttpExceptions': true
    };
    
    // チャンネル情報のテスト (YouTube公式チャンネル)
    var url = 'https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&forUsername=YouTube&key=' + apiKey;
    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    
    if (data && data.items && data.items.length > 0) {
      statusCell.setValue('テスト成功');
      result.success = true;
      
      if (!silent) {
        ui.alert(
          'API接続成功！', 
          'YouTube APIに正常に接続できました。\n\n' +
          'チャンネル名: ' + data.items[0].snippet.title + '\n' +
          '登録者数: ' + parseInt(data.items[0].statistics.subscriberCount).toLocaleString() + '\n' +
          '動画数: ' + data.items[0].statistics.videoCount,
          ui.ButtonSet.OK
        );
      }
    } else {
      // 検索APIをテスト
      url = 'https://www.googleapis.com/youtube/v3/search?part=snippet&q=YouTube&type=channel&maxResults=1&key=' + apiKey;
      response = UrlFetchApp.fetch(url, options);
      data = JSON.parse(response.getContentText());
      
      if (data && data.items && data.items.length > 0) {
        statusCell.setValue('テスト成功');
        result.success = true;
        
        if (!silent) {
          ui.alert(
            'API検索接続成功！', 
            '検索APIは正常に動作しています。\n\n' +
            'チャンネル名: ' + data.items[0].snippet.title,
            ui.ButtonSet.OK
          );
        }
      } else {
        statusCell.setValue('テスト失敗');
        result.error = 'APIは接続できましたが、期待した応答が得られませんでした。';
        
        if (!silent) {
          ui.alert(
            'API応答エラー', 
            result.error + '\n\nレスポンス: ' + response.getContentText(),
            ui.ButtonSet.OK
          );
        }
      }
    }
    
    return result;
  } catch (error) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var statusCell = sheet.getRange('L1');
    statusCell.setValue('エラー');
    
    result.error = error.toString();
    
    if (!silent) {
      ui.alert(
        'APIエラー', 
        'YouTube APIに接続できませんでした: ' + error.toString(),
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
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ui = SpreadsheetApp.getUi();
    
    // APIキーの確認
    var apiKey = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY');
    if (!apiKey) {
      ui.alert('APIキーが設定されていません', '「YouTube ツール > ① API設定・テスト」を実行してAPIキーを設定してください。', ui.ButtonSet.OK);
      return;
    }
    
    // プログレスバーセルを用意
    var statusCell = sheet.getRange('L1');
    statusCell.setValue('準備中...');
    
    // B列のデータを取得（ハンドル名の列）
    var data = sheet.getRange('B:B').getValues();
    var colors = sheet.getRange('B:B').getFontColors(); // フォント色を取得
    var handles = [];
    
    // ハンドル名をリストに追加（空でないセルかつ@で始まるもの）
    for (var i = 1; i < data.length; i++) {  // i=1 からスタート（ヘッダー行をスキップ）
      var handle = data[i][0].toString().trim();
      var color = colors[i][0];
      
      // 実際のハンドル名のみを処理（プレースホルダーでないもの＝黒いテキスト）
      if (handle && 
          handle.startsWith('@') && 
          color !== '#999999' && // グレーのプレースホルダーは除外
          handle !== '@と入力してハンドル名を入力') {
        handles.push({
          handle: handle,
          row: i + 1  // スプレッドシートの行番号（1-indexed）
        });
      }
    }
    
    if (handles.length === 0) {
      statusCell.setValue('エラー');
      ui.alert(
        'エラー',
        'B列に@で始まるハンドル名が見つかりませんでした。例: @YouTube',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 進捗状況を表示
    statusCell.setValue('0%');
    
    // 各ハンドルの情報を取得
    var successCount = 0;
    var errorCount = 0;
    
    for (var i = 0; i < handles.length; i++) {
      var channelInfo = getChannelByHandle(handles[i].handle, apiKey);
      var row = handles[i].row;
      
      if (channelInfo) {
        // 成功した場合、情報を表示
        var statistics = channelInfo.statistics || {};
        
        // 登録者数（非公開の場合は「非公開」と表示）
        var subscriberCount = statistics.hiddenSubscriberCount ? 
          '非公開' : 
          (statistics.subscriberCount ? parseInt(statistics.subscriberCount).toLocaleString() : '0');
          
        // スプレッドシートに情報を書き込み
        sheet.getRange(row, 3).setValue(channelInfo.snippet.title);  // チャンネル名
        sheet.getRange(row, 4).setValue(channelInfo.id);  // チャンネルID
        sheet.getRange(row, 5).setValue(subscriberCount);  // 登録者数
        sheet.getRange(row, 6).setValue(parseInt(statistics.viewCount || 0).toLocaleString());  // 総視聴回数
        sheet.getRange(row, 7).setValue(parseInt(statistics.videoCount || 0).toLocaleString());  // 動画本数
        sheet.getRange(row, 8).setValue(new Date(channelInfo.snippet.publishedAt).toLocaleDateString());  // 作成日
        sheet.getRange(row, 9).setValue('https://www.youtube.com/channel/' + channelInfo.id);  // チャンネルURL
        
        // サムネイル画像を設定（IMAGE関数を使用）
        var thumbnailUrl = channelInfo.snippet.thumbnails.default.url;
        var imageFormula = '=IMAGE("' + thumbnailUrl + '", 1)';  // 1=画像をセルに合わせてサイズ調整
        sheet.getRange(row, 10).setValue(imageFormula);
        
        successCount++;
      } else {
        // 失敗した場合、エラーメッセージを表示
        sheet.getRange(row, 3).setValue('チャンネルが見つかりません');
        sheet.getRange(row, 4, 1, 7).setValue('');  // 他のセルをクリア
        errorCount++;
      }
      
      // 処理状況を更新
      var progress = Math.round(((i + 1) / handles.length) * 100);
      statusCell.setValue(progress + '%');
      SpreadsheetApp.flush();  // 画面を更新
    }
    
    // 処理完了
    statusCell.setValue('完了');
    
    // 行の高さを調整（サムネイル画像用）
    for (var i = 0; i < handles.length; i++) {
      sheet.setRowHeight(handles[i].row, 60);
    }
    
    // 完了メッセージ
    ui.alert(
      '処理完了',
      '合計: ' + handles.length + ' 件\n' +
      '成功: ' + successCount + ' 件\n' +
      '失敗: ' + errorCount + ' 件\n\n' +
      '「YouTube ツール > ③ ベンチマークレポート作成」を実行するとレポートを生成できます',
      ui.ButtonSet.OK
    );
  } catch (error) {
    Logger.log('エラー: ' + error.toString());
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var statusCell = sheet.getRange('L1');
    statusCell.setValue('エラー');
    
    SpreadsheetApp.getUi().alert(
      'エラー',
      'チャンネル情報の取得中にエラーが発生しました: ' + error.toString(),
      ui.ButtonSet.OK
    );
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
    var username = handle.replace('@', '');
    var options = {
      'method': 'get',
      'muteHttpExceptions': true
    };
    
    // 1. 検索APIを使用してハンドル名で検索（そのままの形式で）
    var searchUrl = 'https://www.googleapis.com/youtube/v3/search?part=snippet&q=' + 
                    encodeURIComponent(handle) + '&type=channel&maxResults=10&key=' + apiKey;
    
    var searchResponse = UrlFetchApp.fetch(searchUrl, options);
    var searchData = JSON.parse(searchResponse.getContentText());
    
    if (searchData && searchData.items && searchData.items.length > 0) {
      for (var i = 0; i < searchData.items.length; i++) {
        var channelId = searchData.items[i].id.channelId;
        
        // チャンネルの詳細情報を取得
        var channelUrl = 'https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=' + 
                         channelId + '&key=' + apiKey;
        
        var channelResponse = UrlFetchApp.fetch(channelUrl, options);
        var channelData = JSON.parse(channelResponse.getContentText());
        
        if (channelData && channelData.items && channelData.items.length > 0) {
          var channelDetails = channelData.items[0];
          
          // チャンネル名かcustomUrlがユーザー名を含むものを優先
          if (channelDetails.snippet.title.toLowerCase().includes(username.toLowerCase()) || 
              (channelDetails.snippet.customUrl && 
               channelDetails.snippet.customUrl.toLowerCase().includes(username.toLowerCase()))) {
            return channelDetails;
          }
        }
      }
      
      // 名前に一致するものがなければ最初の結果を返す
      var firstChannelId = searchData.items[0].id.channelId;
      var firstChannelUrl = 'https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=' + 
                           firstChannelId + '&key=' + apiKey;
      
      var firstChannelResponse = UrlFetchApp.fetch(firstChannelUrl, options);
      var firstChannelData = JSON.parse(firstChannelResponse.getContentText());
      
      if (firstChannelData && firstChannelData.items && firstChannelData.items.length > 0) {
        return firstChannelData.items[0];
      }
    }
    
    // 2. @なしのユーザー名で検索
    var usernameSearchUrl = 'https://www.googleapis.com/youtube/v3/search?part=snippet&q=' + 
                           encodeURIComponent(username) + '&type=channel&maxResults=5&key=' + apiKey;
    
    var usernameSearchResponse = UrlFetchApp.fetch(usernameSearchUrl, options);
    var usernameSearchData = JSON.parse(usernameSearchResponse.getContentText());
    
    if (usernameSearchData && usernameSearchData.items && usernameSearchData.items.length > 0) {
      var channelId = usernameSearchData.items[0].id.channelId;
      var channelUrl = 'https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&id=' + 
                       channelId + '&key=' + apiKey;
      
      var channelResponse = UrlFetchApp.fetch(channelUrl, options);
      var channelData = JSON.parse(channelResponse.getContentText());
      
      if (channelData && channelData.items && channelData.items.length > 0) {
        return channelData.items[0];
      }
    }
    
    // 3. forUsername パラメータで検索 (旧式、バックアップとして)
    var usernameUrl = 'https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics&forUsername=' + 
                      encodeURIComponent(username) + '&key=' + apiKey;
    
    var usernameResponse = UrlFetchApp.fetch(usernameUrl, options);
    var usernameData = JSON.parse(usernameResponse.getContentText());
    
    if (usernameData && usernameData.items && usernameData.items.length > 0) {
      return usernameData.items[0];
    }
    
    return null;
  } catch (error) {
    Logger.log('チャンネル取得エラー (' + handle + '): ' + error.toString());
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
    var statusCell = sheet.getRange('L1');
    statusCell.setValue('レポート作成中...');
    
    // 既存データを取得
    var data = sheet.getDataRange().getValues();
    
    // ヘッダー行を除く有効なデータ行を確認
    var validRows = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2] !== 'チャンネルが見つかりません') {
        validRows++;
      }
    }
    
    if (validRows === 0) {
      statusCell.setValue('エラー');
      ui.alert('有効なデータがありません。先にチャンネル情報を取得してください。');
      return;
    }
    
    statusCell.setValue('25%');
    
    // 現在日時からユニークなレポート名を生成
    var now = new Date();
    var reportName = 'ベンチマークレポート_' + 
                     now.getFullYear() + 
                     ('0' + (now.getMonth() + 1)).slice(-2) + 
                     ('0' + now.getDate()).slice(-2) + '_' +
                     ('0' + now.getHours()).slice(-2) + 
                     ('0' + now.getMinutes()).slice(-2);
    
    // 新しいシートを作成
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.insertSheet(reportName);
    
    // レポートヘッダーを設定
    reportSheet.appendRow(['YouTubeチャンネルベンチマークレポート']);
    reportSheet.appendRow(['作成日: ' + now.toLocaleDateString() + ' ' + now.toLocaleTimeString()]);
    reportSheet.appendRow(['チャンネル数: ' + validRows + ' 件']);
    reportSheet.appendRow(['']);
    
    // 統計情報を計算
    var totalChannels = validRows;
    var totalSubscribers = 0;
    var totalViews = 0;
    var totalVideos = 0;
    
    for (var i = 1; i < data.length; i++) {
      // チャンネル名がある行のみ処理
      if (data[i][2] && data[i][2] !== 'チャンネルが見つかりません') {
        var subscribers = data[i][4];
        if (subscribers !== '非公開' && subscribers) {
          // カンマを削除して数値に変換
          subscribers = subscribers.toString().replace(/,/g, '');
          if (!isNaN(parseInt(subscribers))) {
            totalSubscribers += parseInt(subscribers);
          }
        }
        
        var views = data[i][5] ? data[i][5].toString().replace(/,/g, '') : '0';
        if (!isNaN(parseInt(views))) {
          totalViews += parseInt(views);
        }
        
        var videos = data[i][6] ? data[i][6].toString().replace(/,/g, '') : '0';
        if (!isNaN(parseInt(videos))) {
          totalVideos += parseInt(videos);
        }
      }
    }
    
    // 進捗状況更新
    statusCell.setValue('50%');
    
    // 統計サマリーを追加
    reportSheet.appendRow(['統計サマリー:']);
    reportSheet.appendRow(['指標', '合計', '平均', '最大値', '中央値']);
    
    // 登録者数の統計データを計算
    var subscribersStats = calculateStats(data, 4);
    var viewsStats = calculateStats(data, 5);
    var videosStats = calculateStats(data, 6);
    
    // 統計データを追加
    reportSheet.appendRow(['登録者数', 
                          totalSubscribers.toLocaleString(), 
                          Math.round(subscribersStats.average).toLocaleString(), 
                          subscribersStats.max.toLocaleString(), 
                          subscribersStats.median.toLocaleString()]);
    
    reportSheet.appendRow(['総視聴回数', 
                          totalViews.toLocaleString(), 
                          Math.round(viewsStats.average).toLocaleString(), 
                          viewsStats.max.toLocaleString(), 
                          viewsStats.median.toLocaleString()]);
    
    reportSheet.appendRow(['動画本数', 
                          totalVideos.toLocaleString(), 
                          Math.round(videosStats.average).toLocaleString(), 
                          videosStats.max.toLocaleString(), 
                          videosStats.median.toLocaleString()]);
    
    reportSheet.appendRow(['']);
    
    // 進捗状況更新
    statusCell.setValue('70%');
    
    // トップ10チャンネルを追加
    reportSheet.appendRow(['登録者数トップ10チャンネル:']);
    reportSheet.appendRow([
      'ランク',
      'ジャンル',
      'ハンドル名',
      'チャンネル名', 
      '登録者数', 
      '総視聴回数', 
      '動画本数',
      'チャンネルURL',
      'サムネイル'
    ]);
    
    // データを登録者数でソート（非公開は除外）
    var sortableData = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] && data[i][2] !== 'チャンネルが見つかりません' && data[i][4] !== '非公開') {
        var subscribers = data[i][4] ? data[i][4].toString().replace(/,/g, '') : '0';
        if (!isNaN(parseInt(subscribers))) {
          // IMAGE関数からサムネイルURLを抽出
          var thumbnailUrl = '';
          var imageFormula = data[i][9] ? data[i][9].toString() : '';
          if (imageFormula.indexOf('=IMAGE("') === 0) {
            thumbnailUrl = imageFormula.substring(8, imageFormula.indexOf('", 1'));
          }
          
          sortableData.push({
            genre: data[i][0],  // ジャンル
            handle: data[i][1],  // ハンドル名
            name: data[i][2],    // チャンネル名
            subscribers: parseInt(subscribers),
            views: data[i][5] ? data[i][5].toString().replace(/,/g, '') : '0',
            videos: data[i][6] ? data[i][6].toString().replace(/,/g, '') : '0',
            url: data[i][8],
            thumbnail: thumbnailUrl
          });
        }
      }
    }
    
    // 登録者数でソート
    sortableData.sort(function(a, b) {
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
        sortableData[i].thumbnail ? '=IMAGE("' + sortableData[i].thumbnail + '", 1)' : ''
      ]);
    }
    
    // サムネイル用に行の高さを調整
    for (var i = 0; i < maxRows; i++) {
      reportSheet.setRowHeight(13 + i, 60);
    }
    
    // 進捗状況更新
    statusCell.setValue('90%');
    
    // 平均以上の登録者数を持つチャンネル情報
    reportSheet.appendRow(['']);
    reportSheet.appendRow(['平均以上の登録者数を持つチャンネル:']);
    reportSheet.appendRow([
      'ジャンル',
      'ハンドル名',
      'チャンネル名', 
      '登録者数', 
      '総視聴回数', 
      '動画本数',
      'チャンネルURL',
      'サムネイル'
    ]);
    
    var avgSubscribers = subscribersStats.average;
    var aboveAvgChannels = sortableData.filter(function(channel) {
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
        aboveAvgChannels[i].thumbnail ? '=IMAGE("' + aboveAvgChannels[i].thumbnail + '", 1)' : ''
      ]);
      
      // サムネイル用に行の高さを調整
      reportSheet.setRowHeight(16 + maxRows + i, 60);
    }
    
    // ジャンル別の分析を追加
    var genreAnalysis = analyzeByGenre(sortableData);
    if (genreAnalysis.genres.length > 1) {  // 複数のジャンルがある場合のみ
      reportSheet.appendRow(['']);
      reportSheet.appendRow(['ジャンル別分析:']);
      reportSheet.appendRow(['ジャンル', 'チャンネル数', '平均登録者数', '平均視聴回数', '平均動画数']);
      
      for (var i = 0; i < genreAnalysis.genres.length; i++) {
        var genre = genreAnalysis.genres[i];
        var stats = genreAnalysis.stats[genre];
        
        reportSheet.appendRow([
          genre,
          stats.count,
          Math.round(stats.avgSubscribers).toLocaleString(),
          Math.round(stats.avgViews).toLocaleString(),
          Math.round(stats.avgVideos).toLocaleString()
        ]);
      }
    }
    
    // レポートの書式設定
    formatBenchmarkReport(reportSheet, maxRows, aboveAvgChannels.length, genreAnalysis.genres.length > 1);
    
    // レポートシートをアクティブにする
    ss.setActiveSheet(reportSheet);
    
    // 進捗状況更新
    statusCell.setValue('完了');
    
    ui.alert('ベンチマークレポートが作成されました。');
  } catch (error) {
    Logger.log('レポート作成エラー: ' + error.toString());
    SpreadsheetApp.getUi().alert('レポート作成中にエラーが発生しました: ' + error.toString());
    
    var progressCell = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('L1');
    if (progressCell) {
      progressCell.setValue('エラー');
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
    var genre = sortableData[i].genre || '未分類';
    
    if (!genreStats[genre]) {
      genreStats[genre] = {
        count: 0,
        totalSubscribers: 0,
        totalViews: 0,
        totalVideos: 0
      };
      genres.push(genre);
    }
    
    genreStats[genre].count++;
    genreStats[genre].totalSubscribers += sortableData[i].subscribers;
    genreStats[genre].totalViews += parseInt(sortableData[i].views.replace(/,/g, ''));
    genreStats[genre].totalVideos += parseInt(sortableData[i].videos.replace(/,/g, ''));
  }
  
  // 平均値を計算
  for (var genre in genreStats) {
    genreStats[genre].avgSubscribers = genreStats[genre].totalSubscribers / genreStats[genre].count;
    genreStats[genre].avgViews = genreStats[genre].totalViews / genreStats[genre].count;
    genreStats[genre].avgVideos = genreStats[genre].totalVideos / genreStats[genre].count;
  }
  
  return {
    genres: genres,
    stats: genreStats
  };
}

/**
 * レポートの書式設定を行う関数
 */
function formatBenchmarkReport(sheet, topChannelsCount, aboveAvgCount, hasGenreAnalysis) {
  // タイトルと日付の書式設定
  sheet.getRange('A1:I1').merge();
  sheet.getRange('A1').setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A1').setBackground('#4285F4').setFontColor('white');
  
  sheet.getRange('A2:I2').merge();
  sheet.getRange('A2').setFontStyle('italic').setHorizontalAlignment('center');
  
  sheet.getRange('A3:I3').merge();
  sheet.getRange('A3').setFontWeight('bold').setHorizontalAlignment('center');
  
  // 統計サマリーの書式設定
  sheet.getRange('A5').setFontWeight('bold');
  sheet.getRange('A6:E6').setFontWeight('bold').setBackground('#E0E0E0');
  sheet.getRange('A7:E9').setBorder(true, true, true, true, true, true);
  sheet.getRange('A6:E9').setHorizontalAlignment('center');
  
  // トップ10チャンネルの書式設定
  sheet.getRange('A11').setFontWeight('bold');
  sheet.getRange('A12:I12').setFontWeight('bold').setBackground('#E0E0E0');
  sheet.getRange('A13:I' + (12 + topChannelsCount)).setBorder(true, true, true, true, true, true);
  
  // 平均以上の書式設定
  var aboveAvgStartRow = 14 + topChannelsCount;
  sheet.getRange('A' + aboveAvgStartRow).setFontWeight('bold');
  sheet.getRange('A' + (aboveAvgStartRow + 1) + ':H' + (aboveAvgStartRow + 1)).setFontWeight('bold').setBackground('#E0E0E0');
  sheet.getRange('A' + (aboveAvgStartRow + 2) + ':H' + (aboveAvgStartRow + 1 + aboveAvgCount)).setBorder(true, true, true, true, true, true);
  
  // ジャンル分析の書式設定（存在する場合）
  if (hasGenreAnalysis) {
    var genreStartRow = aboveAvgStartRow + aboveAvgCount + 3;
    sheet.getRange('A' + genreStartRow).setFontWeight('bold');
    sheet.getRange('A' + (genreStartRow + 1) + ':E' + (genreStartRow + 1)).setFontWeight('bold').setBackground('#E0E0E0');
    
    // ジャンル行数を取得（ヘッダー+データ行）
    var genreCount = 0;
    for (var i = genreStartRow + 2; i <= sheet.getLastRow(); i++) {
      if (sheet.getRange('A' + i).getValue()) {
        genreCount++;
      }
    }
    
    if (genreCount > 0) {
      sheet.getRange('A' + (genreStartRow + 2) + ':E' + (genreStartRow + 1 + genreCount)).setBorder(true, true, true, true, true, true);
    }
  }
  
  // 列幅の調整
  sheet.setColumnWidth(1, 80);  // ランク/ジャンル
  sheet.setColumnWidth(2, 120); // ジャンル/ハンドル名
  sheet.setColumnWidth(3, 150); // ハンドル名/チャンネル名
  sheet.setColumnWidth(4, 200); // チャンネル名/登録者数
  sheet.setColumnWidth(5, 100); // 登録者数/視聴回数
  sheet.setColumnWidth(6, 120); // 視聴回数/動画数
  sheet.setColumnWidth(7, 100); // 動画数/URL
  sheet.setColumnWidth(8, 250); // URL
  sheet.setColumnWidth(9, 150); // サムネイル
  
  // チャンネルURLを青色のハイパーリンクに
  var urls = sheet.getRange('H13:H' + (12 + topChannelsCount)).getValues();
  for (var i = 0; i < urls.length; i++) {
    if (urls[i][0]) {
      var range = sheet.getRange('H' + (13 + i));
      range.setFontColor('#1155CC').setFontLine('underline');
    }
  }
  
  var aboveAvgUrls = sheet.getRange('G' + (aboveAvgStartRow + 2) + ':G' + (aboveAvgStartRow + 1 + aboveAvgCount)).getValues();
  for (var i = 0; i < aboveAvgUrls.length; i++) {
    if (aboveAvgUrls[i][0]) {
      var range = sheet.getRange('G' + (aboveAvgStartRow + 2 + i));
      range.setFontColor('#1155CC').setFontLine('underline');
    }
  }
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
    if (data[i][2] && data[i][2] !== 'チャンネルが見つかりません' && data[i][colIndex] !== '非公開') {
      var value = data[i][colIndex] ? data[i][colIndex].toString().replace(/,/g, '') : '0';
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
      median: 0
    };
  }
  
  // 昇順ソート
  values.sort(function(a, b) {
    return a - b;
  });
  
  // 平均値計算
  var sum = values.reduce(function(a, b) { return a + b; }, 0);
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
    median: median
  };
}