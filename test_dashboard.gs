/**
 * ダッシュボード入力検証テスト関数
 */
function testInputValidation() {
  var testInputs = [
    "@YouTube",
    "https://www.youtube.com/@YouTube",
    "https://www.youtube.com/c/YouTube", 
    "https://www.youtube.com/channel/UC-9-kyTW8ZkZNDHQJ6FgpwQ",
    "UC-9-kyTW8ZkZNDHQJ6FgpwQ",
    "YouTube",
    "invalid input",
    "@invalid@handle",
    ""
  ];
  
  for (var i = 0; i < testInputs.length; i++) {
    var input = testInputs[i];
    var normalized = normalizeChannelInput(input);
    Logger.log("入力: '" + input + "' → 正規化: '" + normalized + "'");
  }
  
  SpreadsheetApp.getUi().alert(
    "テスト完了",
    "入力検証テストが完了しました。ログを確認してください。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * ダッシュボードを強制的に再作成するテスト関数
 */
function recreateDashboard() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 既存のダッシュボードを削除
    var existingDashboard = ss.getSheetByName("📊 YouTube チャンネル分析");
    if (existingDashboard) {
      ss.deleteSheet(existingDashboard);
    }
    
    // 新しいダッシュボードを作成
    createUnifiedDashboard();
    
    SpreadsheetApp.getUi().alert(
      "再作成完了",
      "ダッシュボードを再作成しました。B8セルに入力してB9セルで「分析」を試してください。",
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