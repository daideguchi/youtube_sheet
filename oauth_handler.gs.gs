/**
 * OAuth認証用のWebアプリハンドラー
 */
function doGet(e) {
  const code = e.parameter.code;
  const state = e.parameter.state;
  const error = e.parameter.error;

  if (error) {
    return HtmlService.createHtmlOutput(`
      <h2>認証エラー</h2>
      <p>エラー: ${error}</p>
      <p>ウィンドウを閉じてスプレッドシートに戻ってください。</p>
    `);
  }

  if (code) {
    // 認証コードを一時保存
    PropertiesService.getUserProperties().setProperty("TEMP_AUTH_CODE", code);

    return HtmlService.createHtmlOutput(`
      <h2>認証成功!</h2>
      <p>認証コードを取得しました。</p>
      <p>このウィンドウを閉じて、スプレッドシートで「認証完了」ボタンをクリックしてください。</p>
      <script>
        // 5秒後に自動でウィンドウを閉じる
        setTimeout(function() {
          window.close();
        }, 5000);
      </script>
    `);
  }

  return HtmlService.createHtmlOutput(`
    <h2>OAuth認証</h2>
    <p>認証に必要なパラメータが見つかりません。</p>
  `);
}
