function addComment() {
  const comment = "Hello!World!";

  // 1. 現在開いているシートを取得
  const sheet = SpreadsheetApp.getActiveSheet();

  // 2. 取得したシートの最終行にコメントを追加
  sheet.appendRow([comment]);
}