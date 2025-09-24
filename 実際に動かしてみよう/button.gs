// ボタンをクリックしたときに実行される関数
function button() {

  // 1. 現在アクティブなシートを取得
  const sheet = SpreadsheetApp.getActiveSheet();

  // 2. 現在の日時を取得
  const now = new Date();

  // 4. 日時を「年/月/日 時:分:秒」の形式の文字列に変換する
  //    "JST"は日本のタイムゾーンを指定
  const formattedDate = Utilities.formatDate(now, "JST", "yyyy/MM/dd HH:mm:ss");

  // 5. 記録する内容
  const comment = "現在の日付と時間";

  // 6. B2:C2の範囲を指定して、それぞれのセルに値を書き込む
  sheet.getRange("B2:C2").setValues([[comment, formattedDate]]);
}