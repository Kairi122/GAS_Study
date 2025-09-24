/**
 * 毎日の朝会アジェンダを生成する関数
 * - テンプレート（A4:D7）をコピーして8行目に挿入します。
 * - コピーしたアジェンダの日付を「翌営業日」に更新します。
 * - 土日・祝日は実行しません。
 */
function createDailyAgenda() {
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0:日曜日,6:土曜日
  const calendarId = 'ja.japanese#holiday@group.v.calendar.google.com';
  const holidayCalendar = CalendarApp.getCalendarById(calendarId);

  // --- 実行条件の判定 ---

  // 1. 土日かどうかを判定
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    console.log('土日のため処理をスキップしました。');
    return; // 処理を終了
  }

  // 2. 日本の祝日かどうかを判定
  const events = holidayCalendar.getEventsForDay(today);
  if (events.length > 0) {
    console.log('祝日のため処理をスキップしました。');
    return; // 処理を終了
  }

  // --- スプレッドシートの操作 ---

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("朝会アジェンダ");

  if (!sheet) {
    console.log('指定されたシートが見つかりませんでした。');
    return;
  }

  const templateRangeAddress = 'A4:D7';
  const numRowsInTemplate = 4;
  const insertRowPosition = 8;

  sheet.insertRowsBefore(insertRowPosition, numRowsInTemplate);

  const templateRange = sheet.getRange(templateRangeAddress);
  const destinationRange = sheet.getRange(insertRowPosition, 1, numRowsInTemplate, 4);
  templateRange.copyTo(destinationRange);

  let nextBusinessDay = new Date();
  nextBusinessDay.setDate(nextBusinessDay.getDate() + 1);

  while (true) {
    const day = nextBusinessDay.getDay();
    const isWeekend = (day === 0 || day === 6); // 日曜日(0)土曜日(6)
    const isHoliday = holidayCalendar.getEventsForDay(nextBusinessDay).length > 0;

    if (isWeekend || isHoliday) {
      nextBusinessDay.setDate(nextBusinessDay.getDate() + 1);
    } else {
      break;
    }
  }

  const dateCell = sheet.getRange('A8');
  const formattedDate = Utilities.formatDate(nextBusinessDay, ss.getSpreadsheetTimeZone(), 'MM/dd');
  dateCell.setValue(formattedDate);

  dateCell.setHorizontalAlignment('center').setVerticalAlignment('middle');

  console.log('翌営業日 (' + formattedDate + ') のアジェンダを作成しました。');
}