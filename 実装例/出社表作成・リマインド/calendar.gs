// -------- 設定項目 --------

// ▼ 1. プルダウンリストの項目
const STATUS_OPTIONS = ["出社", "休み", "リモート", "午後休", "午前休","有給","研修"];

const SS_NAME = "出社表"; // スプレッドシートのシート名と合わせてください

const WEEKEND_AND_HOLIDAY_COLOR = "#fce5cd"; // 土日・祝日の背景色

// -------- 設定はここまで --------

/**
 * メインの関数
 */
function updateAttendanceSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SS_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SS_NAME);
  }
  
  // --- 現在の従業員リストをシートから取得 ---
  const employees = sheet.getRange("A2:A" + sheet.getLastRow()).getValues()
                         .flat() // 2次元配列を1次元に変換
                         .filter(name => name !== ""); // 空の名前を除外

  if (employees.length === 0) {
    sheet.getRange("A2").setValue("ここに名前を入力");
    SpreadsheetApp.getUi().alert("従業員名が入力されていません。A2セル以下に名前を入力してから再度実行してください。");
    return;
  }
  // --- シートのクリア（名前列以外） ---
  if (sheet.getLastColumn() > 1) {
    sheet.getRange(1, 2, sheet.getMaxRows(), sheet.getLastColumn() - 1).clear();
  }

  // --- 日付データの準備 ---
  const today = new Date();
  const dates = [];
  // 今月
  const thisMonthLastDay = new Date(today.getFullYear(), today.getMonth() + 1, 0).getDate();
  for (let i = 1; i <= thisMonthLastDay; i++) {
    dates.push(new Date(today.getFullYear(), today.getMonth(), i));
  }
  // 来月
  const nextMonth = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const nextMonthLastDay = new Date(nextMonth.getFullYear(), nextMonth.getMonth() + 1, 0).getDate();
  for (let i = 1; i <= nextMonthLastDay; i++) {
    dates.push(new Date(nextMonth.getFullYear(), nextMonth.getMonth(), i));
  }

  // 日本の祝日カレンダーを取得
  const holidayCalendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  const holidays = holidayCalendar.getEvents(dates[0], dates[dates.length - 1]).map(e => e.getStartTime().toLocaleString());

  const header = [];
  const backgroundColors = [];
  const weekdays = ["(日)", "(月)", "(火)", "(水)", "(木)", "(金)", "(土)"];

  dates.forEach(date => {
    header.push(`${date.getMonth() + 1}/${date.getDate()}\n${weekdays[date.getDay()]}`);
    const dateString = date.toLocaleString();
    if (date.getDay() === 0 || date.getDay() === 6 || holidays.includes(dateString)) {
      backgroundColors.push(WEEKEND_AND_HOLIDAY_COLOR);
    } else {
      backgroundColors.push(null);
    }
  });

  // --- シートへの書き込み ---
  sheet.getRange(1, 2, 1, header.length).setValues([header])
    .setHorizontalAlignment("center").setFontWeight("bold").setWrap(true);

  // --- プルダウンリスト（データの入力規則）の設定 ---
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS).setAllowInvalid(false).build();
  sheet.getRange(2, 2, employees.length, dates.length).setDataValidation(rule);

  // --- 書式設定 ---
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setRowHeight(1, 40);
  sheet.autoResizeColumn(1);
  sheet.setColumnWidths(2, dates.length, 50);

  const entireSheetRange = sheet.getRange(1, 1, sheet.getMaxRows(), header.length + 1);
  const sheetBackgrounds = entireSheetRange.getBackgrounds();
  for(let col = 0; col < backgroundColors.length; col++){
    if(backgroundColors[col]){
       for(let row = 0; row < sheetBackgrounds.length; row++){
         sheetBackgrounds[row][col+1] = backgroundColors[col];
       }
    }
  }
  entireSheetRange.setBackgrounds(sheetBackgrounds);

  sheet.getRange(2, 1, employees.length, 1).setFontWeight("bold");
  sheet.getRange(1, 1).setValue("名前").setFontWeight("bold").setHorizontalAlignment("center");
}