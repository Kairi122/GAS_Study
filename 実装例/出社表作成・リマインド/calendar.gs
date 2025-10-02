// -------- 設定項目 --------

// ▼ 1. プルダウンリストの項目
const STATUS_OPTIONS = ["出社", "休み", "リモート", "午後休", "午前休", "有給", "研修"];

// ▼ 2. スプレッドシートのシート名
const SS_NAME = "出社表"; // スプレッドシートのシート名と合わせてください

// ▼ 3. 色の設定 (カラーコードで指定)
const WEEKEND_AND_HOLIDAY_COLOR = "#fce5cd"; // 土日・祝日の背景色

// -------- 設定はここまで --------


/**
 * メイン関数：この関数を毎月1日に実行するようにトリガー設定します。
 */
function monthlyUpdateAttendanceSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SS_NAME);

  if (!sheet) {
    // シートが存在しない場合は初期作成
    console.log(`シート「${SS_NAME}」が存在しないため、初期作成を実行します。`);
    createInitialSheet(spreadsheet);
  } else {
    // 既存シートを「今月」と「来月」の2ヶ月分に再構築します
    console.log(`シート「${SS_NAME}」を再構築します。`);
    rebuildCalendarSheet(sheet);
  }
}

/**
 * 初回実行時にシートを新規作成する関数
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - 対象のスプレッドシート
 */
function createInitialSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet(SS_NAME);

  // 基本的な書式設定
  sheet.getRange(1, 1).setValue("名前").setFontWeight("bold").setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
  sheet.setRowHeight(1, 40);

  // 初回は従業員がいないため、サンプルテキストを配置
  sheet.getRange("A2").setValue("ここに名前を入力");
  SpreadsheetApp.getUi().alert(`シート「${SS_NAME}」を新規作成しました。A2セル以下に従業員名を入力してください。`);

  // カレンダー部分を構築
  rebuildCalendarSheet(sheet);
}


/**
 * シートを「今月」と「来月」の2ヶ月分で再構築する関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象のシート
 */
function rebuildCalendarSheet(sheet) {
  // --- 1. 既存の予定データを一時保存 ---
  const preservedData = {};
  if (sheet.getLastColumn() > 1 && sheet.getLastRow() > 1) {
    const dataRange = sheet.getDataRange();
    const allValues = dataRange.getValues();
    const header = allValues[0]; // [名前, 10/1, 10/2, ...]
    const employeeRows = allValues.slice(1);

    employeeRows.forEach(row => {
      const employeeName = row[0];
      if (employeeName) {
        preservedData[employeeName] = {};
        for (let i = 1; i < header.length; i++) {
          const dateStr = header[i].split('\n')[0]; // "10/1"
          if (row[i]) { // ステータスが入力されていれば保存
            preservedData[employeeName][dateStr] = row[i];
          }
        }
      }
    });
    console.log("既存の予定データを一時保存しました。");
  }

  // --- 2. B列以降（カレンダー部分）を一度すべてクリア ---
  if (sheet.getLastColumn() > 1) {
    sheet.getRange(1, 2, sheet.getMaxRows(), sheet.getLastColumn() - 1).clear();
  }

  // --- 3. カレンダーの再構築 (今月 + 来月) ---
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

  // ヘッダーと背景色データを生成
  const { header, backgrounds } = generateCalendarData(dates);

  // 新しいヘッダーを書き込み
  sheet.getRange(1, 2, 1, header.length)
    .setValues([header])
    .setHorizontalAlignment("center").setFontWeight("bold").setWrap(true);

  console.log(`カレンダーを${today.getMonth()+1}月と${nextMonth.getMonth()+1}月分で再構築しました。`);

  // --- 4. 予定データを復元 ---
  const employeeData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  if (employeeData.length > 0) {
    const restoredValues = [];
    employeeData.forEach(name => {
      const newRow = [];
      if (name && preservedData[name]) {
        header.forEach(h => {
          const dateStr = h.split('\n')[0];
          newRow.push(preservedData[name][dateStr] || ""); // 保存したデータがあればそれを、なければ空文字
        });
      } else {
        // 名前はあるが保存データがない場合（新規従業員など）
        newRow.fill("", 0, header.length);
      }
      restoredValues.push(newRow);
    });

    sheet.getRange(2, 2, restoredValues.length, header.length).setValues(restoredValues);
    console.log("予定データを復元しました。");
  }

  // --- 5. 書式設定を適用 ---
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidths(2, header.length, 50);

  // 背景色を適用
  sheet.getRange(1, 2, sheet.getMaxRows(), header.length).setBackgrounds(backgrounds);

  // 入力規則（プルダウン）を設定
  if (sheet.getLastRow() > 1) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS).setAllowInvalid(false).build();
    sheet.getRange(2, 2, sheet.getLastRow() - 1, header.length).setDataValidation(rule);
  }
}

/**
 * 日付の配列からヘッダーと背景色のデータを生成するヘルパー関数
 * @param {Date[]} dates - 日付オブジェクトの配列
 * @return {{header: string[], backgrounds: (string|null)[][]}}
 */
function generateCalendarData(dates) {
  const holidayCalendar = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
  const holidays = holidayCalendar.getEvents(dates[0], dates[dates.length - 1])
    .map(e => new Date(e.getStartTime()).setHours(0, 0, 0, 0));
  const holidaySet = new Set(holidays);

  const header = [];
  const singleRowBackgrounds = [];
  const weekdays = ["(日)", "(月)", "(火)", "(水)", "(木)", "(金)", "(土)"];

  dates.forEach(date => {
    header.push(`${date.getMonth() + 1}/${date.getDate()}\n${weekdays[date.getDay()]}`);
    const dateWithoutTime = new Date(date).setHours(0, 0, 0, 0);
    if (date.getDay() === 0 || date.getDay() === 6 || holidaySet.has(dateWithoutTime)) {
      singleRowBackgrounds.push(WEEKEND_AND_HOLIDAY_COLOR);
    } else {
      singleRowBackgrounds.push(null); // 色なし
    }
  });

  // シートの最大行数分の背景色データを作成
  const backgrounds = Array(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SS_NAME)?.getMaxRows() || 1000)
    .fill(null).map(() => [...singleRowBackgrounds]);
  
  return { header, backgrounds };
}