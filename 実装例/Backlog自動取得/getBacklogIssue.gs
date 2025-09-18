// ▼▼▼ 設定項目 ▼▼▼

// (1) BacklogスペースのURL
const SPACE_URL = 'https://*******.backlog.com';

// (2) Backlog APIキー
const API_KEY = '******************************************';

// (3) 対象のプロジェクトID
const PROJECT_ID = ******;

// (4) 取得したい課題のステータスIDの配列 (空の場合は完了済みも含む全て)
const TARGET_STATUS_IDS = [];

// (5) スプレッドシートに出力するヘッダー
const HEADERS = [
  "ID",
  "課題キー",
  "件名",
  "サイト種別",
  "完了予定日",
  "実施ステータス",
];

// 'スプレッドシートの列名': 先ほど取得したカスタムフィールドのID(数値)
const CUSTOM_FIELD_ID_MAPPING = {
  'サイト種別': ******
};

// (7) 操作対象のスプレッドシート名
const TARGET_SHEET_NAME = 'Backlog連携テスト';

// ▲▲▲ 設定項目はここまで ▲▲▲


function syncBacklogIssues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (!sheet) {
    Logger.log(`エラー: 指定されたシート "${TARGET_SHEET_NAME}" が見つかりません。`);
    SpreadsheetApp.getUi().alert(`エラー: シート "${TARGET_SHEET_NAME}" が見つかりませんでした。`);
    return;
  }

  // --- 1. 既存データを読み込み（API対象外の課題を維持するため）---
  const existingData = sheet.getDataRange().getValues();
  const headerRow = existingData.length > 0 ? existingData[0] : HEADERS;
  const colIndexMap = createHeaderIndexMap(headerRow);

  const existingIssuesMap = new Map();
  if (existingData.length > 1) {
    const keyCol = colIndexMap['課題キー'];
    if (keyCol === undefined) {
      SpreadsheetApp.getUi().alert('エラー: スプレッドシートのヘッダーに "課題キー" が見つかりませんでした。');
      return;
    }
    for (let i = 1; i < existingData.length; i++) {
      const row = existingData[i];
      const issueKey = row[keyCol];
      if (issueKey) {
        existingIssuesMap.set(issueKey, row);
      }
    }
  }

  // --- 2. Backlog APIから最新の課題データを取得 ---
  const query = { 'projectId[]': [PROJECT_ID], 'sort': 'id', 'order': 'asc', 'count': 100 };
  if (TARGET_STATUS_IDS.length > 0) {
    query['statusId[]'] = TARGET_STATUS_IDS;
  }
  const latestIssues = getIssues(query);

  // --- 3. 最新データと既存データをマージして、新しいシートデータを作成 ---
  const outputData = [HEADERS];
  const processedKeys = new Set();

  for (const issue of latestIssues) {
    // 課題キーが存在しないアイテムは、課題として扱わずスキップする
    if (!issue.issueKey) {
      continue;
    }

    const newRow = Array(HEADERS.length).fill('');

    // HEADERSの順番通りに直接データを格納します
    newRow[0] = issue.id; // ID
    newRow[1] = issue.issueKey; // 課題キー
    newRow[2] = issue.summary; // 件名
    // ★カスタムフィールドIDを使って "サイト種別" の値を取得
    newRow[3] = getCustomFieldValueById(issue.customFields, CUSTOM_FIELD_ID_MAPPING['サイト種別']);
    newRow[4] = issue.dueDate ? issue.dueDate.split('T')[0] : ''; // 完了予定日
    newRow[5] = issue.status ? issue.status.name : ''; // 実施ステータス

    outputData.push(newRow);
    processedKeys.add(issue.issueKey);
  }

  // --- 4. APIの取得対象外になったが、シートには残っている課題を追加 ---
  for (const [key, row] of existingIssuesMap) {
    if (!processedKeys.has(key)) {
      outputData.push(row);
    }
  }

  // --- 5. スプレッドシートにデータを一括書き込み ---
  if (outputData.length > 0) {
    sheet.clearContents();
    sheet.getRange(1, 1, outputData.length, HEADERS.length).setValues(outputData);
  }

  // --- 6. シートの書式設定を適用 ---
  setupSheetFormatting(sheet);
}


/**
 * スプレッドシートの書式（条件付き書式）を設定します。
 */
function setupSheetFormatting(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexMap = createHeaderIndexMap(headerRow);

  // 実施ステータスの条件付き書式設定
  const implStatusCol = colIndexMap['実施ステータス'];
  if (implStatusCol !== undefined) {
    const range = sheet.getRange(2, implStatusCol + 1, lastRow - 1, 1);
    
    sheet.clearConditionalFormatRules();
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('未対応').setBackground('#f4cccc').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('保留').setBackground('#fce5cd').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('処理中').setBackground('#cfe2f3').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('対応不要').setBackground('#666666').setFontColor('#ffffff').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('完了').setBackground('#d9ead3').setRanges([range]).build()
    ];
    sheet.setConditionalFormatRules(rules);
  }
}

/**
 * ヘッダー名から列番号（0始まり）へのマップを作成します。
 */
function createHeaderIndexMap(headerRow) {
  const map = {};
  headerRow.forEach((header, i) => {
    map[header] = i;
  });
  return map;
}

function getCustomFieldValueById(customFields, fieldId) {
  if (!customFields || !Array.isArray(customFields) || !fieldId) {
    return '';
  }

  const field = customFields.find(cf => cf.id === fieldId);

  if (field && typeof field.value !== 'undefined' && field.value !== null) {
    // 値が配列（チェックボックスなど）の場合
    if (Array.isArray(field.value)) {
      // 各要素の.nameプロパティを抽出し、カンマ区切りで結合
      return field.value.map(item => item.name || '').join(', ');
    }
    // 値がオブジェクト（単一選択リストなど）の場合
    if (typeof field.value === 'object') {
      // .nameプロパティを返す
      return field.value.name || '';
    }
    // 値が文字列や数値などの単純な値の場合
    return field.value;
  }

  return '';
}


// ----------------------------------------------------------------
// 以下はAPI通信のための関数群
// ----------------------------------------------------------------

function getIssues(query = {}) {
  query.offset = 0;
  query.count = 100;
  query.apiKey = API_KEY;
  return getAllPage(_getIssues, query);
}

function _getIssues(query) {
  const options = { 'method': 'GET', 'muteHttpExceptions': true };
  const response = UrlFetchApp.fetch(buildUrl(SPACE_URL, '/api/v2/issues', query), options);
  const statusCode = response.getResponseCode();
  if (statusCode === 200) {
    return JSON.parse(response.getContentText());
  } else {
    Logger.log('エラー: ステータスコード ' + statusCode);
    Logger.log('レスポンス: ' + response.getContentText());
    throw new Error('APIリクエストに失敗しました。詳細はログを確認してください。');
  }
}

function getAllPage(f, query) {
  let allResponses = [];
  let response;
  do {
    response = f(query);
    allResponses = allResponses.concat(response);
    query.offset += query.count;
  } while (response.length === query.count);
  return allResponses;
}

function buildUrl(spaceUrl, path, query) {
  return spaceUrl + path + '?' + objectToQueryString(query);
}

function objectToQueryString(obj) {
  return Object.keys(obj)
    .map(function (key) {
      if (key.endsWith("[]")) {
        return obj[key].map(function (value) {
          return key + "=" + encodeURIComponent(value);
        }).join("&");
      } else {
        return key + "=" + encodeURIComponent(obj[key]);
      }
    })
    .join("&");
}

