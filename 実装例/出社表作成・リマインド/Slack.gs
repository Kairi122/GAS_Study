// -------- 設定項目 --------

// ▼ 1. SlackのIncoming Webhook URL
const SLACK_WEBHOOK_URL = "*********";

// ▼ 2. 出社表のスプレッドシートID
const SPREADSHEET_ID = "**********";

// ▼ 3. 出社表のシート名
const SHEET_NAME = "出社表";

// ▼ 4. 通知の対象とするステータス
const TARGET_STATUSES = ["出社", "有給", "研修"];

// -------- 設定はここまで --------


/**
メインの関数
 */
function notifyTodaysAttendanceToSlack() {
  try {
    const today = new Date();
    const todayString = `${today.getMonth() + 1}/${today.getDate()}`;

    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    if (!sheet) {
      console.error("指定されたシートが見つかりません。");
      return;
    }
    const dataRange = sheet.getDataRange();
    const allValues = dataRange.getValues();
    
    const header = allValues[0]; // 1行目（日付）
    const employees = allValues.slice(1); // 2行目以降（従業員データ）

    let todayColumnIndex = -1;
    for (let i = 1; i < header.length; i++) {
      if (header[i].startsWith(todayString)) {
        todayColumnIndex = i;
        break;
      }
    }

    if (todayColumnIndex === -1) {
      console.log(`本日の日付(${todayString})の列が見つからなかったため、通知をスキップします。`);
      return;
    }
    const targetMembers = {};
    TARGET_STATUSES.forEach(status => {
      targetMembers[status] = [];
    });

    employees.forEach(row => {
      const name = row[0];
      const status = row[todayColumnIndex];
      if (name && TARGET_STATUSES.includes(status)) {
        targetMembers[status].push(name);
      }
    });

    // --- Slackに送信するメッセージを作成 ---
    const message = createSlackMessage(today, targetMembers);
    if (message) {
      sendToSlack(message);
    } else {
      console.log("通知対象者がいなかったため、Slackへの通知は行いませんでした。");
    }

  } catch (e) {
    console.error(`エラーが発生しました: ${e}`);
    // エラー発生をSlackに通知することも可能
    // sendToSlack(`出社表のSlack通知処理でエラーが発生しました。\n\`\`\`${e}\`\`\``);
  }
}

/**
 * Slackに送信するメッセージ本文を組み立てる関数
 * @param {Date} date - 今日の日付
 * @param {Object} members - ステータスごとの従業員リスト
 * @return {string | null} - Slackに送信するメッセージ、または通知対象がいない場合はnull
 */
function createSlackMessage(date, members) {
  const weekdays = ["日", "月", "火", "水", "木", "金", "土"];
  const formattedDate = `${date.getMonth() + 1}月${date.getDate()}日 (${weekdays[date.getDay()]})`;

  let message = `おはようございます！\n本日の出社・休暇予定です。\n\n`;
  let hasTarget = false; // 通知対象者がいるかどうかのフラグ

  TARGET_STATUSES.forEach(status => {
    if (members[status] && members[status].length > 0) {
      hasTarget = true;
      message += `■ ${status}\n`;
      message += "```\n" + members[status].join("\n") + "\n```\n";
    }
  });

  return hasTarget ? message : null;
}

/**
 * Slackにメッセージを送信する関数
 * @param {string} text - 送信するメッセージ
 */
function sendToSlack(text) {
  const payload = {
    "text": text
  };
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };
  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
}