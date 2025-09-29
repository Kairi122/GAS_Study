// ====== 設定項目 ======

// SlackのIncoming Webhook URL
const SLACK_WEBHOOK_URL = '*********';

const SHEET_NAME = 'シート1'; //

// ====== 設定はここまで ======


/**
 * メインの処理を実行する関数
 */
function checkTaskDeadlineAndSendReminder() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      console.error(`シート「${SHEET_NAME}」が見つかりません。`);
      return;
    }

// I2セルからタスク期限を取得
const deadlineValue = sheet.getRange('I2').getValue();
const deadline = new Date(deadlineValue);

if (isNaN(deadline.getTime())) {
  console.error(`I2セルの値「${deadlineValue}」が有効な日付ではありません。セルの表示形式を確認してください。`);
  return;
}

    deadline.setHours(0, 0, 0, 0);

    const reminderDate = new Date(deadline);
    reminderDate.setDate(deadline.getDate() - 3);//リマインドする日付（三日前）

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (today.getTime() !== reminderDate.getTime()) {
      console.log('本日はリマインドの送信日ではありません。');
      return; // リマインド送信日でなければ処理を終了
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      console.log('データがありません。');
      return;
    }

    const dataRange = sheet.getRange(2, 3, lastRow - 1, 3);
    const values = dataRange.getValues();


    const taskTitle = sheet.getRange('A1').getValue() || 'タスク'

    const incompleteUsers = [];
    values.forEach(row => {
      const slackName = row[0]; // C列 (配列のインデックスは0)
      const status = row[2];    // E列 (配列のインデックスは2)

      if (slackName && status !== '完了') {
        incompleteUsers.push(slackName);
      }
    });

    // 未完了者がいればSlackに通知
    if (incompleteUsers.length > 0) {
      // メンション部分を作成
      const mentions = incompleteUsers.map(name => `<@${name}>`).join(' ');

      // Slackに送信するメッセージを作成
      const message = `【タスク期限リマインド：3日前】\n` +
                      `担当者の皆さん、タスク「${taskTitle}」の期限が迫っています！\n\n` +
                      `対象者: ${mentions} さん\n` +
                      `ステータスが「完了」となっていない方は、ご対応をお願いします。\n` +
                      `期限: ${deadline.getFullYear()}/${deadline.getMonth() + 1}/${deadline.getDate()}`;

      sendToSlack(message);
      console.log(`リマインドを送信しました: ${incompleteUsers.join(', ')}`);
    } else {
      console.log('全員タスク完了済みです。');
    }

  } catch (e) {
    console.error(`エラーが発生しました: ${e.message}`);
    // エラー発生をSlackに通知することも可能
    // sendToSlack(`リマインダーBOTでエラーが発生しました。\n${e.message}`);
  }
}

/**
 * Slackにメッセージを送信する関数
 * @param {string} message 送信するテキストメッセージ
 */
function sendToSlack(message) {
  const payload = {
    "text": message,
    "link_names": 1, // これにより <@username> 形式のメンションが有効になる
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload)
  };

  UrlFetchApp.fetch(SLACK_WEBHOOK_URL, options);
}