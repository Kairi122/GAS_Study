// Backlogの課題からカスタム属性の値を取得する関数
function getBacklogCustomField() {

  const SPACE_ID = '*******';             // 指定されたスペースID
  const API_KEY = '*********************';
  const ISSUE_KEY = '**********';
  const CUSTOM_FIELD_ID = ******;             // 指定されたカスタム属性ID
  // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

  // Backlog APIのエンドポイントURLを作成
  const url = `https://${SPACE_ID}.backlog.com/api/v2/issues/${ISSUE_KEY}?apiKey=${API_KEY}`;

  try {
    const response = UrlFetchApp.fetch(url);
    const issueData = JSON.parse(response.getContentText());
    const customFields = issueData.customFields;

    if (customFields && customFields.length > 0) {
      // 指定したIDのカスタム属性を探す
      const targetField = customFields.find(field => field.id === CUSTOM_FIELD_ID);

      if (targetField && targetField.value) {
        let fieldValue = '';

        if (Array.isArray(targetField.value)) {
          fieldValue = targetField.value.map(item => item.name).join(', ');
        } else if (typeof targetField.value === 'object') {
          fieldValue = targetField.value.name;
        } else {
          fieldValue = targetField.value;
        }

        console.log(`課題「${ISSUE_KEY}」のカスタム属性(ID: ${CUSTOM_FIELD_ID})の値は「${fieldValue}」です。`);
        return fieldValue;

      } else {
// (以下、省略)
        console.log(`課題「${ISSUE_KEY}」に、指定されたカスタム属性(ID: ${CUSTOM_FIELD_ID})が見つかりませんでした。`);
      }
    } else {
      console.log(`課題「${ISSUE_KEY}」にカスタム属性が設定されていません。`);
    }
  } catch (e) {
    console.error('APIリクエストに失敗しました: ' + e.toString());
  }
}