function getCustomFieldList() {
  // --- 既存の設定項目をそのまま利用します ---
  const projectId = PROJECT_ID;
  const apiKey = API_KEY;
  const spaceUrl = SPACE_URL;
  // -----------------------------------------

  const path = `/api/v2/projects/${projectId}/customFields`;
  const url = `${spaceUrl}${path}?apiKey=${apiKey}`;

  const options = {
    'method': 'GET',
    'muteHttpExceptions': true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();

    if (statusCode === 200) {
      const customFields = JSON.parse(response.getContentText());

      if (customFields.length === 0) {
        Logger.log('このプロジェクトにはカスタムフィールドが設定されていません。');
        return;
      }

      Logger.log('--- カスタムフィールド一覧 ---');
      customFields.forEach(field => {
        Logger.log(`フィールド名: "${field.name}", ID: ${field.id}`);
      });
      Logger.log('--------------------------');

    } else {
      Logger.log(`エラー: APIリクエストに失敗しました。ステータスコード: ${statusCode}`);
      Logger.log(`レスポンス: ${response.getContentText()}`);
    }
  } catch (e) {
    Logger.log(`予期せぬエラーが発生しました: ${e.message}`);
  }
}