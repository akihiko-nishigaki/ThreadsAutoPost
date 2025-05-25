function testgetAttachmentImageMovies() {
  let test = getAttachmentImageMovies(7);
}

function inspectOAuthService() {
  // サービスインスタンスを取得
  const service = getServiceThreads();
  
  // 基本的なサービス情報を取得
  const serviceInfo = {
    // サービスの認証状態を確認
    isAuthorized: service.hasAccess(),
    
    // 現在のアクセストークン
    accessToken: service.getAccessToken(),
    
    // トークンの有効期限
    tokenExpiry: PropertiesService.getUserProperties().getProperty('oauth2.threads.expires'),
    
    // リフレッシュトークン
    refreshToken: PropertiesService.getUserProperties().getProperty('oauth2.threads.refresh_token'),
    
    // 設定値を取得
    config: {
      authorizationUrl: THREADS_AUTH_URL,
      tokenUrl: THREADS_TOKEN_URL,
      clientId: THREADS_CLIENT_ID,
      scope: THREADS_SCOPE,
    }
  };
  
  // ログに出力
  Logger.log('Service Authorization Status: ' + serviceInfo.isAuthorized);
  Logger.log('Access Token: ' + serviceInfo.accessToken);
  Logger.log('Token Expiry: ' + new Date(parseInt(serviceInfo.tokenExpiry)));
  Logger.log('Has Refresh Token: ' + (serviceInfo.refreshToken ? 'Yes' : 'No'));
  Logger.log('Service Configuration:', serviceInfo.config);
  
  return serviceInfo;
}

function resetOAuthService() {
  // サービスの認証情報をリセット
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('oauth2.threads.access_token');
  userProperties.deleteProperty('oauth2.threads.refresh_token');
  userProperties.deleteProperty('oauth2.threads.expires');
  
  Logger.log('OAuth service credentials have been reset');
}

function forceTokenRefresh() {
  const service = getServiceThreads();
  
  // 現在のトークンを削除して強制的に更新
  PropertiesService.getUserProperties().deleteProperty('oauth2.threads.access_token');
  
  // 新しいトークンを取得
  const newToken = service.getAccessToken();
  Logger.log('New access token obtained: ' + newToken);
  
  return newToken;
}



/**
 * Google DriveのファイルIDからダイレクトダウンロードリンクを生成します
 * 
 * @param {string} fileId - Google DriveのファイルID
 * @return {string} ダイレクトダウンロード用URL
 */
function generateDirectDownloadUrl(fileId) {
  // ベースとなるGoogle DriveのダウンロードURL
  const baseUrl = 'https://drive.google.com/uc';
  
  // URLSearchParamsを使用してクエリパラメータを構築
  const params = {
    'export': 'download',
    'id': fileId
  };
  
  // クエリパラメータを連結
  const queryString = Object.entries(params)
    .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
    .join('&');
  
  // 最終的なURLを生成
  const downloadUrl = `${baseUrl}?${queryString}`;
  
  return downloadUrl;
}

function showSelectionDialog() {
  // カスタムダイアログを表示するための HTML を読み込む
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setWidth(400)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, '行の選択');
}

/***********************************************
 * 選択されている行番号を返す
 ***********************************************/
function getSourceRowNumber() {
  // アクティブなセルの行番号を取得
  const activeCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  return activeCell.getRow();
}

function testfunc(){
  copyValuesToActiveRow(1)
}

/***********************************************
 * 選択された行の引用情報を、アクティブ行へ返す
 ***********************************************/
function copyValuesToActiveRow(sourceRowNum) {
  try {
    Logger.log('Function started with sourceRowNum: ' + sourceRowNum); // 開始ログ
    
    const sheet = SpreadsheetApp.getActiveSheet();
    Logger.log('Active sheet: ' + sheet.getName()); // シート名をログ
    
    const activeRow = sheet.getActiveCell().getRow();
    Logger.log('Active row: ' + activeRow); // アクティブな行をログ
    
    // Q列とU列の値を取得
    let targetRow = Number(sourceRowNum) + CONFIG.START_ROW - 1;
    let targetXCol = CONFIG.X_POST_URL + CONFIG.SHEET_ARRAY_COL_DIF;
    let targetThreadsCol = CONFIG.POST_ID_COL + CONFIG.SHEET_ARRAY_COL_DIF;

    const qValue = sheet.getRange(targetRow, targetXCol).getValue();
    const uValue = sheet.getRange(targetRow, targetThreadsCol).getValue();
    
    Logger.log('Retrieved values - Q: ' + qValue + ', U: ' + uValue); // 取得した値をログ
    
    // I列とJ列に値をセット
    sheet.getRange(activeRow, CONFIG.POST_QUOTE + CONFIG.SHEET_ARRAY_COL_DIF).setValue(qValue);
    sheet.getRange(activeRow, CONFIG.POST_QUOTE_THREADS + CONFIG.SHEET_ARRAY_COL_DIF).setValue(uValue);
    
    Logger.log('Values set successfully'); // 設定完了ログ
    
    return {
      success: true,
      message: '値のコピーが完了しました'
    };
  } catch (error) {
    Logger.log('Error occurred: ' + error.toString()); // エラーログ
    return {
      success: false,
      message: 'エラーが発生しました: ' + error.toString()
    };
  }
}
