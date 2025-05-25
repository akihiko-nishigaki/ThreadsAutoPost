/**
 * 投稿処理を実行する関数
 * @param {Object} data - 投稿データ
 * @returns {Object} 投稿結果
 */
function submitPost(data) {
  try {
    const webPost = new WebPost();
    return webPost.submitPost(data);
  } catch (error) {
    console.error('投稿処理でエラーが発生しました:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * ファイルアップロード処理
 * @param {string} base64Data - Base64エンコードされたファイルデータ
 * @param {string} fileName - ファイル名
 * @param {string} mimeType - MIMEタイプ
 * @returns {Object} アップロード結果
 */
function uploadFile(base64Data, fileName, mimeType) {
  try {
    // Google Driveにファイルをアップロード
    const folder = DriveApp.getFolderById(CONFIG.GOOGLE_DRIVE.FOLDER_ID);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    
    return {
      success: true,
      url: file.getUrl()
    };
  } catch (error) {
    console.error('ファイルアップロードエラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フロントエンドからの投稿リクエストを処理する
 * @param {Object} postData - 投稿データ
 * @return {Object} 投稿結果
 */
function submitPost(postData) {
  try {
    Logger.log('submitPost開始');
    
    // スプレッドシートを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEETS_NAME.RESERVATION);
    
    // F列のデータを取得（7行目から1000行まで）
    const fColumnData = sheet.getRange("F7:F1000").getValues();
    
    // 7行目以降から空きを探す
    let targetRow = 7;  // 7行目から開始
    for (let i = 0; i < fColumnData.length; i++) {
      if (fColumnData[i][0] === "") {
        targetRow = i + 7;  // 7行目からのインデックスを考慮
        break;
      }
    }
    
    // 投稿データを準備
    const rowData = new Array(30).fill(''); // 十分な長さの配列を確保
    
    // 日時情報を設定
    const postDate = new Date(postData.date);
    postDate.setHours(parseInt(postData.hour), parseInt(postData.minute));
    
    // 基本情報を設定
    rowData[2] = postDate;  // C列: 投稿日時
    rowData[3] = postData.hour;  // D列: 時間
    rowData[4] = postData.minute;  // E列: 分
    rowData[5] = postData.text;  // F列: 文章
    
    // 投稿先のチェックを設定
    rowData[10] = postData.platforms.includes('x');  // K列: X
    rowData[11] = postData.platforms.includes('threads');  // L列: Threads
    rowData[12] = postData.platforms.includes('instagram');  // M列: Instagram
    
    // ファイルのアップロード処理
    if (postData.files && postData.files.length > 0) {
      let fileCount = 0;
      for (const file of postData.files) {
        if (fileCount >= 10) break; // 最大10ファイルまで
        
        // URLとチェックを設定
        const urlColumn = 24 + (fileCount * 2);  // Y列から始まる（24 = Y列のインデックス）
        const checkColumn = urlColumn + 1;  // チェック列はURL列の右隣
        
        // シートに直接書き込み
        sheet.getRange(targetRow, urlColumn + 1).setValue(file.url);  // URLを直接設定
        sheet.getRange(targetRow, checkColumn + 1).setValue(fileCount < 4);  // 4つ目まではチェックオン
        
        fileCount++;
      }
    }
    
    // データをシートに書き込み（G列を除く）
    const columnsToUpdate = [3, 4, 5, 6, 11, 12, 13];  // C, D, E, F, K, L, M列
    for (let i = 0; i < columnsToUpdate.length; i++) {
      const col = columnsToUpdate[i];
      sheet.getRange(targetRow, col).setValue(rowData[col - 1]);
    }
    
    return {
      success: true,
      message: '投稿が予約投稿シートに保存されました'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
} 