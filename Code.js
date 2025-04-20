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