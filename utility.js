/**
 * JSONレスポンスから最初のIDを取得する
 * @param {string} jsonResponse - JSONレスポンス文字列
 * @return {string} 最初のID。見つからない場合は空文字列を返す
 */
function getFirstIdFromResponse(jsonResponse) {
  try {
    // 文字列がJSONの場合はパース、オブジェクトの場合はそのまま使用
    const data = typeof jsonResponse === 'string' ? 
      JSON.parse(jsonResponse) : jsonResponse;
    
    // データ配列が存在し、少なくとも1つの要素がある場合
    if (data.data && Array.isArray(data.data) && data.data.length > 0) {
      const firstId = data.data[0].id;
      return firstId || '';
    }
    
    return '';
  } catch (error) {
    console.error('JSONの解析中にエラーが発生しました:', error);
    return '';
  }
}

/**
 * ログを出力する関数
 * @param {string} message - ログメッセージ
 * @param {string} [category='INFO'] - ログカテゴリ
 */
function logMessage(message, category = 'INFO') {
  const timestamp = new Date().toISOString();
  Logger.log(`[${timestamp}] [${category}] ${message}`);
}


/**
 * メンション情報変換
 * @param {string} text - 投稿文章
 * @param {string} snsCol - 指定したSNSのカラム番号
 * @returns {string} - 変換後の投稿文章
 */
function convertMentionInfo(text, snsCol) {

  // メンションシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.MENTION);
  const lastRow = sheet.getLastRow();

  // メンション情報を取得
  const mentionData = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  // 投稿文章のサンプル
  let transformedPost = text;

  // メンションデータを使用して変換
  mentionData.forEach(row => {
    const keyword = row[SNS_CONVERT.COL.KEYWORD]; // キーワード
    const name = row[SNS_CONVERT.COL.NAME];   // 名前
    const handle = row[snsCol]; // 指定したSNSカラムのハンドル

    if (keyword && name && handle) {
      const mentionTag = `@${handle}`;
      transformedPost = transformedPost.replace(
        new RegExp(`\\${keyword}`, 'g'),
        `${name}さん(${mentionTag})`
      );
    }
  });

  // 結果をログに出力
  Logger.log("## 変換前\n" + text);
  Logger.log("## 変換後\n" + transformedPost);

  return transformedPost;
}

/***************************************
 * システムプロパティ取得
 ***************************************/
function getSystemProperty(){
  
  // システムシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.SYSTEM);

    // プロパティを取得する
  let uploadImageFoloderId = sheet.getRange(PROPERTY_CELL.UL_IMAGE_FOLDER_ID).getValue();
  let uploadVideoFoloderId = sheet.getRange(PROPERTY_CELL.UL_VIDEO_FOLDER_ID).getValue();
  let xApiClient = sheet.getRange(PROPERTY_CELL.X_CLIENT_KEY).getValue();
  let xApiClientSecret = sheet.getRange(PROPERTY_CELL.X_CLIENT_SECRET).getValue();
  let xApiCodeVerifier = sheet.getRange(PROPERTY_CELL.X_CODE_VERIFIER).getValue();
  let xApiOauth2 = sheet.getRange(PROPERTY_CELL.X_OAUTH2_TWITTER).getValue();
  let xUserId = sheet.getRange(PROPERTY_CELL.X_USER_ID).getValue();
  let instaAppId = sheet.getRange(PROPERTY_CELL.INSTA_APP_ID).getValue();
  let instaAppSecret = sheet.getRange(PROPERTY_CELL.INSTA_APP_SECRET).getValue();
  let instaUserId = sheet.getRange(PROPERTY_CELL.INSTA_USER_ID).getValue();
  let instaBusinessId = sheet.getRange(PROPERTY_CELL.INSTA_BUSINESS_ID).getValue();
  let instaShortTimeToken = sheet.getRange(PROPERTY_CELL.INSTA_SHORT_ACCESS_TOKEN).getValue();
  let instaLongTimeToken = sheet.getRange(PROPERTY_CELL.INSTA_LONG_ACCESS_TOKEN).getValue();
  let instaLongTimeTokenExpiry = sheet.getRange(PROPERTY_CELL.INSTA_LONG_ACCESS_TOKEN_EXPIRY).getValue();
  let threadsLongTimeToken = sheet.getRange(PROPERTY_CELL.THREADS_LONG_TIME_TOKEN).getValue();
  let threadsLongTimeTokenExpiry = sheet.getRange(PROPERTY_CELL.THREADS_LONG_TIME_TOKEN_EXPIRY).getValue();
  let threadsClientId = sheet.getRange(PROPERTY_CELL.THREADS_CLIENT_ID).getValue();
  let threadsClientSecret = sheet.getRange(PROPERTY_CELL.THREADS_CLIENT_SECRET).getValue();
  let selectedSheetName = sheet.getRange(PROPERTY_CELL.SELECTED_SHEET_NAME).getValue();
  let webPostDefaultX = sheet.getRange(PROPERTY_CELL.WEB_POST_DEFAULT_X).getValue();
  let webPostDefaultThreads = sheet.getRange(PROPERTY_CELL.WEB_POST_DEFAULT_THREADS).getValue();
  let webPostDefaultInstagram = sheet.getRange(PROPERTY_CELL.WEB_POST_DEFAULT_INSTAGRAM).getValue();
  let cloudflareAccessKey = sheet.getRange(PROPERTY_CELL.CLOUD_FLARE_ACCESS_KEY).getValue();
  let cloudflareSecretKey = sheet.getRange(PROPERTY_CELL.CLOUD_FLARE_SECRET_KEY).getValue();
  let cloudflareAccountId = sheet.getRange(PROPERTY_CELL.CLOUD_FLARE_ACCOUNT_ID).getValue();
  let cloudflareBucket = sheet.getRange(PROPERTY_CELL.CLOUD_FLARE_BUCKET).getValue();
  let cloudflarePublicUrl = sheet.getRange(PROPERTY_CELL.CLOUD_FLARE_PUBLIC_URL).getValue();

  return{
    uploadImageFoloderId: uploadImageFoloderId,
    uploadVideoFoloderId: uploadVideoFoloderId,
    xApiClient: xApiClient,
    xApiClientSecret: xApiClientSecret,
    xApiCodeVerifier: xApiCodeVerifier,
    //xCodeVerifier: xApiCodeVerifier,
    xApiOauth2: xApiOauth2,
    //xOauth2Twitter: xApiOauth2,
    xUserId: xUserId,
    instaAppId: instaAppId,
    instaAppSecret: instaAppSecret,
    instaUserId: instaUserId,
    instaBusinessId: instaBusinessId,
    instaShortTimeToken: instaShortTimeToken,
    instaLongTimeToken: instaLongTimeToken,
    instaLongTimeTokenExpiry: instaLongTimeTokenExpiry,
    threadsLongTimeToken: threadsLongTimeToken,
    threadsLongTimeTokenExpiry: threadsLongTimeTokenExpiry,
    threadsClientId: threadsClientId,
    threadsClientSecret: threadsClientSecret,
    selectedSheetName: selectedSheetName,
    webPostDefaultX: webPostDefaultX,
    webPostDefaultThreads: webPostDefaultThreads,
    webPostDefaultInstagram: webPostDefaultInstagram,
    cloudflareAccessKey: cloudflareAccessKey,
    cloudflareSecretKey: cloudflareSecretKey,
    cloudflareAccountId: cloudflareAccountId,
    cloudflareBucket: cloudflareBucket,
    cloudflarePublicUrl: cloudflarePublicUrl
  }
}

/***************************************
 * システムプロパティ設定
 ***************************************/
function setSystemProperty(keyCell, value){
  
  // システムシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.SYSTEM);
  sheet.getRange(keyCell).setValue(value);

  Logger.log(`プロパティ設定しました。Key:${keyCell}、Value:${value}`);
}

/***************************************
 * アカウント設定済の値をセットする
 ***************************************/
function setSnsAccountSettingStatus(keyCell){

  // 予約投稿シートを取得
  const sheetReserve = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
  sheetReserve.getRange(keyCell).setValue("設定済");

  // 自動投稿シートを取得
  const sheetAuto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);
  sheetAuto.getRange(keyCell).setValue("設定済");

  Logger.log(`SNSステータスを設定しました。Key:${keyCell}、Value:設定済`);

}

/***************************************
 * SNSテンプレートチェックをオンにする
 ***************************************/
function setSnsCheck(keyCell){

  // システムシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.TEMPLATE);
  sheet.getRange(keyCell).setValue(true);
  const sheetAuto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.TEMPLATE_AUTO);
  sheetAuto.getRange(keyCell).setValue(true);

  Logger.log(`SNSテンプレートチェック。Key:${keyCell}、Value:チェックオン`);

}
/***************************************
 * アカウント設定済の値をセットする
 ***************************************/
function clearSnsAccountSettingStatus(keyCell){

  // システムシートを取得
  const sheetReserve = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
  sheetReserve.getRange(keyCell).setValue("未設定");

  const sheetAuto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);
  sheetAuto.getRange(keyCell).setValue("未設定");

  Logger.log(`SNSステータスを設定しました。Key:${keyCell}、Value:未設定`);

}

/***************************************
 * SNSテンプレートチェックをオフにする
 ***************************************/
function clearSnsCheck(keyCell){

  // システムシートを取得
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.TEMPLATE);
  sheet.getRange(keyCell).setValue(false);
  const sheetAuto = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.TEMPLATE_AUTO);
  sheetAuto.getRange(keyCell).setValue(false);

  Logger.log(`SNSテンプレートチェック。Key:${keyCell}、Value:チェックオフ`);

}

/**
 * URLから拡張子を抽出するユーティリティ関数
 * @param {string} url - 拡張子を抽出するURL
 * @return {string} 小文字の拡張子。拡張子がない場合は空文字を返す
 */
function getExtensionFromUrl(url) {
  // URLが未定義または空の場合は空文字を返す
  if (!url){
    return '';
  }
  
  // URLの最後のドット以降を取得
  const match = url.split('?')[0].match(/\.([^.]+)$/);
  
  // マッチした場合は小文字の拡張子を返す、そうでない場合は空文字を返す
  return match ? match[1].toLowerCase() : '';
}

/**
 * 拡張子からファイルタイプを判定する関数
 * @param {string} extension - 判定する拡張子
 * @return {string} ファイルタイプ（"image"または"video"）。該当しない場合は空文字を返す
 */
function getFileTypeFromExtension(extension) {
  // 拡張子が未定義または空の場合は空文字を返す
  if (!extension){
     return '';
  }
  
  // 小文字に変換して判定
  const lowerExtension = extension.toLowerCase();
  
  // 画像拡張子の判定
  if (CONFIG.IMAGE_EXTENSIONS.includes(lowerExtension)) {
    return CONFIG.STRING_IMAGE;
  }
  
  // 動画拡張子の判定
  if (CONFIG.VIDEO_EXTENSIONS.includes(lowerExtension)) {
    return CONFIG.STRING_VIDEO;
  }
  
  // いずれにも該当しない場合は空文字を返す
  return '';
}

/**
 * values配列内で上方向に検索を行う関数
 * @param {Array<Array>} values - スプレッドシートから取得した値の2次元配列
 * @param {number} currentRowIndex - 現在の行インデックス
 * @param {string|number} searchValue - 検索する値
 * @param {number} searchColIndex - 検索する列のインデックス
 * @param {number} resultColIndex - 結果を取得する列のインデックス
 * @param {number} headerRowIndex - ヘッダー行のインデックス
 * @param {string} sheetName 対象シート名
 * @returns {string|number|null} 見つかった場合はS列の値、見つからない場合はnull
 */
function searchUpwardsInValues(
  values,
  currentRowIndex,
  searchValue,
  searchColIndex,
  resultColIndex,
  headerRowIndex,
  sheetName
) {
  // 検索開始行を設定（現在の1つ上の行から）
  const startRow = currentRowIndex - headerRowIndex - 1;
  
  // 上方向に向かって検索
  for (let row = startRow; row >= 0; row--) {
    // 値の比較
    if (values[row][searchColIndex] === searchValue) {

      // 値を再度取得する
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const targetSheet = ss.getSheetByName(sheetName);

      // データを再取得
      const dataRange = targetSheet.getRange(row + headerRowIndex, 2, 1, 32);  // B-AG列を取得
      const reValues = dataRange.getValues();

      Logger.log(`row:${row}`)
      Logger.log(`resultColIndex:${resultColIndex}`)
      
      // let retVaeluesValue = values[row][resultColIndex];
      let retVaeluesValue = reValues[0][resultColIndex];
      Logger.log(`retValuesValue:${retVaeluesValue}`)

      return retVaeluesValue;
    }
  }
  
  return "";
}

function getVideoInfo(url) {
  try {

    // URLFetchAppを使用してHTTPリクエストを送信
    var options = {
      'method': 'get',
      'muteHttpExceptions': true,
      'followRedirects': true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    
    // レスポンスコードチェック
    if (responseCode !== 200) {
      throw new Error('Failed to fetch video. Status code: ' + responseCode);
    }
    
    // blobを取得
    var blob = response.getBlob();
    
    // 動画情報を取得
    var videoInfo = {
      fileName: getFileNameFromUrl(url),
      blob: blob,
      url: url,
      contentType: blob.getContentType(),
      fileBytes: blob.getBytes(),
      totalBytes: blob.getBytes().length,
      sizeMB: Math.round((blob.getBytes().length / (1024 * 1024)) * 100) / 100,
      responseCode: responseCode
    };
    
    return videoInfo;
    
  } catch (error) {
    Logger.log('Error: ' + error.message);
    return {
      error: error.message,
      responseCode: error.responseCode || null
    };
  }
}

// URLからファイル名を抽出する関数
function getFileNameFromUrl(url) {
  return url.split('/').pop();
}

/**
 * フォルダURLからフォルダIDを抽出する
 * @param {string} url フォルダURL
 * @return {string|null} フォルダID
 */
function getFolderIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * フォルダ内のファイルを名前の昇順で取得する
 * @param {Folder} folder DriveAppのフォルダオブジェクト
 * @return {File[]} ソート済みのファイル配列
 */
function getSortedFiles(folder) {
  const files = [];
  const fileIterator = folder.getFiles();
  
  // イテレータからファイルを配列に変換
  while (fileIterator.hasNext()) {
    files.push(fileIterator.next());
  }
  
  // ファイル名で昇順ソート
  return files.sort((a, b) => {
    const nameA = a.getName().toLowerCase();
    const nameB = b.getName().toLowerCase();
    return nameA.localeCompare(nameB);
  });
}



/**
 * 指定された行の値と数式をコピーする
 * 
 * @param {GoogleAppsScript.Spreadsheet.Range} sourceRange - コピー元の行の範囲
 * @param {GoogleAppsScript.Spreadsheet.Sheet} targetSheet - コピー先のシート
 * @param {number} targetRow - コピー先の行番号
 * 
 * @example
 * // 使用例
 * const sourceRange = templateSheet.getRange("2:2");
 * const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Target");
 * copyRowWithFormulas(sourceRange, targetSheet, 5); // 5行目にコピー
 */
function copyRowWithFormulas(sourceRange, targetSheet, targetRow) {
  // コピー先の範囲を取得
  const targetRange = targetSheet.getRange(targetRow, 1, 1, sourceRange.getNumColumns());
  
  // 数式を相対参照形式で取得
  const formulasR1C1 = sourceRange.getFormulasR1C1();
  
  // 値をコピー
  targetRange.setValues(sourceRange.getValues());
  
  // 数式をコピー（相対参照を保持）
  targetRange.setFormulasR1C1(formulasR1C1);
}


/**
 * 時間ユーティリティ関数
 * @param {Date} date 対象の日付
 * @returns {string} HH:mm形式の時間文字列
 */
function getTimeString(date) {
  return Utilities.formatDate(date, 'Asia/Tokyo', 'HH:mm');
}

/**
 * 時間比較用のオブジェクトを生成
 * @param {Date} date 対象の日付
 * @returns {Object} 時間と分を持つオブジェクト
 */
function createTimeObject(date) {
  return {
    hours: date.getHours(),
    minutes: date.getMinutes()
  };
}

/**
 * 2つの時間を比較して、指定時間が範囲内かチェック
 * @param {Object} target 対象時間
 * @param {Object} start 開始時間
 * @param {Object} end 終了時間
 * @returns {boolean} 範囲内の場合はtrue
 */
function isTimeInRange(target, start, end) {
  // 終了時間が開始時間より小さい場合（日付をまたぐ場合）の処理
  if (end.hours < start.hours || (end.hours === start.hours && end.minutes < start.minutes)) {
    return isTimeInRange24h(target, start, end);
  }
  
  // 通常の時間比較
  const targetMinutes = target.hours * 60 + target.minutes;
  const startMinutes = start.hours * 60 + start.minutes;
  const endMinutes = end.hours * 60 + end.minutes;
  
  return targetMinutes >= startMinutes && targetMinutes <= endMinutes;
}

/**
 * 日付をまたぐ場合の時間比較
 * @param {Object} target 対象時間
 * @param {Object} start 開始時間
 * @param {Object} end 終了時間
 * @returns {boolean} 範囲内の場合はtrue
 */
function isTimeInRange24h(target, start, end) {
  const targetMinutes = target.hours * 60 + target.minutes;
  const startMinutes = start.hours * 60 + start.minutes;
  const endMinutes = end.hours * 60 + end.minutes;
  
  return targetMinutes >= startMinutes || targetMinutes <= endMinutes;
}



/**
 * メールアドレスからローカル部分（@の前の部分）を取得する関数。
 *
 * @param {string} email メールアドレス（例: "user@example.com"）
 * @returns {string|null} ローカル部分（例: "user"）、無効な場合はnullを返す
 */
function getLocalPartFromEmail(email) {
  // メールアドレスに「@」が含まれているかを確認
  if (!email.includes("@")) {
    Logger.log("無効なメールアドレスです");
    return null;  // 無効な場合はnullを返す
  }

  // '@' 記号を基準に分割し、前半部分（ローカル部分）を取得
  var localPart = email.split("@")[0];

  // 取得したローカル部分をログに出力
  // Logger.log("メールアドレスのローカル部分: " + localPart);
  
  // ローカル部分を返す
  return localPart;
}

/**
 * 現在の日本時間（JST）をyyyymmdd形式の文字列で取得する
 * @return {string} yyyymmddHHmmss形式の日付文字列（例：20250128234510）
 */
function getJstDateString() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
}

// 別のフォーマットが必要な場合の実装例
/**
 * カスタムフォーマットで日本時間の日付文字列を取得する
 * @param {string} format - 日付フォーマット（例：'yyyy-MM-dd'）
 * @return {string} 指定されたフォーマットの日付文字列
 * 
 * 使用可能なフォーマットパターン：
 * - yyyy：年（4桁）
 * - MM：月（2桁）
 * - dd：日（2桁）
 * - HH：時（24時間形式、2桁）
 * - mm：分（2桁）
 * - ss：秒（2桁）
 */
function getJstDateStringWithFormat(format) {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', format);
}
