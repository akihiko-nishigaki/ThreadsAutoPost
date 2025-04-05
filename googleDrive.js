// 設定値の定数
const SYSTEM_SHEET_NAME = 'system';
const IMAGE_FOLDER_CELL = 'B2';
const VIDEO_FOLDER_CELL = 'B3';
const MAX_IMAGE_WIDTH = 1440; // 最大画像幅
const MAX_VIDEO_SIZE = 30 * 1024 * 1024; // 30MB
const IMAGE_FOLDER_NAME = '投稿用画像';
const VIDEO_FOLDER_NAME = '投稿用動画';

function showFileUploadDialog() {
  try {
    // 選択しているシートを記録する：予約投稿シート
    setSystemPropertyValue(PROPERTY_CELL.SELECTED_SHEET_NAME, SHEETS_NAME.RESERVATION);
    Logger.log('選択シートを設定: ' + SHEETS_NAME.RESERVATION);
    
    
    const html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(600)
      .setHeight(400);
    SpreadsheetApp.getUi().showModalDialog(html, 'ファイルアップロード');
  } catch (error) {
    console.error('ダイアログ表示エラー:', error);
    throw error;
  }
}

function showFileUploadDialogAutoSheet() {
  try {
    // 選択しているシートを記録する：予約投稿シート
    setSystemPropertyValue(PROPERTY_CELL.SELECTED_SHEET_NAME, SHEETS_NAME.AUTO);
    Logger.log('選択シートを設定: ' + SHEETS_NAME.AUTO);
    
    const html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(800)
      .setHeight(800);
    SpreadsheetApp.getUi().showModalDialog(html, 'ファイルアップロード');
  } catch (error) {
    console.error('ダイアログ表示エラー:', error);
    throw error;
  }
}


/**
 * フォルダIDを取得する関数
 */
function getFolderIds() {
  const ss = SpreadsheetApp.getActive();
  const systemSheet = ss.getSheetByName(SYSTEM_SHEET_NAME);
  
  if (!systemSheet) {
    throw new Error('systemシートが見つかりません');
  }

  const imageFolderId = systemSheet.getRange(IMAGE_FOLDER_CELL).getValue();
  const videoFolderId = systemSheet.getRange(VIDEO_FOLDER_CELL).getValue();

  if (!imageFolderId || !videoFolderId) {
    throw new Error('フォルダIDが設定されていません');
  }

  return {
    imageFolderId: imageFolderId,
    videoFolderId: videoFolderId
  };
}

/**
 * ファイルアップロード処理
 */
function uploadFile(base64Data, fileName, mimeType) {
  try {
    const folderIds = getFolderIds();
    const isImage = mimeType.startsWith('image/');
    const isVideo = mimeType.startsWith('video/');
    
    if (!isImage && !isVideo) {
      throw new Error('未対応のファイル形式です');
    }
    
    const folderId = isImage ? folderIds.imageFolderId : folderIds.videoFolderId;
    
    // base64をBlobに変換
    let blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    
    // ファイルサイズチェック
    if (isVideo && blob.getBytes().length > MAX_VIDEO_SIZE) {
      throw new Error('動画ファイルは30MB以内にしてください');
    }
    
    // 画像ファイルの場合のリサイズ処理
    let retIsSuccess;
    let retFileUrl;

    if (isImage) {
      let result = handleImageUpload(blob, folderId);

      retIsSuccess = result.success;
      retFileUrl = result.url;
      
    } else {
      // 動画ファイルの通常アップロード
      const folder = DriveApp.getFolderById(folderId);
      const file = folder.createFile(blob);
      const url = file.getUrl();

      // 現在選択中の行の添付ファイルセルへURLを書き込む
      setAttachmentFileUrl(url);

      retIsSuccess = true;
      retFileUrl = url;
    }

    console.log('retIsSuccess:'+retIsSuccess);
    console.log('retFileUrl:'+retFileUrl);

    return{
      success: retIsSuccess,
      url: retFileUrl
    }

    
  } catch (error) {
    console.error('Upload failed:', error);
    console.error('Error details:', error.stack);
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 画像アップロード処理
 */
function handleImageUpload(blob, folderId) {
  try {
    const size = ImgApp.getSize(blob);
    console.log('Checking image size:', blob.getName());
    console.log('Original size:', size.width, 'x', size.height);
    
    console.log('Image requires resizing');
    const folder = DriveApp.getFolderById(folderId);
    const tempFile = folder.createFile(blob);
    const tempFileId = tempFile.getId();

    let url = tempFile.getUrl();

    // 幅が制限を超える場合はリサイズ
    if (size.width > MAX_IMAGE_WIDTH) {
      try {
        const aspectRatio = size.height / size.width;
        const newWidth = MAX_IMAGE_WIDTH;
        const newHeight = Math.round(newWidth * aspectRatio);
        
        console.log('Resizing to:', newWidth, 'x', newHeight);
        
        const res = ImgApp.doResize(tempFileId, newWidth);
        const resizedFile = folder.createFile(res.blob.setName(blob.getName()));
        
        // 一時ファイルを削除
        tempFile.setTrashed(true);
        
        url = resizedFile.getUrl();
        console.log('Resize and upload completed successfully');
      } catch (resizeError) {
        tempFile.setTrashed(true);
        throw new Error('画像のリサイズに失敗しました: ' + resizeError.message);
      }
    }        

    try {
      // 現在選択中の行の添付ファイルセルへURLを書き込む
      setAttachmentFileUrl(url);
    } catch (error) {
      console.error('セルの更新に失敗しました:', error);
      // エラーが発生しても、アップロード自体は成功として扱う
    }

    return {
      success: true,
      url: url
    };
  } catch (error) {
    throw new Error('画像処理に失敗しました: ' + error.message);
  }
}


/***********************************
 * アップロードしたファイルのURLをセットする
 ***********************************/
function setAttachmentFileUrl(url) {
  try {
    // システムプロパティを取得
    const prop = getSystemProperty();
    Logger.log('取得したシステムプロパティ:', prop);
    Logger.log('選択中のシート名:', prop.xApiClientSecret);
    
    if (!prop || !prop.selectedSheetName) {
      console.error('選択中のシート名が取得できません');
      console.error('システムプロパティ:', prop);
      return;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(prop.selectedSheetName);
    if (!sheet) {
      console.error('シートが見つかりません:', prop.selectedSheetName);
      return;
    }

    // アクティブなセルが取得できない場合は、最終行の次の行を使用
    let currentRow;
    const activeCell = sheet.getActiveCell();
    if (!activeCell) {
      currentRow = sheet.getLastRow() + 1;
      console.log('アクティブなセルが見つからないため、最終行の次の行を使用:', currentRow);
    } else {
      currentRow = activeCell.getRowIndex();
    }

    let attachmentValues = sheet.getRange(currentRow
                                        , CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF
                                        , 1
                                        , CONFIG.ATTACH_END_COLUMN - CONFIG.ATTACH_START_COLUMN + 1).getValues();

    // Xの添付数を取得する
    let attachmentXCount = sheet.getRange(currentRow, CONFIG.ATTACH_COUNT_X_COL + CONFIG.SHEET_ARRAY_COL_DIF).getValue();

    // 添付ファイル領域の列へアップロードしたファイルのURLをセットする
    for(let currentCol = 0 ; currentCol < attachmentValues[0].length ; currentCol += CONFIG.ATTACH_COLUMN_DIF){
      // 添付ファイルが入っていない箇所へセットする
      if(attachmentValues[0][currentCol] == ""){
        sheet.getRange(currentRow, currentCol + CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF).setValue(url);

        // Xの添付数が4を超えていない場合はチェックを入れる
        if(attachmentXCount < 4){
          sheet.getRange(currentRow, currentCol + CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF + CONFIG.ATTACH_X_CHECK_DIF).setValue(true);
        }
        break;
      }
    }
    
    Logger.log('ファイルURLを設定完了:', url);
  } catch (error) {
    console.error('setAttachmentFileUrl でエラーが発生しました:', error);
    throw error;
  }
}

/***********************************
 * 初期処理：画像/動画フォルダ作成
 ***********************************/
function createUploadFolders() {
  try {
    // スプレッドシートのファイルIDを取得
    const spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    const spreadsheetFile = DriveApp.getFileById(spreadsheetId);
    const parentFolder = spreadsheetFile.getParents().next();
    
    // 既存のフォルダをチェック
    const existingFolders = parentFolder.getFolders();
    let imageFolder = null;
    let videoFolder = null;
    let imageFolderId = null;
    let videoFolderId = null;
    let existingFolderNames = [];
    
    while (existingFolders.hasNext()) {
      const folder = existingFolders.next();
      const folderName = folder.getName();
      
      if (folderName === IMAGE_FOLDER_NAME) {
        imageFolder = folder;
        imageFolderId = folder.getId();
        existingFolderNames.push(IMAGE_FOLDER_NAME);
      }
      if (folderName === VIDEO_FOLDER_NAME) {
        videoFolder = folder;
        videoFolderId = folder.getId();
        existingFolderNames.push(VIDEO_FOLDER_NAME);
      }
    }
    
    // 必要なフォルダを作成
    if (!imageFolder) {
      imageFolder = parentFolder.createFolder(IMAGE_FOLDER_NAME);
      imageFolderId = imageFolder.getId();
    }
    
    if (!videoFolder) {
      videoFolder = parentFolder.createFolder(VIDEO_FOLDER_NAME);
      videoFolderId = videoFolder.getId();
    }
    
    // systemシートの作成または取得
    let systemSheet = SpreadsheetApp.getActive().getSheetByName(SYSTEM_SHEET_NAME);
    if (!systemSheet) {
      systemSheet = SpreadsheetApp.getActive().insertSheet(SYSTEM_SHEET_NAME);
      
      // ヘッダーの設定
      systemSheet.getRange('A2').setValue('画像フォルダID');
      systemSheet.getRange('A3').setValue('動画フォルダID');
    }
    
    // フォルダIDの設定
    systemSheet.getRange(IMAGE_FOLDER_CELL).setValue(imageFolderId);
    systemSheet.getRange(VIDEO_FOLDER_CELL).setValue(videoFolderId);
    
    // 結果メッセージの作成
    let message = '';
    if (existingFolderNames.length > 0) {
      message += '以下のフォルダは既に存在しています：\n';
      message += existingFolderNames.map(name => `・${name}`).join('\n');
      message += '\n\n';
    }
    
    const newFolders = [];
    if (!existingFolderNames.includes(IMAGE_FOLDER_NAME)) newFolders.push(IMAGE_FOLDER_NAME);
    if (!existingFolderNames.includes(VIDEO_FOLDER_NAME)) newFolders.push(VIDEO_FOLDER_NAME);
    
    if (newFolders.length > 0) {
      message += '以下のフォルダを新規作成しました：\n';
      message += newFolders.map(name => `・${name}`).join('\n');
      message += '\n\n';
    }
    
    message += 'フォルダIDをsystemシートに設定しました。';
    
    // 結果の表示
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'フォルダ設定完了',
      message,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'エラー',
      `フォルダの作成中にエラーが発生しました：\n${error.toString()}`,
      ui.ButtonSet.OK
    );
  }
}

/**
 * URLからファイル拡張子を取得
 * @param {string} url - GoogleドライブのURL
 * @return {string} ファイル拡張子
 */
function getFileExtension(url) {
  const match = url.match(/\.[0-9a-z]+$/i);
  return match ? match[0].substring(1).toLowerCase() : '';
}

/**
 * ファイルの種類を判定
 * @param {string} extension - ファイル拡張子
 * @return {string} 'image' または 'video'
 */
function getFileCategory(extension) {
  if (CONFIG.IMAGE_EXTENSIONS.includes(extension)) return 'image';
  if (CONFIG.VIDEO_EXTENSIONS.includes(extension)) return 'video';
  return 'unknown';
}

/**
 * 指定行の添付ファイル情報を取得
 * @param {sheet} sheet - 予約投稿/自動投稿シート
 * @param {number} row - 処理対象の行番号
 * @return {Object} URLとファイルカテゴリを含むオブジェクト
 */
function getAttachmentImageMovies(sheet, row) {

  // const sheet = SpreadsheetApp.getActiveSheet();
  const startCol = CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF;
  const endCol = CONFIG.ATTACH_END_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF;
  
  // 戻り値となる配列を初期化
  const attachments = [];

  // 指定範囲のURLを取得
  for (let col = startCol; col <= endCol; col += CONFIG.ATTACH_COLUMN_DIF) {
    
    Logger.log('CurrentCol:' + col);
    
    const cellValue = sheet.getRange(row, col).getValue();
    
    // URLが存在する場合の処理
    if (cellValue && cellValue.toString().includes('drive.google.com')) {

      let fileInfo = getFileInfo(cellValue);
      
      const extension = fileInfo.extension;
      
      // 拡張子が見つからない場合はスキップ
      if (!extension){
        continue;
      } 
      
      const fileCategory = getFileCategory(extension);
      
      // 画像もしくは動画の場合のみ返す
      if (fileCategory !== 'unknown') {
        let wkUrl = "";

        if (fileCategory == CONFIG.STRING_VIDEO) {
          // ビデオの場合はファイルをcloudflareへアップロードする
          wkUrl = cloudflareMovieUpload(fileInfo.fileId[0]);
        }
        
        attachments.push({
          url: fileInfo.filedirecturl,
          movieDirectUrl: wkUrl,
          fileId: fileInfo.fileId[0],
          fileCategory: fileCategory
        });
      }
    }else if(cellValue){
      // URLの処理
      const extension = getExtensionFromUrl(cellValue);
      const fileType = getFileTypeFromExtension(extension);
      console.log(`Extension: ${extension}, Type: ${fileType}`);

      attachments.push({
        url: cellValue,
        fileId: "",
        fileCategory: fileType
      });
    }
  }
  
  // 該当するファイルが見つからない場合は空のオブジェクトを含む配列を返す
  return attachments;
}

/**
 * 指定行の添付ファイル情報を取得(X専用)
 * @param {sheet} sheet - 予約投稿/自動投稿シート
 * @param {number} row - 処理対象の行番号
 * @return {Object} URLとファイルカテゴリを含むオブジェクト
 */
function getAttachmentImageMoviesForX(sheet, row) {

  // const sheet = SpreadsheetApp.getActiveSheet();
  const startCol = CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF;
  const endCol = CONFIG.ATTACH_END_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF;
  
  // 戻り値となる配列を初期化
  const attachments = [];

  // 指定範囲のURLを取得
  for (let col = startCol; col <= endCol; col += CONFIG.ATTACH_COLUMN_DIF) {
    
    
    Logger.log('CurrentCol:' + col);
    const cellValue = sheet.getRange(row, col).getValue();
    const cellCheck = sheet.getRange(row, col+1).getValue();
    
    // URLが存在する場合の処理
    // Xはチェックがオンが入っている場合のみ処理を行う
    if (cellValue && cellCheck && cellValue.toString().includes('drive.google.com')) {

      let fileInfo = getFileInfo(cellValue);
      
      const extension = fileInfo.extension;
      
      // 拡張子が見つからない場合はスキップ
      if (!extension){
        continue;
      } 
      
      const fileCategory = getFileCategory(extension);
      
      // 画像もしくは動画の場合のみ返す
      if (fileCategory !== 'unknown') {
        attachments.push({
          url: fileInfo.filedirecturl,
          fileId: fileInfo.fileId[0],
          fileCategory: fileCategory
        });
      }
    }else if(cellValue && cellCheck){
      // URLの処理
      const extension = getExtensionFromUrl(cellValue);
      const fileType = getFileTypeFromExtension(extension);
      console.log(`Extension: ${extension}, Type: ${fileType}`);

      attachments.push({
        url: cellValue,
        fileId: "",
        fileCategory: fileType
      });
    }
  }
  
  // 該当するファイルが見つからない場合は空のオブジェクトを含む配列を返す
  return attachments;
}


/******************************************
 * GoogleDriveのURLからファイルIDを取得して、
 * 詳細情報を取得する
 ******************************************/
function getFileInfo(fileUrl) {
  try {
    // IDを抽出
    const fileId = fileUrl.match(/[-\w]{25,}/);
    
    if (!fileId) {
      throw new Error('Invalid Google Drive URL');
    }
    
    // ファイルを取得
    const file = DriveApp.getFileById(fileId[0]);
    
    let extension = file.getName().split('.').pop();

    let fileDirectUrl = convertDriveUrl(fileUrl, extension);

    return {
      name: file.getName(),
      extension: extension,
      fileId: fileId,
      filedirecturl: fileDirectUrl,
      mimeType: file.getMimeType()
    };
  } catch (error) {
    console.error('Error:', error.message);
    return null;
  }
}

/****************************************
 * GoogleDriveのファイルを直接参照させるURLを作成する
 ****************************************/
function convertDriveUrl(url, extension) {
  const fileId = url.match(/[-\w]{25,}/)[0];
  return `https://drive.google.com/uc?id=${fileId}&${extension}`;
}

/****************************************
 * GoogleDriveフォルダを指定して、フォルダ内の直接参照させるURLを作成する
 ****************************************/
function showAttachFileInfoForGoogleFolderInputDialog() {
 
   // 選択しているシートを記録する：予約投稿シート
  setSystemProperty(PROPERTY_CELL.SELECTED_SHEET_NAME, SHEETS_NAME.RESERVATION);

  // HTMLテンプレートを取得
  const html = HtmlService.createTemplateFromFile('googleFolderInputDialog')
    .evaluate()
    .setWidth(CONFIG.DIALOG_WIDTH)
    .setHeight(CONFIG.DIALOG_HEIGHT)
    .setTitle(CONFIG.DIALOG_TITLE);
  
  // モーダルダイアログとして表示
  SpreadsheetApp.getUi().showModalDialog(html, CONFIG.DIALOG_TITLE.GOOGLE_FOLDER_INPUT);
}

/****************************************
 * GoogleDriveフォルダを指定して、フォルダ内の直接参照させるURLを作成する
 ****************************************/
function showAttachFileInfoForGoogleFolderInputDialogAutoSheet() {
 
   // 選択しているシートを記録する：予約投稿シート
  setSystemProperty(PROPERTY_CELL.SELECTED_SHEET_NAME, SHEETS_NAME.AUTO);

  // HTMLテンプレートを取得
  const html = HtmlService.createTemplateFromFile('googleFolderInputDialog')
    .evaluate()
    .setWidth(CONFIG.DIALOG_WIDTH)
    .setHeight(CONFIG.DIALOG_HEIGHT)
    .setTitle(CONFIG.DIALOG_TITLE);
  
  // モーダルダイアログとして表示
  SpreadsheetApp.getUi().showModalDialog(html, CONFIG.DIALOG_TITLE.GOOGLE_FOLDER_INPUT);
}


// フォルダURL入力後の処理を実行
function processFilesInFolder(folderUrl) {
  
  //let folderUrl = "https://drive.google.com/drive/folders/16haggvysXOmnG_jokj9QkoFV3jAIgx6E";
  
  try {
    // フォルダURLからフォルダIDを取得
    const folderId = getFolderIdFromUrl(folderUrl);
    if (!folderId) {
      throw new Error('フォルダIDの取得に失敗しました。');
    }
    
    // フォルダオブジェクトを取得
    const folder = DriveApp.getFolderById(folderId);
    
    // フォルダ内のファイルを名前の昇順で取得
    const files = getSortedFiles(folder);
    
    // ファイルを順次処理
    files.forEach((file, index) => {
      Logger.log(`処理開始: ${index + 1}/${files.length} - ${file.getName()}`);
      
      // 現在選択中の行の添付ファイルセルへURLを書き込む
      const prop = getSystemProperty();
      let ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(prop.selectedSheetName);
      let currentRow = sheet.getActiveCell().getRowIndex();
      let attachmentValues = sheet.getRange(currentRow
                                          , CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF
                                          , 1
                                          , CONFIG.ATTACH_END_COLUMN - CONFIG.ATTACH_START_COLUMN + 1).getValues();
    
      // 添付ファイル領域の列へアップロードしたファイルのURLをセットする
      for(let currentCol = 0 ; currentCol < attachmentValues[0].length ; currentCol += CONFIG.ATTACH_COLUMN_DIF){

        // 添付ファイルが入っていない箇所へセットする
        if(attachmentValues[0][currentCol] == ""){
          sheet.getRange(currentRow, currentCol + CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF).setValue(file.getUrl());

          // Xの添付数を取得する
          let attachmentXCount = sheet.getRange(currentRow, CONFIG.ATTACH_COUNT_X_COL + CONFIG.SHEET_ARRAY_COL_DIF).getValue();

          // Xの添付数が4を超えていない場合はチェックを入れる
          if(attachmentXCount < 4){
            sheet.getRange(currentRow, currentCol + CONFIG.ATTACH_START_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF + CONFIG.ATTACH_X_CHECK_DIF).setValue(true);
          }

          // スプレッドシートをリフレッシュする
          SpreadsheetApp.flush();
          break;
        }
      }
    });
    
    return {
      success: true,
      message: `${files.length}件のファイル処理が完了しました。`
    };
    
  } catch (error) {
    Logger.log(`エラーが発生しました: ${error.message}`);
    return {
      success: false,
      message: `エラーが発生しました: ${error.message}`
    };
  }
}