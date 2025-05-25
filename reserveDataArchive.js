/**
 * メイン処理を実行する関数
 * 予約投稿データを過去データシートに転記する
 */
function transferReservationData() {

  // 予約投稿データのデータを過去シートへ転記する
  doTransferReservationData(SHEETS_NAME.RESERVATION);
}
/**
 * メイン処理を実行する関数
 * 予約投稿データを過去データシートに転記する
 */
function doTransferReservationData(sheetName) {
  logMessage('データ転記処理を開始します');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let targetSheetName;
    let historySheetName;

    // 自動投稿 or 予約投稿を区別してシート名を取得する
    if(sheetName == SHEETS_NAME.RESERVATION){
      targetSheetName = SHEETS_NAME.RESERVATION;
      historySheetName = SHEETS_NAME.HISTORY;
    }else if(sheetName == SHEETS_NAME.AUTO){
      targetSheetName = SHEETS_NAME.AUTO;
      historySheetName = SHEETS_NAME.AUTO_HISTORY;
    }
    
    // ターゲットシートと過去データシートを宣言する
    let targetSheet = ss.getSheetByName(targetSheetName);
    let historySheet = ss.getSheetByName(historySheetName);
    
    if (!targetSheet || !historySheet) {
      throw new Error('シートが見つかりません');
    }

    logMessage(`対象シート: ${targetSheetName}, 過去データシート: ${historySheet} を取得しました`);
    
    // データ転記処理の実行
    const dataHandler = new DataHandler(targetSheet, historySheet, sheetName);
    dataHandler.processData(sheetName, historySheet, targetSheet);
    
    logMessage('データ転記処理が完了しました');
  } catch (error) {
    logMessage(`エラーが発生しました: ${error.message}`, 'ERROR');
    throw error;
  }
}

/**
 * 予約投稿シートのチェックボックスで選択されたデータを削除
 */
function deleteTargetData(){
  // 予約投稿シートのデータを削除する
  doDeleteTargetData(SHEETS_NAME.RESERVATION);
}
/**
 * 自動投稿シートのチェックボックスで選択されたデータを削除
 */
function deleteTargetDataAutoSheet(){
  // 自動投稿シートのデータを削除する
  doDeleteTargetData(SHEETS_NAME.AUTO);
}

/**
 * チェックボックスで選択されたデータを削除する関数
 */
function doDeleteTargetData(sheetName) {
  logMessage('データ削除処理を開始します');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    logMessage(`予約投稿シート: ${sheetName} を取得しました`);

    let checkBoxCol;    
    // シート名によってチェックボックスの位置が違うため、取得する
    if(sheetName == SHEETS_NAME.RESERVATION){
      // 予約投稿
      checkBoxCol = CONFIG.SHEET_COL_DELETE_CHECKBOX;
    }else{
      // 自動投稿
      checkBoxCol = CONFIG_AUTO.SHEET_COL_DELETE_CHECKBOX;
    }

    // データ削除処理の実行
    const dataHandler = new DataHandler(sheet, null, sheetName);
    const deletedCount = processDeleteData(dataHandler, sheet, checkBoxCol, sheetName);
    
    logMessage(`${deletedCount}件のデータを削除しました`);
  } catch (error) {
    logMessage(`エラーが発生しました: ${error.message}`, 'ERROR');
    throw error;
  }
}

/**
 * データ削除処理のメイン関数
 * @param {DataHandler} dataHandler - データハンドラーインスタンス
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - 対象シート
 * @param {number} checkBoxCol - チェックボックスがあるシート列番号
 * @param {string} sheetName - 対象シート名
 * @returns {number} 削除したデータ数
 */
function processDeleteData(dataHandler, sheet, checkBoxCol, sheetName) {
  logMessage('削除対象データの検索を開始します');
  
  try {
    // 最終データ行を取得
    let lastRow;

    if(this.sheetName == SHEETS_NAME.RESERVATION){
      lastRow = dataHandler.getLastDataRow();
    }else{
      lastRow = dataHandler.getLastDataRowAutoSheet();
    }

    
    // チェックボックスの値を取得
    const checkboxRange = sheet.getRange(CONFIG.START_ROW, checkBoxCol, lastRow - CONFIG.START_ROW + 1, 1);
    const checkboxValues = checkboxRange.getValues();
    
    // 削除対象の行を特定
    const deleteTargets = [];
    checkboxValues.forEach((value, index) => {
      if (value[0] === true) {
        deleteTargets.push({
          rowIndex: index + CONFIG.START_ROW
        });
      }
    });
    
    if (deleteTargets.length === 0) {
      logMessage('削除対象のデータが存在しません');
      return 0;
    }

    logMessage(`${deleteTargets.length}件の削除対象データを特定しました`);
    
    // 削除処理（既存のcleanupSourceDataを利用）
    dataHandler.cleanupSourceData(deleteTargets, sheet, sheetName);
    
    return deleteTargets.length;
  } catch (e) {
    logMessage(`データ削除処理中にエラーが発生しました: ${e.message}`, 'ERROR');
    throw e;
  }
}

/**
 * J列のチェックボックスをすべてオンにする
 */
function checkAllXColumn() {
  setCheckboxStatus(SHEETS_NAME.RESERVATION, CONFIG.POST_CHECK_X + CONFIG.SHEET_ARRAY_COL_DIF, true);
}

/**
 * J列のチェックボックスをすべてオフにする
 */
function uncheckAllXColumn() {
  setCheckboxStatus(SHEETS_NAME.RESERVATION,CONFIG.POST_CHECK_X + CONFIG.SHEET_ARRAY_COL_DIF, false);
}

/**
 * K列のチェックボックスをすべてオンにする
 */
function checkAllThreadsColumn() {
  setCheckboxStatus(SHEETS_NAME.RESERVATION,CONFIG.POST_CHECK_THREADS + CONFIG.SHEET_ARRAY_COL_DIF, true);
}

/**
 * K列のチェックボックスをすべてオフにする
 */
function uncheckAllThreadsColumn() {
  setCheckboxStatus(SHEETS_NAME.RESERVATION,CONFIG.POST_CHECK_THREADS + CONFIG.SHEET_ARRAY_COL_DIF, false);
}

/**
 * L列のチェックボックスをすべてオンにする
 */
function checkAllInstaColumn() {
  setCheckboxStatus(SHEETS_NAME.RESERVATION,CONFIG.POST_CHECK_INSTA + CONFIG.SHEET_ARRAY_COL_DIF, true);
}

/**
 * L列のチェックボックスをすべてオフにする
 */
function uncheckAllInstaColumn() {
  setCheckboxStatus(SHEETS_NAME.RESERVATION, CONFIG.POST_CHECK_INSTA + CONFIG.SHEET_ARRAY_COL_DIF, false);
}

/**
 * J列のチェックボックスをすべてオンにする
 */
function checkAllXColumnAutoSheet() {
  setCheckboxStatus(SHEETS_NAME.AUTO, CONFIG.POST_CHECK_X + CONFIG.SHEET_ARRAY_COL_DIF, true);
}

/**
 * J列のチェックボックスをすべてオフにする
 */
function uncheckAllXColumnAutoSheet() {
  setCheckboxStatus(SHEETS_NAME.AUTO,CONFIG.POST_CHECK_X + CONFIG.SHEET_ARRAY_COL_DIF, false);
}

/**
 * K列のチェックボックスをすべてオンにする
 */
function checkAllThreadsColumnAutoSheet() {
  setCheckboxStatus(SHEETS_NAME.AUTO,CONFIG.POST_CHECK_THREADS + CONFIG.SHEET_ARRAY_COL_DIF, true);
}

/**
 * K列のチェックボックスをすべてオフにする
 */
function uncheckAllThreadsColumnAutoSheet() {
  setCheckboxStatus(SHEETS_NAME.AUTO,CONFIG.POST_CHECK_THREADS + CONFIG.SHEET_ARRAY_COL_DIF, false);
}

/**
 * L列のチェックボックスをすべてオンにする
 */
function checkAllInstaColumnAutoSheet() {
  setCheckboxStatus(SHEETS_NAME.AUTO,CONFIG.POST_CHECK_INSTA + CONFIG.SHEET_ARRAY_COL_DIF, true);
}

/**
 * L列のチェックボックスをすべてオフにする
 */
function uncheckAllInstaColumnAutoSheet() {
  setCheckboxStatus(SHEETS_NAME.AUTO, CONFIG.POST_CHECK_INSTA + CONFIG.SHEET_ARRAY_COL_DIF, false);
}

/**
 * 指定された列のチェックボックスの状態を一括で変更する
 * パフォーマンスを考慮して、バッチ処理で実行する
 * 
 * @param {number} columnNumber - 対象の列番号
 * @param {boolean} status - 設定する状態（true: チェックオン、false: チェックオフ）
 */
function setCheckboxStatus(sheetName, columnNumber, status) {
  logMessage(`チェックボックス一括制御処理を開始します（列: ${columnNumber}, 状態: ${status}）`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // 総処理行数を計算
    let totalRows;
    
    if (sheetName == SHEETS_NAME.RESERVATION){
      totalRows = CONFIG.END_ROW - CONFIG.START_ROW + 1;
    }else if(sheetName == SHEETS_NAME.AUTO){
      totalRows = CONFIG_AUTO.END_ROW - CONFIG_AUTO.START_ROW + 1;
    }

    // バッチ処理の回数を計算
    const batchCount = Math.ceil(totalRows / CONFIG.ROWS_PER_BATCH);
    
    logMessage(`処理対象行数: ${totalRows}, バッチ数: ${batchCount}`);

    // バッチ処理でチェックボックスの状態を更新
    for (let i = 0; i < batchCount; i++) {
      const startRow = CONFIG.START_ROW + (i * CONFIG.ROWS_PER_BATCH);
      const rowCount = Math.min(CONFIG.ROWS_PER_BATCH, 
                               CONFIG.END_ROW - startRow + 1);
      
      // 更新用の2次元配列を作成（すべての要素が同じ値）
      const values = Array(rowCount).fill([status]);
      
      // 範囲を一括で更新
      sheet.getRange(startRow, columnNumber, rowCount, 1).setValues(values);
      
      logMessage(`バッチ処理完了: ${startRow}行目から${rowCount}行を処理`);
    }

    logMessage('チェックボックス一括制御処理が完了しました');
  } catch (error) {
    logMessage(`エラーが発生しました: ${error.message}`, 'ERROR');
    throw error;
  }
}


/**
 * データ処理を行うクラス
 */
class DataHandler {
  /**
   * コンストラクタ
   * @param {GoogleAppsScript.Spreadsheet.Sheet} reservationSheet - 予約投稿シート
   * @param {GoogleAppsScript.Spreadsheet.Sheet} historySheet - 過去データシート
   */
  constructor(reservationSheet, historySheet, sheetName) {
    this.reservationSheet = reservationSheet;
    this.historySheet = historySheet;
    this.sheetName = sheetName;
    this.startCol = CONFIG.DATE_COL + CONFIG.SHEET_ARRAY_COL_DIF;
    this.endCol = CONFIG.ATTACH_END_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF;
    this.postDateCol = CONFIG.POST_STATUS_COL + CONFIG.SHEET_ARRAY_COL_DIF;
    logMessage('DataHandlerを初期化しました');
  }

  /**
   * メイン処理を実行
   */
  processData(sheetName, sheet, targetSheet) {
    logMessage('データ処理を開始します');

    // データ範囲を取得
    const dataRange = this.getDataRange();
    if (!dataRange || dataRange.length === 0) {
      logMessage('処理対象のデータが存在しません');
      return;
    }

    logMessage(`${dataRange.length}行のデータを取得しました`);

    // 転記可能なデータを抽出
    const transferableData = this.getTransferableData(dataRange);
    if (transferableData.length === 0) {
      logMessage('転記可能なデータが存在しません');
      return;
    }

    logMessage(`${transferableData.length}行の転記可能データを特定しました`);

    // データを転記
    this.transferData(transferableData);

    // 元データを削除し、上詰め処理を実行
    this.cleanupSourceData(transferableData, targetSheet, sheetName);
  }

  /**
   * データ範囲を取得
   * @returns {Array<Array>|null} データ範囲の値
   */
  getDataRange() {
    logMessage('データ範囲の取得を開始します');

    let lastRow;

    if(this.sheetName == SHEETS_NAME.RESERVATION){
      lastRow = this.getLastDataRow();
    }else{
      lastRow = this.getLastDataRowAutoSheet();
    }

    logMessage(`最終データ行: ${lastRow}`);
    
    if (lastRow < CONFIG.START_ROW) {
      logMessage('有効なデータ範囲が存在しません');
      return null;
    }

    // 範囲の計算
    const numRows = lastRow - CONFIG.START_ROW + 1;
    const numCols = CONFIG.ATTACH_END_COLUMN - CONFIG.DATE_COL + 1;

    logMessage(`取得する範囲: 行数=${numRows}, 列数=${numCols}`);

    try {
      const values = this.reservationSheet.getRange(
        CONFIG.START_ROW,
        CONFIG.DATE_COL + CONFIG.SHEET_ARRAY_COL_DIF,
        numRows,
        numCols
      ).getValues();
      
      logMessage('データ範囲の取得に成功しました');
      return values;
    } catch (e) {
      logMessage(`データ範囲の取得に失敗しました: ${e.message}`, 'ERROR');
      return null;
    }
  }

  /**
   * 最終データ行を取得
   * @returns {number} 最終データ行
   */
  getLastDataRow() {
    logMessage('最終データ行の検索を開始します');
    
    try {
      // C列の最終データ行を取得
      const lastRow = this.reservationSheet.getRange(this.reservationSheet.getMaxRows(), 3) // 3はC列を示す
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      
      // CONFIG.START_ROWより前の場合は開始行を返す
      const finalRow = Math.max(lastRow, CONFIG.START_ROW);
      
      logMessage(`最終データ行を特定しました: ${finalRow}`);
      return finalRow;

    } catch (e) {
      logMessage(`最終行の検索中にエラーが発生しました: ${e.message}`, 'ERROR');
      logMessage(`エラーの詳細: ${e.stack}`, 'ERROR');
      return CONFIG.START_ROW; // エラーの場合は開始行を返す
    }
  }

    /**
   * 最終データ行を取得
   * @returns {number} 最終データ行
   */
  getLastDataRowAutoSheet() {
    logMessage('最終データ行の検索を開始します');
    
    try {
      // F列の最終データ行を取得
      const lastRow = this.reservationSheet.getRange(this.reservationSheet.getMaxRows(), 6) // 6はF列を示す
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      
      // CONFIG.START_ROWより前の場合は開始行を返す
      const finalRow = Math.max(lastRow, CONFIG.START_ROW);
      
      logMessage(`最終データ行を特定しました: ${finalRow}`);
      return finalRow;

    } catch (e) {
      logMessage(`最終行の検索中にエラーが発生しました: ${e.message}`, 'ERROR');
      logMessage(`エラーの詳細: ${e.stack}`, 'ERROR');
      return CONFIG.START_ROW; // エラーの場合は開始行を返す
    }
  }

  /**
   * 転記可能なデータを抽出
   * @param {Array<Array>} dataRange - データ範囲
   * @returns {Array<Object>} 転記可能なデータ
   */
  getTransferableData(dataRange) {
    if (!dataRange) {
      logMessage('データ範囲が無効です', 'WARNING');
      return [];
    }
    
    logMessage('転記可能なデータの抽出を開始します');
    
    const transferableData = [];
    const postDateColIndex = CONFIG.POST_STATUS_COL - CONFIG.DATE_COL;

    dataRange.forEach((row, index) => {
      if (row[postDateColIndex]) {
        transferableData.push({
          rowIndex: index + CONFIG.START_ROW,
          data: row
        });
      }
    });

    logMessage(`${transferableData.length}件の転記可能なデータを抽出しました`);
    return transferableData;
  }

  /**
   * データを転記
   * @param {Array<Object>} transferableData - 転記するデータ
   */
  transferData(transferableData) {
    if (transferableData.length === 0) {
      logMessage('転記するデータがありません', 'WARNING');
      return;
    }

    logMessage('データの転記を開始します');
    const historyLastRow = this.getHistoryLastRow();
    logMessage(`過去データの最終行: ${historyLastRow}`);

    try {
      const targetRange = this.historySheet.getRange(
        historyLastRow + 1,
        CONFIG.DATE_COL + CONFIG.SHEET_ARRAY_COL_DIF,
        transferableData.length,
        CONFIG.ATTACH_END_COLUMN - CONFIG.DATE_COL + 1
      );

      const dataToTransfer = transferableData.map(item => item.data);
      targetRange.setValues(dataToTransfer);
      
      logMessage(`${transferableData.length}行のデータを転記しました`);
    } catch (e) {
      logMessage(`データ転記中にエラーが発生しました: ${e.message}`, 'ERROR');
      throw e;
    }
  }

  /**
   * 履歴シートの最終行を取得
   * @returns {number} 最終行番号
   */
  getHistoryLastRow() {
    logMessage('履歴シートの最終行の検索を開始します');
    
    try {
      // C列の最終データ行を取得
      const lastRow = this.historySheet.getRange(this.historySheet.getMaxRows(), 3) // 3はC列を示す
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      
      // 開始行より前の場合は開始行を返す
      const finalRow = Math.max(lastRow, CONFIG.ARCHIVE_START_ROW - 1);
      
      logMessage(`履歴シートの最終行を特定しました: ${finalRow}`);
      return finalRow;
    } catch (e) {
      logMessage(`最終行の検索中にエラーが発生しました: ${e.message}`, 'ERROR');
      logMessage(`エラーの詳細: ${e.stack}`, 'ERROR');
      return CONFIG.ARCHIVE_START_ROW - 1; // エラーの場合は開始行を返す
    }
  }


  /**
   * テンプレートシートの内容（値と数式）をターゲットシートにコピーする
   * メイン処理を実行する関数
   * @throws {Error} テンプレートシートまたはターゲットシートが見つからない場合
   */
  cleanupSourceData(transferableData, sheet, sheetName) {
    
    logMessage('元データのクリーンアップを開始します');
    
    try {
      // テンプレートシートを取得
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const templateSheet = ss.getSheetByName(SHEETS_NAME.TEMPLATE);
      const templateAutoSheet = ss.getSheetByName(SHEETS_NAME.TEMPLATE_AUTO);
      if (!templateSheet) {
        throw new Error('テンプレートシートが見つかりません');
      }
      if (!templateAutoSheet) {
        throw new Error('テンプレートシートが見つかりません');
      }
      

      // コピー先のシートを取得
      const targetSheet = sheet;
      if (!targetSheet) {
        throw new Error('コピー先シートが見つかりません');
      }

      // コピー元範囲を取得
      let sourceRange;
      if(sheetName == SHEETS_NAME.RESERVATION){
        sourceRange= templateSheet.getRange(CONFIG.START_ROW, 1, 1, CONFIG.RESERVE_SHEET_COPY_COL_COUNT + CONFIG.SHEET_ARRAY_COL_DIF);
      }else{
        sourceRange= templateAutoSheet.getRange(CONFIG.START_ROW, 1, 1, CONFIG.RESERVE_SHEET_COPY_COL_COUNT + CONFIG.SHEET_ARRAY_COL_DIF);
      }
      // 転記したデータをテンプレートで上書き
      transferableData.forEach(item => {

        let targetRange = targetSheet.getRange(item.rowIndex, 1, 1, CONFIG.RESERVE_SHEET_COPY_COL_COUNT + CONFIG.SHEET_ARRAY_COL_DIF);

        // コピー処理を実行
        sourceRange.copyTo(targetRange,SpreadsheetApp.CopyPasteType.PASTE_NORMAL,false);

      });

      // 上詰め処理
      this.compressData();

      Logger.log('行のコピーが完了しました');
    } catch (error) {
      Logger.log('エラーが発生しました: ' + error.message);
      throw error;
    }
  }

  /**
   * データの上詰め処理
   * 行単位でのコピーペーストとテンプレート適用
   */
  compressData() {
    logMessage('データの上詰め処理を開始します');
    
    try {
      
      let lastRow;
      let targetCol;
      if(this.sheetName == SHEETS_NAME.RESERVATION){
        lastRow = this.getLastDataRow();
        targetCol = CONFIG.DATE_COL + CONFIG.SHEET_ARRAY_COL_DIF;  // C列
      }else{
        lastRow = this.getLastDataRowAutoSheet();
        targetCol = CONFIG_AUTO.SHEET_COL_TEXT;  // F列
      }
      
      // C列の値に基づいて有効な行を特定
      const cColRange = this.reservationSheet.getRange(
        CONFIG.START_ROW,
        targetCol,
        lastRow - CONFIG.START_ROW + 1,
        1
      );
      const cColValues = cColRange.getValues();
      
      // 有効な行のインデックスを収集
      const validRows = cColValues.map((value, index) => ({
        isValid: value[0] !== '',
        sourceRow: CONFIG.START_ROW + index,
        targetRow: null
      })).filter(row => row.isValid);

      if (validRows.length === 0) {
        logMessage('有効なデータが見つかりませんでした');
        return;
      }

      // 移動先の行番号を設定
      validRows.forEach((row, index) => {
        row.targetRow = CONFIG.START_ROW + index;
      });

      logMessage(`有効なデータ行数: ${validRows.length}`);

      // 移動が必要な行のみを処理
      validRows.forEach(row => {
        if (row.sourceRow !== row.targetRow) {
          logMessage(`${row.sourceRow}行目を${row.targetRow}行目に移動します`);
          
          // C列からBL列までの範囲を指定
          const sourceRange = this.reservationSheet.getRange(row.sourceRow, CONFIG.DATE_COL+CONFIG.SHEET_ARRAY_COL_DIF, 1, CONFIG.RESERVE_SHEET_COPY_COL_COUNT); // C列からBL列
          const targetRange = this.reservationSheet.getRange(row.targetRow, CONFIG.DATE_COL+CONFIG.SHEET_ARRAY_COL_DIF, 1, CONFIG.RESERVE_SHEET_COPY_COL_COUNT); // C列からBL列
          
          // 数式を含めてコピー
          sourceRange.copyTo(targetRange);
          
          // 元の行をクリア（移動した行が後続の行の場合のみ）
          if (row.sourceRow > row.targetRow) {
            // テンプレートをコピー
            this.copyTemplateToRow(row.sourceRow, this.sheetName);
          }
        }
      });

      // 未使用行にテンプレートを適用
      if (validRows.length > 0) {
        const lastValidRow = CONFIG.START_ROW + validRows.length - 1;
        if (lastValidRow < lastRow) {
          logMessage(`${lastValidRow + 1}行目から${lastRow}行目までにテンプレートを適用します`);
          
          for (let row = lastValidRow + 1; row <= lastRow; row++) {
            this.copyTemplateToRow(row, this.sheetName);
          }
        }
      }

      logMessage('データの上詰め処理が完了しました');
    } catch (e) {
      logMessage(`上詰め処理中にエラーが発生しました: ${e.message}`, 'ERROR');
      throw e;
    }
  }

  /**
   * 指定行にテンプレートをコピー
   * @param {number} targetRow - コピー先の行番号
   */
  copyTemplateToRow(targetRow, sheetName) {
    try {
      // テンプレートシートを取得
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let templateSheet;

      if(sheetName == SHEETS_NAME.RESERVATION){
        templateSheet = ss.getSheetByName(SHEETS_NAME.TEMPLATE);
      }else if(sheetName == SHEETS_NAME.AUTO){
        templateSheet = ss.getSheetByName(SHEETS_NAME.TEMPLATE_AUTO);
      }

      if (!templateSheet) {
        throw new Error('テンプレートシートが見つかりません');
      }

      // テンプレートの7行目のB列からBK列を取得
      
      const templateRange = templateSheet.getRange(7, 2, 1, CONFIG.ATTACH_END_COLUMN); // B列からBK列
      const targetRange = this.reservationSheet.getRange(targetRow, 2, 1, CONFIG.ATTACH_END_COLUMN);
      
      // テンプレートをコピー
      templateRange.copyTo(targetRange);
      logMessage(`${targetRow}行目にテンプレートを適用しました`);
    } catch (e) {
      logMessage(`テンプレートのコピー中にエラーが発生しました: ${e.message}`, 'ERROR');
      throw e;
    }
  }

}