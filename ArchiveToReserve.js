/**
 * メイン処理を実行する関数
 * 過去データから予約投稿シートへデータを転記する
 */
function transferHistoryToReservation() {
  logMessage('============================================');
  logMessage('過去データから予約投稿シートへの転記処理を開始します');
  logMessage('============================================');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    logMessage('スプレッドシートを取得しました');

    const reservationSheet = ss.getSheetByName(SHEETS_NAME.RESERVATION);
    const historySheet = ss.getSheetByName(SHEETS_NAME.HISTORY);
    
    if (!reservationSheet || !historySheet) {
      throw new Error('シートが見つかりません');
    }

    logMessage(`予約投稿シート: ${SHEETS_NAME.RESERVATION}, 過去データシート: ${SHEETS_NAME.HISTORY} を取得しました`);
    logMessage(`予約投稿シートの開始行: ${CONFIG.START_ROW}, 過去データシートの開始行: ${CONFIG_HISTORY.SHEET_ROW_START}`);
    
    const dataHandler = new HistoryDataHandler(reservationSheet, historySheet);
    dataHandler.processData();
    
    // チェックボックスのクリア処理を最終行情報付きで呼び出し
    clearHistoryCheckboxes(historySheet, dataHandler.lastDataRow);
    
    logMessage('============================================');
    logMessage('データ転記処理が正常に完了しました');
    logMessage('============================================');
  } catch (error) {
    logMessage('============================================');
    logMessage(`エラーが発生しました: ${error.message}`, 'ERROR');
    logMessage(`エラーの詳細: ${error.stack}`, 'ERROR');
    logMessage('============================================');
    throw error;
  }
}

/**
 * 過去データ処理を行うクラス
 */
class HistoryDataHandler {
  /**
   * コンストラクタ
   */
  constructor(reservationSheet, historySheet) {
    this.reservationSheet = reservationSheet;
    this.historySheet = historySheet;
    this.columnRanges = {
      firstRange: {
        start: CONFIG.DATE_COL + CONFIG.SHEET_ARRAY_COL_DIF,  // C列
        end: CONFIG.POST_TEXT + CONFIG.SHEET_ARRAY_COL_DIF    // F列
      },
      secondRange: {
        start: CONFIG.POST_RES + CONFIG.SHEET_ARRAY_COL_DIF,         // H列
        end: CONFIG.POST_CHECK_INSTA + CONFIG.SHEET_ARRAY_COL_DIF    // M列
      },
      thirdRange: {
        start: CONFIG.MEMO + CONFIG.SHEET_ARRAY_COL_DIF,              // X列
        end: CONFIG.ATTACH_END_COLUMN + CONFIG.SHEET_ARRAY_COL_DIF    // BL列
      }
    };
    // 初期化時にC列の最終行を取得
    this.lastDataRow = this.getHistoryLastRow();
    
    logMessage('HistoryDataHandlerを初期化しました');
    logMessage(`転記範囲1: C列(${this.columnRanges.firstRange.start})～F列(${this.columnRanges.firstRange.end})`);
    logMessage(`転記範囲2: H列(${this.columnRanges.secondRange.start})～M列(${this.columnRanges.secondRange.end})`);
    logMessage(`転記範囲3: X列(${this.columnRanges.thirdRange.start})～BL列(${this.columnRanges.thirdRange.end})`);
    logMessage(`過去データの最終行: ${this.lastDataRow}行目`);
  }

  /**
   * 過去データシートのC列の最終行を取得
   * getNextDataCell()を使用して効率的に最終行を特定
   */
  getHistoryLastRow() {
    logMessage('過去データシートのC列最終行の検索を開始します');
    
    try {
      // C列の最終データ行を取得
      const lastRow = this.historySheet.getRange(this.historySheet.getMaxRows(), CONFIG.DATE_COL + CONFIG.SHEET_ARRAY_COL_DIF)
        .getNextDataCell(SpreadsheetApp.Direction.UP)
        .getRow();
      
      // 最終行が開始行（5行目）より前の場合は、開始行を返す
      const finalRow = Math.max(lastRow, CONFIG_HISTORY.SHEET_ROW_START);
      
      logMessage(`過去データシートのC列最終行を特定しました: ${finalRow}行目`);
      return finalRow;

    } catch (e) {
      logMessage(`最終行の検索中にエラーが発生しました: ${e.message}`, 'ERROR');
      logMessage(`エラーの詳細: ${e.stack}`, 'ERROR');
      return 5; // エラーの場合は開始行を返す
    }
  }


  /**
   * メイン処理を実行
   */
  processData() {
    logMessage('----------------------------------------');
    logMessage('データ処理を開始します');

    // 転記対象データを取得
    const transferableData = this.getTransferableData();
    if (!transferableData || transferableData.length === 0) {
      logMessage('転記可能なデータが存在しません', 'WARNING');
      return;
    }

    logMessage(`${transferableData.length}行の転記可能データを特定しました`);
    logMessage('----------------------------------------');

    // データを転記
    this.transferData(transferableData);
  }

/**
 * 転記対象データを取得
 */
getTransferableData() {
  logMessage('転記対象データの取得処理を開始します');
  
  try {
    // 5行目から最終行までの行数を計算
    const dataLength = this.lastDataRow - 4;
    
    logMessage(`データ取得範囲: 5行目から${this.lastDataRow}行目まで（${dataLength}行）`);
    
    logMessage('チェックボックス列（B列）の取得を開始');
    const checkboxRange = this.historySheet.getRange(CONFIG_HISTORY.SHEET_ROW_START, CONFIG_HISTORY.SHEET_COL_COPY_CHECK, dataLength, 1);
    const checkboxValues = checkboxRange.getValues();
    logMessage(`チェックボックス列を取得しました（${dataLength}行）`);
    
    logMessage('最初の転記範囲（C～F列）のデータ取得を開始');
    const firstRangeData = this.historySheet.getRange(
      5,
      this.columnRanges.firstRange.start,
      dataLength,
      this.columnRanges.firstRange.end - this.columnRanges.firstRange.start
    ).getValues();
    logMessage('最初の転記範囲のデータを取得しました');

    logMessage('二番目の転記範囲（H～M列）のデータ取得を開始');
    const secondRangeData = this.historySheet.getRange(
      5,
      this.columnRanges.secondRange.start,
      dataLength,
      this.columnRanges.secondRange.end - this.columnRanges.secondRange.start
    ).getValues();
    logMessage('二番目の転記範囲のデータを取得しました');

    logMessage('三番目の転記範囲（W～BL列）のデータ取得を開始');
    const thirdRangeData = this.historySheet.getRange(
      5,
      this.columnRanges.thirdRange.start,
      dataLength,
      this.columnRanges.thirdRange.end - this.columnRanges.thirdRange.start
    ).getValues();
    logMessage('三番目の転記範囲のデータを取得しました');

    // チェックされているデータを抽出
    const transferableData = [];
    let checkedCount = 0;
    let validDataCount = 0;

    checkboxValues.forEach((checkbox, index) => {
      if (checkbox[0] === true) {
        checkedCount++;
        if (firstRangeData[index][0]) { // C列のデータが存在する行のみ
          validDataCount++;
          transferableData.push({
            rowIndex: index + CONFIG_HISTORY.SHEET_ROW_START, // 実際の行番号（デバッグ用）
            firstRange: firstRangeData[index],
            secondRange: secondRangeData[index],
            thirdRange: thirdRangeData[index]
          });
        }
      }
    });

    logMessage(`チェックされている行数: ${checkedCount}`);
    logMessage(`C列にデータがある行数: ${validDataCount}`);
    logMessage(`最終的な転記対象データ数: ${transferableData.length}`);
    
    if (transferableData.length > 0) {
      logMessage('転記対象の行番号:');
      transferableData.forEach(item => {
        logMessage(`- ${item.rowIndex}行目`);
      });
    }

    return transferableData;
  } catch (e) {
    logMessage(`データ取得中にエラーが発生しました: ${e.message}`, 'ERROR');
    logMessage(`エラーの詳細: ${e.stack}`, 'ERROR');
    return null;
  }
}
  /**
   * 予約投稿シートの最終データ行を取得
   */
  getLastReservationRow() {
    logMessage('予約投稿シートの最終行の検索を開始します');
    
    const cColumn = this.reservationSheet.getRange("C:C").getValues();
    let lastRow = CONFIG.START_ROW - 1;
    
    for (let i = CONFIG.START_ROW - 1; i < CONFIG.END_ROW; i++) {
      if (cColumn[i][0]) {
        lastRow = i + 1;
      }
    }
    
    logMessage(`予約投稿シートの最終行を特定しました: ${lastRow}行目`);
    logMessage(`次のデータは${lastRow + 1}行目から転記されます`);
    return lastRow;
  }

  /**
   * データを転記
   */
  transferData(transferableData) {
    logMessage('----------------------------------------');
    logMessage('データの転記処理を開始します');

    try {
      const lastRow = this.getLastReservationRow();
      
      if (transferableData.length > 0) {
        logMessage('最初の範囲（C～F列）の転記を開始');
        const firstRangeTarget = this.reservationSheet.getRange(
          lastRow + 1,
          this.columnRanges.firstRange.start,
          transferableData.length,
          this.columnRanges.firstRange.end - this.columnRanges.firstRange.start
        );
        const firstRangeValues = transferableData.map(item => item.firstRange);
        firstRangeTarget.setValues(firstRangeValues);
        logMessage('最初の範囲の転記が完了しました');

        logMessage('二番目の範囲（H～M列）の転記を開始');
        const secondRangeTarget = this.reservationSheet.getRange(
          lastRow + 1,
          this.columnRanges.secondRange.start,
          transferableData.length,
          this.columnRanges.secondRange.end - this.columnRanges.secondRange.start
        );
        const secondRangeValues = transferableData.map(item => item.secondRange);
        secondRangeTarget.setValues(secondRangeValues);
        logMessage('二番目の範囲の転記が完了しました');

        logMessage('三番目の範囲（W～BL列）の転記を開始');
        const thirdRangeTarget = this.reservationSheet.getRange(
          lastRow + 1,
          this.columnRanges.thirdRange.start,
          transferableData.length,
          this.columnRanges.thirdRange.end - this.columnRanges.thirdRange.start
        );
        const thirdRangeValues = transferableData.map(item => item.thirdRange);
        thirdRangeTarget.setValues(thirdRangeValues);
        logMessage('三番目の範囲の転記が完了しました');

        logMessage(`合計${transferableData.length}行のデータを${lastRow + 1}行目から${lastRow + transferableData.length}行目に転記しました`);
      }
    } catch (e) {
      logMessage(`データ転記中にエラーが発生しました: ${e.message}`, 'ERROR');
      logMessage(`エラーの詳細: ${e.stack}`, 'ERROR');
      throw e;
    }
  }
}


/**
 * 過去データシートのチェックボックスをクリアする関数
 * @param {GoogleAppsScript.Spreadsheet.Sheet} historySheet - 過去データシート
 * @param {number} lastDataRow - データの最終行
 */
function clearHistoryCheckboxes(historySheet, lastDataRow) {
  logMessage('============================================');
  logMessage('チェックボックスのクリア処理を開始します');
  logMessage('============================================');

  try {
    if (!historySheet) {
      throw new Error('過去データシートが見つかりません');
    }

    logMessage(`過去データシート: ${SHEETS_NAME.HISTORY} を取得しました`);

    // チェックボックスの範囲を取得
    const startRow = CONFIG.ARCHIVE_START_ROW + 1; // ヘッダー行の次から開始
    const numRows = lastDataRow - startRow + 1; // 最終データ行までの行数
    
    logMessage(`クリア対象範囲: ${startRow}行目から${lastDataRow}行目まで（${numRows}行）`);

    // チェックボックスをクリア
    const checkboxRange = historySheet.getRange(startRow, CONFIG.ARCHIVE_COPY_CHECK_BOX_COL, numRows, 1); // B列
    checkboxRange.setValue(false);

    logMessage(`${numRows}行分のチェックボックスをクリアしました`);
    logMessage('============================================');
    logMessage('チェックボックスのクリア処理が完了しました');
    logMessage('============================================');

  } catch (error) {
    logMessage('============================================');
    logMessage(`エラーが発生しました: ${error.message}`, 'ERROR');
    logMessage(`エラーの詳細: ${error.stack}`, 'ERROR');
    logMessage('============================================');
    throw error;
  }
}