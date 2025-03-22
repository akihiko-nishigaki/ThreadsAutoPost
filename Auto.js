/***********************************
 * 自動投稿の実行メソッド
 * 今の時間が開始時間と終了時間内に入っていたら処理を行う
 ***********************************/
function autoPostScheduler() {
  try {
    // 起動時間と終了時間を取得する
    const autoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);
    const scheduleStartTime = autoSheet.getRange(CONFIG_AUTO.CELL_SCHEDULE_START_TIME).getValue();
    const scheduleEndTime = autoSheet.getRange(CONFIG_AUTO.CELL_SCHEDULE_END_TIME).getValue();

    // 現在のJST時間を取得
    const jstDate = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Tokyo" }));
    
    // 時間オブジェクトを生成
    const currentTime = createTimeObject(jstDate);
    const startTime = createTimeObject(scheduleStartTime);
    const endTime = createTimeObject(scheduleEndTime);
    
    // デバッグログ出力
    Logger.log(`Current Time: ${getTimeString(jstDate)}`);
    Logger.log(`Schedule Start: ${getTimeString(scheduleStartTime)}`);
    Logger.log(`Schedule End: ${getTimeString(scheduleEndTime)}`);
    
    // 時間範囲内かチェック
    if (isTimeInRange(currentTime, startTime, endTime)) {
      // ここに実行したい処理を記述
      autoPostExecute();
    } else {
      Logger.log('現在時刻は実行時間範囲外です');
    }
    
  } catch (error) {
    Logger.log(`Error in autoPostScheduler: ${error.message}`);
    throw error;
  }
}

/***********************************
 * 自動投稿の実行メソッド
 * トリガーによって定期的に実行される
 ***********************************/
function autoPostExecute() {

  Logger.log('autoPostExecute:Start');

  // シート取得
  const autoSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);

  try {
    // 処理する番号を取得する
    let targetRowNo = autoSheet.getRange(CONFIG_AUTO.CELL_CURRENT_ROW).getValue();
    
    // 上限の番号を取得する
    let limitRowNo = autoSheet.getRange(CONFIG_AUTO.CELL_LIMIT_ROW).getValue();

    // 実際の投稿を行う
    const now = new Date();
    Logger.log(`現在時刻: ${now.toLocaleString()}`);
    
    // データがない場合は処理しない
    if(limitRowNo == 0){
      Logger.log("データがないため処理終了する")
      return;
    }

    // // データ範囲を一括で取得
    // const dataRange = autoSheet.getRange(
    //   CONFIG.START_ROW + targetRowNo - 1,
    //   2,    // C列
    //   1,    // 1行分
    //   32    // B-AG列を取得
    // );
    // データ範囲を一括で取得
    const dataRange = autoSheet.getRange(
      CONFIG_AUTO.START_ROW,
      2,
      limitRowNo,
      32 // B-AG列を取得
    );
    const values = dataRange.getValues();

    // データを処理（1行分のみ）
    let rowData = values[targetRowNo - 1];
 
    const rowIndex = targetRowNo + CONFIG_AUTO.START_ROW - 1;
    
    Logger.log(`行${rowIndex}: 投稿処理を開始`);
    const statusUpdate = executePost(autoSheet, rowData, rowIndex, values, SHEETS_NAME.AUTO);
    
    // 次回投稿番号を更新する
    updateTagetNo(autoSheet, targetRowNo, limitRowNo);
  
    }catch(error){
      Logger.log(`スケジュールチェック中にエラーが発生: ${error.message}`);
      throw error;

  }
  Logger.log('autoPostExecute:End');

}

/***********************************
 * 次回投稿番号の更新
 ***********************************/
function updateTagetNo(autoSheet, targetRowNo, limitRowNo){

  Logger.log('updateTagetNo:Start');
  Logger.log(`targetRowNo:${targetRowNo}`);
  Logger.log(`limitRowNo:${limitRowNo}`);

  // 次回投稿番号をインクリメントする
  targetRowNo++;

  // 次回投稿番号が上限番号を超えていたら1に戻す
  if(targetRowNo > limitRowNo){
    targetRowNo = 1;
  }
  
  // 次回投稿番号を更新する
  autoSheet.getRange(CONFIG_AUTO.CELL_CURRENT_ROW).setValue(targetRowNo);

}


/**                                                                                                                                                 
 * 起動確認ダイアログを表示
 */
function showAutoPostStartDialog() {
  Logger.log('起動ダイアログを表示');
  const ui = SpreadsheetApp.getUi();
  // const response = ui.alert(
  //   '自動投稿スケジューラーの起動',
  //   '自動投稿スケジューラーを起動しますか？',
  //   ui.ButtonSet.YES_NO
  // );
  
  // if (response === ui.Button.YES) {
  Logger.log('ユーザーが起動を承認。自動投稿トリガー設定を開始');
    
  // トリガーをセットする
  setAutoPostTrigger();

  // 予定投稿のステータスを更新する
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);
  sheet.getRange(CONFIG.CELL_TRIGGER_STATUS).setValue(CONFIG.STRING_SHUT);

  ui.alert('自動投稿スケジューラーを起動しました');
  // } else {
  //   Logger.log('ユーザーが起動をキャンセル');
  // }
}

/**
 * 停止確認ダイアログを表示
 */
function showAutoPostStopDialog() {
  Logger.log('自動投稿停止ダイアログを表示');
  const ui = SpreadsheetApp.getUi();
  // const response = ui.alert(
  //   '自動投稿スケジューラーの停止',
  //   '自動投稿スケジューラーを停止しますか？',
  //   ui.ButtonSet.YES_NO
  // );
  
  // if (response === ui.Button.YES) {
  Logger.log('ユーザーが停止を承認。自動投稿トリガー削除を開始');
  deleteTriggers();

  // トリガーを削除する
  deleteAutoPostTrigger();

  // 予定投稿のステータスを更新する
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);
  sheet.getRange(CONFIG.CELL_TRIGGER_STATUS).setValue(CONFIG.STRING_SHUTDOWN);

  ui.alert('自動投稿スケジューラーを停止しました');
  // } else {
  //   Logger.log('ユーザーが停止をキャンセル');
  // }
}


/**
 * トリガーを設定する
 * 既存のトリガーがある場合は一旦削除して新規作成する
 */
function setAutoPostTrigger() {
  try {
    // 現在の時間間隔を取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.AUTO);
    const intervalHours = sheet.getRange(CONFIG_AUTO.CELL_SCHEDULE_EXEC_TIME).getValue();
    
    // 既存のトリガーを削除
    deleteAutoPostTrigger();
    
    // 新しいトリガーを作成
    const trigger = ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTION_NAME_AUTO)
      .timeBased()
      .everyHours(intervalHours)
      .create();
    
    // トリガーIDをプロパティに保存
    PropertiesService.getScriptProperties().setProperty(CONFIG.AUTO_SCHEDULE_TRIGGER_ID, trigger.getUniqueId());
    
    Logger.log(`Trigger set successfully for every ${intervalHours} hours`);
    return {
      success: true,
      message: `Trigger set successfully for every ${intervalHours} hours`,
      triggerId: trigger.getUniqueId()
    };
    
  } catch (error) {
    Logger.log(`Error in setTrigger: ${error.message}`);
    return {
      success: false,
      message: `Failed to set trigger: ${error.message}`
    };
  }
}

/**
 * 既存のトリガーを削除する
 */
function deleteAutoPostTrigger() {
  try {
    // 保存されているトリガーIDを取得
    const triggerId = PropertiesService.getScriptProperties().getProperty(CONFIG.AUTO_SCHEDULE_TRIGGER_ID);
    
    // 全てのトリガーをチェック
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === CONFIG.TRIGGER_FUNCTION_NAME_AUTO) {
        // トリガーIDが一致するか、トリガーIDが保存されていない場合は削除
        if (!triggerId || trigger.getUniqueId() === triggerId) {
          ScriptApp.deleteTrigger(trigger);
          Logger.log('Existing trigger deleted');
        }
      }
    }
    
    // トリガーIDをプロパティから削除
    PropertiesService.getScriptProperties().deleteProperty(CONFIG.AUTO_SCHEDULE_TRIGGER_ID);
    
    return {
      success: true,
      message: 'Trigger deleted successfully'
    };
    
  } catch (error) {
    Logger.log(`Error in deleteAutoPostTrigger: ${error.message}`);
    return {
      success: false,
      message: `Failed to delete trigger: ${error.message}`
    };
  }
}