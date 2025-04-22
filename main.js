/**
 * スプレッドシートを開いた時に実行される関数
 * メニューを追加する（階層構造に変更）
 */
function onOpen() {
  Logger.log('メニューの初期化を開始');
  const ui = SpreadsheetApp.getUi();

  // 共通メニュー
  const commonMenu = ui.createMenu('共通')
    .addItem('アップロードフォルダを作成', 'createUploadFolders')
    .addItem('スクリプトID取得', 'getScriptIDForNewTwitterBotWithImage');

  // Threads投稿メニュー
  const threadsMenu = ui.createMenu('Threads設定')
    .addItem('初期設定', 'showThreadsAuthDialog')
    .addItem('定期更新設定', 'setThreadsTokenRefreshTrigger');

  // X（Twitter）メニュー
  const xMenu = ui.createMenu('X設定')
    .addItem('アカウント認証', 'authorizeLinkForNewTwitterBotWithImage')
    .addItem('ユーザーID取得', 'setXUserId');

  // Instagramメニュー
  const instaMenu = ui.createMenu('Instagram設定')
    .addItem('長期トークン取得', 'getLongTermAccessToken')
    .addItem('定期更新設定', 'setThreadsTokenRefreshTrigger');

  // 解除メニュー
  const clearMenu = ui.createMenu('認証解除')
    .addItem('Threads認証解除', 'clearStoredTokens')
    .addItem('X認証解除', 'clearServiceForNewTwitterBotWithImage')
    .addItem('Instagram認証解除', 'clearServiceForInstagram');

  // メインメニューとして一つにまとめる
  ui.createMenu('設定メニュー')
    .addSubMenu(commonMenu)
    .addSubMenu(threadsMenu)
    .addSubMenu(xMenu)
    .addSubMenu(instaMenu)
    .addSubMenu(clearMenu)
    .addToUi();

  Logger.log('メニューの初期化が完了しました');
}


/**                                                                                                                                                 
 * 起動確認ダイアログを表示
 */
function showStartDialog() {
  Logger.log('起動ダイアログを表示');
  const ui = SpreadsheetApp.getUi();
  // const response = ui.alert(
  //   '予約投稿スケジューラーの起動',
  //   '予約投稿スケジューラーを起動しますか？\n(1分間隔でスケジュールをチェックします)',
  //   ui.ButtonSet.YES_NO
  // );
  
  // if (response === ui.Button.YES) {
  //   Logger.log('ユーザーが起動を承認。予約投稿トリガー設定を開始');
  //   setTrigger();

  //   // 予定投稿のステータスを更新する
  //   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
  //   sheet.getRange(CONFIG.CELL_TRIGGER_STATUS).setValue(CONFIG.STRING_SHUT);

  //   ui.alert('予約投稿スケジューラーを起動しました');
  // } else {
  //   Logger.log('ユーザーが起動をキャンセル');
  // }

  Logger.log('ユーザーが起動を承認。予約投稿トリガー設定を開始');
  setTrigger();

  // 予定投稿のステータスを更新する
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
  sheet.getRange(CONFIG.CELL_TRIGGER_STATUS).setValue(CONFIG.STRING_SHUT);

  ui.alert('予約投稿スケジューラーを起動しました');

}

/**
 * 停止確認ダイアログを表示
 */
function showStopDialog() {
  Logger.log('予約投稿停止ダイアログを表示');
  const ui = SpreadsheetApp.getUi();
  // const response = ui.alert(
  //   '予約投稿スケジューラーの停止',
  //   '予約投稿スケジューラーを停止しますか？',
  //   ui.ButtonSet.YES_NO
  // );
  
  // if (response === ui.Button.YES) {
  //   Logger.log('ユーザーが停止を承認。予約投稿トリガー削除を開始');
  //   deleteTriggers();

  //   // 予定投稿のステータスを更新する
  //   const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
  //   sheet.getRange(CONFIG.CELL_TRIGGER_STATUS).setValue(CONFIG.STRING_SHUTDOWN);

  //   ui.alert('予約投稿スケジューラーを停止しました');
  // } else {
  //   Logger.log('ユーザーが停止をキャンセル');
  // }
  Logger.log('ユーザーが停止を承認。予約投稿トリガー削除を開始');
  deleteTriggers();

  // 予定投稿のステータスを更新する
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
  sheet.getRange(CONFIG.CELL_TRIGGER_STATUS).setValue(CONFIG.STRING_SHUTDOWN);

  ui.alert('予約投稿スケジューラーを停止しました');
}

// trigger.gs
/**
 * トリガーを設定
 */
function setTrigger() {
  Logger.log('トリガー設定を開始');
  
  // 既存のトリガーを削除
  deleteTriggers();
  
  try {
    // 1分おきに実行するトリガーを設定
    ScriptApp.newTrigger(CONFIG.TRIGGER_FUNCTION_NAME)
      .timeBased()
      .everyMinutes(1)
      .create();
    
    // トリガー作成日時を記録
    const now = new Date().toISOString();
    PropertiesService.getScriptProperties().setProperty('TRIGGER_CREATED_AT', now);
    Logger.log(`トリガーを正常に設定しました。作成日時: ${now}`);
  } catch (error) {
    Logger.log(`トリガー設定中にエラーが発生: ${error.message}`);
    throw error;
  }
}

/**
 * 既存のトリガーを全て削除
 */
function deleteTriggers() {
  Logger.log('トリガーの削除を開始');
  let deletedCount = 0;
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === CONFIG.TRIGGER_FUNCTION_NAME) {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    });
    
    PropertiesService.getScriptProperties().deleteProperty('TRIGGER_CREATED_AT');
    Logger.log(`${deletedCount}個のトリガーを削除しました`);
  } catch (error) {
    Logger.log(`トリガー削除中にエラーが発生: ${error.message}`);
    throw error;
  }
}

/**
 * 現在のトリガーの状態を取得
 * @return {Object} トリガーの状態
 */
function getTriggerStatus() {
  Logger.log('トリガーの状態確認を開始');
  
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const isActive = triggers.some(trigger => 
      trigger.getHandlerFunction() === CONFIG.TRIGGER_FUNCTION_NAME
    );
    
    const createdAt = PropertiesService.getScriptProperties().getProperty('TRIGGER_CREATED_AT');
    Logger.log(`トリガーの状態: アクティブ=${isActive}, 作成日時=${createdAt || 'なし'}`);
    
    return {
      isActive: isActive,
      createdAt: createdAt ? new Date(createdAt) : null
    };
  } catch (error) {
    Logger.log(`トリガー状態確認中にエラーが発生: ${error.message}`);
    throw error;
  }
}

// scheduler.gs
/**
 * スケジュールされた日時を取得する
 * @param {Array} rowData - 行データの配列
 * @return {Date|null} スケジュール日時
 */
function getScheduledDateTime(rowData) {
  const dateValue = rowData[CONFIG.DATE_COL]; // DATE_COL の値
  const hourValue = rowData[CONFIG.TIME_COL]; // TIME_COL の値
  const miniteValue = rowData[CONFIG.MINUTE_COL]; // TIME_COL の値
  
  // 修正された条件文
  if (dateValue == null || hourValue == null || miniteValue == null || isNaN(hourValue) || isNaN(miniteValue)) return null;

  try {
    const date = new Date(dateValue);
    date.setHours(Number(hourValue), Number(miniteValue), 0, 0);
    return date;
  } catch (error) {
    Logger.log(`日時の解析に失敗: ${error.message}`);
    return null;
  }
}

/**
 * スケジュールをチェックして投稿を実行
 */
function checkSchedule() {
  Logger.log('スケジュールチェックを開始');
  const startTime = new Date();
  let processedCount = 0;
  let postedCount = 0;
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.RESERVATION);
    const now = new Date();
    Logger.log(`現在時刻: ${now.toLocaleString()}`);
    
    // C列の最終データ行を取得
    const lastRow = sheet.getRange(sheet.getMaxRows(), 3) // 3はC列を示す
      .getNextDataCell(SpreadsheetApp.Direction.UP)
      .getRow();

    // データがない場合は処理しない
    if(CONFIG.START_ROW > lastRow){
      Logger.log("データがないため処理終了する")
      return;
    }

    // データ範囲を一括で取得
    const dataRange = sheet.getRange(
      CONFIG.START_ROW,
      2,
      lastRow - CONFIG.START_ROW + 1,
      32 // B-AG列を取得
    );
    const values = dataRange.getValues();

    // データを処理
    for (const [index, rowData] of values.entries()) {
      processedCount++;
      const rowIndex = CONFIG.START_ROW + index;
      
      // 投稿状態を確認
      if (rowData[CONFIG.POST_STATUS_COL]){
        Logger.log(`行${rowIndex}: 既に投稿済み (${rowData[0]})`);
        continue;  // continueを使用して次の行へ
      }
      
      // 予定日時を取得
      const scheduledDate = getScheduledDateTime(rowData);
      if (!scheduledDate) {
        Logger.log(`行${rowIndex}: 日時が正しく設定されていないためスキップします`);
        break;  // breakでループを完全に終了
      }
      
      Logger.log(`行${rowIndex}: 予定日時=${scheduledDate.toLocaleString()}`);
      
      // 現在時刻が予定日時を過ぎている場合、投稿を実行
      if (now >= scheduledDate) {
        Logger.log(`行${rowIndex}: 投稿処理を開始`);
        const statusUpdate = executePost(sheet, rowData, rowIndex, values, SHEETS_NAME.RESERVATION);
        postedCount++;
      }
    }    
    
    
    const endTime = new Date();
    const executionTime = (endTime - startTime) / 1000;
    Logger.log(`スケジュールチェック完了: 処理件数=${processedCount}, 投稿件数=${postedCount}, 実行時間=${executionTime}秒`);
  } catch (error) {
    Logger.log(`スケジュールチェック中にエラーが発生: ${error.message}`);
    throw error;
  }
}

/**
 * スレッズトークンリフレッシュのトリガーを設定する
 */
function setThreadsTokenRefreshTrigger() {
  Logger.log(`setThreadsTokenRefreshTrigger:Start`);
  setRefreshTrigger(CONFIG.THREADS_REFRESH_TRIGGER_ID, CONFIG.TRIGGER_FUNCTION_NAME_THREADS_REFRESH);
  Logger.log(`setThreadsTokenRefreshTrigger:END`);

}


/**
 * インスタトークンリフレッシュのトリガーを設定する
 */
function setInstaTokenRefreshTrigger() {
  Logger.log(`setInstaTokenRefreshTrigger:Start`);
  setRefreshTrigger(CONFIG.INSTA_REFRESH_TRIGGER_ID, CONFIG.TRIGGER_FUNCTION_NAME_INSTA_REFRESH);
  Logger.log(`setInstaTokenRefreshTrigger:END`);

}


/**
 * 対象のトリガーを設定する
 */
function setRefreshTrigger(triggerId, execMethod) {
  Logger.log(`setAutoPostTrigger:Start(${triggerId})`);
  
  try {
    // 既存のトリガーを削除
    deleteAutoPostTrigger(triggerId, execMethod);
    
    // 新しいトリガーを作成
    const trigger = ScriptApp.newTrigger(execMethod)
      .timeBased()
      .atHour(3)     // 午前3時に実行
      .everyDays(1)  // 毎日実行
      .create();

    // トリガーIDをプロパティに保存
    PropertiesService.getScriptProperties().setProperty(triggerId, trigger.getUniqueId());
    
    Logger.log(`Trigger set successfully`);
    return {
      success: true,
      message: `Trigger set successfully`,
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
function deleteAutoPostTrigger(triggerId, execMethod) {
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


/**
 * 認証用のモーダルダイアログを表示
 */
function showThreadsAuthDialog() {
  var html = HtmlService.createTemplateFromFile('threadsAuthDialog')
    .evaluate()
    .setWidth(600)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, '認証');
}

/**
 * フロントエンド用に認証URLを取得
 * @returns {string|null} 認証URL
 */
function getAuthUrl() {
  return getThreadsAuthorizationUrl();
}

/**
 * スクリプトのURLを取得
 * @return {string} スクリプトのURL
 */
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * ウェブアプリケーションのエントリーポイント
 * @return {HtmlOutput} HTML出力
 */
function doGet() {
  try {
    console.log('doGet called'); // 実行ログ
    
    // テンプレートを作成
    const template = HtmlService.createTemplateFromFile('Index');
    
    // システムプロパティを取得
    const systemProps = getSystemProperty();
    console.log('System Properties:', systemProps);
    
    // 設定値をログ出力
    console.log('Raw values:', {
      x: systemProps.webPostDefaultX,
      threads: systemProps.webPostDefaultThreads,
      instagram: systemProps.webPostDefaultInstagram
    });
    
    // テンプレートにデータを渡す
    template.defaultSettings = {
      x: String(systemProps.webPostDefaultX).toUpperCase() === 'ON',
      threads: String(systemProps.webPostDefaultThreads).toUpperCase() === 'ON',
      instagram: String(systemProps.webPostDefaultInstagram).toUpperCase() === 'ON'
    };
    
    console.log('Final template settings:', template.defaultSettings);
    
    // テンプレートを評価してHTMLを生成
    const htmlOutput = template.evaluate()
      .setTitle('まとめ投稿くん')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setFaviconUrl('https://www.google.com/images/favicon.ico')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
    
    console.log('HTML output generated successfully');
    return htmlOutput;
    
  } catch (error) {
    console.error('Error in doGet:', error);
    // エラーが発生した場合はエラーページを表示
    return HtmlService.createHtmlOutput(
      `<h1>エラーが発生しました</h1><p>${error.message}</p>`
    );
  }
}

/**
 * HTMLファイルをインクルードするためのヘルパー関数
 * @param {string} filename - インクルードするファイル名
 * @return {string} ファイルの内容
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * デフォルト設定を取得する
 * @return {Object} デフォルト設定オブジェクト
 */
function getDefaultSettings() {
  try {
    Logger.log('デフォルト設定の取得を開始');
    
    // システムプロパティを取得
    const systemProps = getSystemProperty();
    
    // 設定値の比較をログ出力
    Logger.log('設定値の比較:', {
      x: `"${systemProps.webPostDefaultX}" === "ON"`,
      threads: `"${systemProps.webPostDefaultThreads}" === "ON"`,
      instagram: `"${systemProps.webPostDefaultInstagram}" === "ON"`
    });
    
    // 設定オブジェクトを作成
    const settings = {
      x: systemProps.webPostDefaultX === "ON",
      threads: systemProps.webPostDefaultThreads === "ON",
      instagram: systemProps.webPostDefaultInstagram === "ON"
    };
    
    Logger.log('取得した設定:', settings);
    return settings;
    
  } catch (error) {
    Logger.log('デフォルト設定の取得中にエラーが発生:', error);
    // エラーが発生した場合はデフォルト値を返す
    return {
      x: false,
      threads: false,
      instagram: false
    };
  }
}

/**
 * デフォルト設定を保存する
 * @param {Object} settings - 保存する設定
 * @return {Object} 保存結果
 */
function saveDefaultSettings(settings) {
  try {
    Logger.log('デフォルト設定の保存を開始:', settings);
    
    // システムシートを取得
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS_NAME.SYSTEM);
    
    // 設定を保存
    sheet.getRange(PROPERTY_CELL.WEB_POST_DEFAULT_X).setValue(settings.x ? 'ON' : 'OFF');
    sheet.getRange(PROPERTY_CELL.WEB_POST_DEFAULT_THREADS).setValue(settings.threads ? 'ON' : 'OFF');
    sheet.getRange(PROPERTY_CELL.WEB_POST_DEFAULT_INSTAGRAM).setValue(settings.instagram ? 'ON' : 'OFF');
    
    Logger.log('デフォルト設定を保存しました');
    return { success: true };
    
  } catch (error) {
    Logger.log('デフォルト設定の保存中にエラーが発生:', error);
    return {
      success: false,
      error: error.message
    };
  }
}
