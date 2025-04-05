const CONFIG = {
  SHEET_NAME: '予約投稿',    // スケジュールを管理するシート名
  START_ROW: 7,             // データ開始行
  END_ROW: 10006,            // データ終了行
  NO_COL : 0,               // No列
  DATE_COL: 1,              // 日付列
  TIME_COL: 2,              // 時間列
  MINUTE_COL: 3,            // 分列
  POST_TEXT : 4,            // 投稿文章
  POST_RES: 6,              // 返信投稿先
  POST_QUOTE: 7,            // 引用投稿先
  POST_QUOTE_THREADS: 8,    // 引用投稿先
  POST_CHECK_X: 9,          // 投稿ステータス:X
  POST_CHECK_THREADS: 10,   // 投稿ステータス:Threads
  POST_CHECK_INSTA: 11,     // 投稿ステータス:Instagram
  ATTACH_COUNT_X_COL: 13,   // 添付数(X)
  POST_STATUS_COL: 14,      // 投稿日時列
  X_POST_URL : 15,          // X投稿URL
  POST_URL : 16,            // Threads投稿URL
  INSTA_POST_URL : 17,      // Instagram投稿URL
  X_POST_ID_COL: 18,        // X投稿ID列 
  POST_ID_COL: 19,          // Threads投稿ID列 
  INSTA_POST_ID_COL: 20,    // Instagram投稿ID列 
  POST_ERROR: 21,           // エラー内容
  MEMO: 22,                 // メモ
  ATTACH_START_COLUMN: 23,  // 添付ファイル開始列
  ATTACH_END_COLUMN: 62,    // 添付ファイル終了列
  ATTACH_COLUMN_DIF: 2,     // 添付ファイル増加分
  ATTACH_X_CHECK_DIF: 1,    // 添付ファイルのチェックボックスの列差分 
  SHEET_ARRAY_COL_DIF : 2,   // シートと配列の列差分
  RESERVE_SHEET_COPY_COL_COUNT: 62,   // 予約投稿過去データへコピ－する対象列数

  ARCHIVE_COPY_CHECK_BOX_COL: 2, // 予約投稿過去データシートのコピー対象列
  ARCHIVE_START_ROW: 3,      // 予約投稿過去データシートの開始行
  ARCHIVE_END_ROW: 50005,    // 予約投稿過去データシートの終了行

  ROWS_PER_BATCH: 1000,  // 1回のバッチで処理する行数

  SHEET_COL_DELETE_CHECKBOX: 1,       // 削除用チェックボックス

  CELL_TRIGGER_STATUS: "G3",          // セル：実行ステータス
  CELL_SETTING_STATUS_X: "K5",        // セル：X設定ステータス
  CELL_SETTING_STATUS_THREADS: "L5",  // セル：Threads設定ステータス
  CELL_SETTING_STATUS_INSTA: "M5",    // セル：Instagram設定ステータス

  CELL_SETTING_CHECKBOX_X: "K7",        // セル：X設定ステータス
  CELL_SETTING_CHECKBOX_THREADS: "L7",  // セル：Threads設定ステータス
  CELL_SETTING_CHECKBOX_INSTA: "M7",    // セル：Instagram設定ステータス

  TRIGGER_FUNCTION_NAME_THREADS_REFRESH: "refreshLongTermToken",
  TRIGGER_FUNCTION_NAME_INSTA_REFRESH: "getLongTokenRefresh",

  TRIGGER_FUNCTION_NAME: 'checkSchedule', // トリガーで実行する関数名
  TRIGGER_FUNCTION_NAME_AUTO: "autoPostScheduler",         //
  AUTO_SCHEDULE_TRIGGER_ID: "AutoScheduleTrrigerID",
  THREADS_REFRESH_TRIGGER_ID: "ThreadsRefreshTrrigerID",
  INSTA_REFRESH_TRIGGER_ID: "InstaRefreshTrrigerID",

  IMAGE_EXTENSIONS: ['jpg', 'jpeg', 'png'],
  STRING_IMAGE: "image",
  STRING_VIDEO: "video",
  STRING_SHUT: "起動中",
  STRING_SHUTDOWN: "停止中",
  VIDEO_EXTENSIONS: ['mp4', 'mov'],
  X_SERVICE_NAME: 'twitter',

  DIALOG_TITLE: {
    GOOGLE_FOLDER_INPUT: "GoogleDriveのフォルダURL入力",
  },

  ENUM_INSTA_UL_TYPE:{
    SINGLE: 1,
    MULTI: 2,
  }

};

const SNS_CONVERT = {
  COL:{
    KEYWORD: 0,
    NAME: 1,
    X: 2,
    THREADS: 3,
    INSTAGRAM: 4,
  },
}

const CONFIG_HISTORY = {
  SHEET_ROW_START: 5,

  SHEET_COL_COPY_CHECK: 2,

};

const CONFIG_SETTING = {
  CELL_SCRIPT_ID: "B1",
};

const CONFIG_AUTO = {
  START_ROW: 7,     // 開始行
  END_ROW: 1006,    // 終了行
  SHEET_COL_DELETE_CHECKBOX: 4, // シート列：削除チェックボックス
  SHEET_COL_TEXT: 6,            // シート列：投稿文章

  CELL_CURRENT_ROW: "E3",         // セル位置：次投稿番号
  CELL_LIMIT_ROW: "E4",           // セル位置：最大番号
  CELL_SCHEDULE_START_TIME: "H3", // セル位置：スケジュール開始時間
  CELL_SCHEDULE_END_TIME: "I3",   // セル位置：スケジュール終了時間
  CELL_SCHEDULE_EXEC_TIME: "J3",  // セル位置：スケジュール実行間隔

}

const SHEETS_NAME = {
  RESERVATION: "予約投稿",
  AUTO: "自動投稿",
  HISTORY: "予約投稿過去データ",
  AUTO_HISTORY: "自動投稿過去データ",
  TEMPLATE: "template",
  TEMPLATE_AUTO: "Autotemplate",
  SYSTEM: 'system',
  SETTING: '設定用シート',
  MENTION: 'メンション情報',
};


const PROPERTY_CELL ={
  UL_IMAGE_FOLDER_ID: 'B2',
  UL_VIDEO_FOLDER_ID: 'B3',
  X_CLIENT_KEY: 'B4',
  X_CLIENT_SECRET: 'B5',
  X_CODE_VERIFIER: 'B6',
  X_OAUTH2_TWITTER: 'B7',
  X_USER_ID: 'B8',
  INSTA_APP_ID: 'B9',
  INSTA_APP_SECRET: 'B10',
  INSTA_USER_ID: 'B11',
  INSTA_BUSINESS_ID: 'B12',
  INSTA_SHORT_ACCESS_TOKEN: 'B13',
  INSTA_LONG_ACCESS_TOKEN: 'B14',
  INSTA_LONG_ACCESS_TOKEN_EXPIRY: 'B15',
  THREADS_LONG_TIME_TOKEN: 'B16',
  THREADS_LONG_TIME_TOKEN_EXPIRY: 'B17',
  THREADS_CLIENT_ID: 'B18',
  THREADS_CLIENT_SECRET: 'B19',
  SELECTED_SHEET_NAME: 'B20',
};
