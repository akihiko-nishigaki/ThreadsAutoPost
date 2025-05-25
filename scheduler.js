// // scheduler.gs
// // スケジュール管理に関する機能を提供するファイル

// /**
//  * シートから予定日時を生成
//  * @param {Sheet} sheet スプレッドシート
//  * @param {number} row 行番号
//  * @return {Date|null} 生成された日時、不正な場合はnull
//  */
// function getScheduledDateTime(sheet, row) {
//   const date = sheet.getRange(CONFIG.DATE_COL + row).getValue();
//   const hour = sheet.getRange(CONFIG.TIME_COL + row).getValue();
//   const minute = sheet.getRange(CONFIG.MINUTE_COL + row).getValue();
  
//   // 値の存在チェック
//   if (!date || !hour || !minute) return null;
  
//   try {
//     const scheduledDate = new Date(date);
//     scheduledDate.setHours(hour);
//     scheduledDate.setMinutes(minute);
//     scheduledDate.setSeconds(0);
//     return scheduledDate;
//   } catch (e) {
//     console.error(`Invalid date format at row ${row}: ${e.message}`);
//     return null;
//   }
// }
