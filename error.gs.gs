function outputErrorLog(e, triggerFlag) {
  // ログシートを取得
  let TODAY = new Date();
  let logsheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_LOG);
  
  // A列のデータを取得して、最終行を計算
  let lastRow = logsheet.getLastRow();
  let values = [];

  if (lastRow >= 1) {
    values = logsheet.getRange(1, 1, lastRow).getValues();
  } else {
    values = []
  }

  let outputRow = values.filter(row => row[0] !== "").length + 1; // 次の空いている行
  
  // 現在の時刻を取得
  let formattedDate = Utilities.formatDate(TODAY, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  
  // 実行ステータスの設定
  let extStatus;
  if (triggerFlag == true) {
    extStatus = MANUAL;
  } else if (triggerFlag == false) {
    extStatus = TORIGGER;
  }

  // 出力するデータをまとめる
  let logData = [
  formattedDate,
  `実行ステータス：${extStatus}`,
  "エラーメッセージ : " + e.toString(),
  "エラー発生箇所 : " + e.stack
].join(' ');

  // ログ出力
  logsheet.getRange(outputRow, 1).setValue(logData);
  Logger.log(JSON.stringify(logData));
}
