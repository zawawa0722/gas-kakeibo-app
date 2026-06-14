function outputErrorLog(e, triggerFlag) {
  // ログシートを取得
  let logsheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("ログ");
  
  // A列のデータを取得して、最終行を計算
  let values = logsheet.getRange(1, 1, logsheet.getLastRow()).getValues();
  let outputRow = values.filter(row => row[0] !== "").length + 1; // 次の空いている行
  
  // 現在の時刻を取得
  let now = new Date();
  let formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  
  // 実行ステータスの設定
  let extStatus;
  if (triggerFlag == true) {
    extStatus = "MANUAL"; // 手動実行
  } else if (triggerFlag == false) {
    extStatus = "TORIGGER"; // トリガー実行
  }

  // 出力するデータをまとめる
  let logData = [
    [formattedDate],
    [`実行ステータス：${extStatus}`],
    ["エラーメッセージ : " + e.toString()],
    ["エラー発生箇所 : " + e.stack]
  ];

  // 出力先の範囲に3行分を一度に出力
  logsheet.getRange(outputRow, 1, 3, 1).setValues(logData);
}
