/**
 * ログ出力用（成功・失敗どちらもここを使う）
 * ボタンクリック時の INFO ログはスキップ
 */
function outputLog(level, message, imButtonPressed) {
  try {
    // ボタンクリック時の INFO は記録しない
    if (imButtonPressed && level === "INFO") {
      return;
    }
    
    let ss = getSpreadsheet(); 
    let logSheet = ss.getSheetByName("ログ");
    if (!logSheet) {
      logSheet = ss.insertSheet();
      logSheet.setName("ログ");
      logSheet.appendRow(["日時", "実行タイプ", "レベル", "メッセージ"]);
    }

    let now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
    let runType = imButtonPressed ? "ボタンクリック" : "トリガー";

    logSheet.appendRow([now, runType, level, message]);

  } catch (e) {
    Logger.log("outputLog failed: " + e);
  }
}
