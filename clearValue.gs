// 履歴シートすべての2行目以降を削除する関数
function deleteAllHistorySheetsRows() {
  let sheetNames = [
    SHEET_THOUBOKANRI_BK,
    SHEET_SYUNYUKANRI_BK,
    SHEET_SISYUTUKANRI_BK,
    SHEET_ONLINE_BK
  ];

  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  sheetNames.forEach(sheetName => {
    let sheet = spreadsheet.getSheetByName(sheetName);

        // シートが存在しない場合はスキップ
    if (!sheet) {
        console.warn(`シート '${sheetName}' が見つかりません。`);
        return;
    }

    // 最終行を取得
    let lastRow = sheet.getLastRow();

    // 2行目以降を削除
    if (lastRow >= 2) {
        // 行を2行目から最終行まで削除
        sheet.deleteRows(2, lastRow - 1);
    }
  });
}

function clearSheetRanges() {
  // スプレッドシートのアクティブなドキュメントを取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // 各シートを処理
  for (var sheetName in CLEAR_RANGE) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (sheet) {
      // 定義された範囲をクリア
      CLEAR_RANGE[sheetName].backupRange.forEach(function(range) {
        sheet.getRange(range).clearContent();
      });
    } else {
      Logger.log("シートが見つかりません: " + sheetName);
    }
  }
}