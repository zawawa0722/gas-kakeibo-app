// 4つのシートに対する値とフォーマットのコピーを追加
function backupSheets() {
  let sheetPairs = [
    { source: SHEET_THOUBOKANRI, backup: SHEET_THOUBOKANRI_BK },
    { source: SHEET_SYUNYUKANRI, backup: SHEET_SYUNYUKANRI_BK },
    { source: SHEET_SISYUTUKANRI, backup: SHEET_SISYUTUKANRI_BK },
    { source: SHEET_ONLINE, backup: SHEET_ONLINE_BK },
  ];

  sheetPairs.forEach(pair => {
    let sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pair.source);
    let historySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(pair.backup);

    if (!sourceSheet || !historySheet) {
      throw new Error(`シート '${pair.source}' または '${pair.backup}' が見つかりません。`);
    }

    let backupRange = BACKUP_RANGE[pair.source].backupRange;
    let backupValues = sourceSheet.getRange(backupRange).getValues();
    let backupColors = sourceSheet.getRange(backupRange).getBackgrounds();
    let backupFormats = sourceSheet.getRange(backupRange).getNumberFormats(); // 書式も取得

    // 最下行を取得
    let lastRow = historySheet.getRange("A:A").getValues().filter(String).length;

    // lastRow に応じて pasteRow を設定
    let pasteRow = lastRow === 1 ? 2 : (lastRow === 0 ? 2 : 3);

    // 値を貼り付け
    let pasteRange = historySheet.getRange(pasteRow, 1, backupValues.length, backupValues[0].length);
    pasteRange.setValues(backupValues);

    // 背景色を貼り付け
    pasteRange.setBackgrounds(backupColors);

    // 書式を貼り付け
    pasteRange.setNumberFormats(backupFormats);

    // 枠線を設定
    pasteRange.setBorder(
      true, true, true, true,  // 上, 左, 下, 右の枠線を有効化
      true, true               // 内部の横線, 縦線も有効化
    );
  });
}
