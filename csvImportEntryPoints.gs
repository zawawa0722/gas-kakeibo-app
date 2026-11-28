// ボタンフラグ取得
function onButtonClick() {
  // 実行前に確認メッセージを表示
  let result = Browser.msgBox("確認", "家計簿出力処理を実行してもよろしいですか？", Browser.Buttons.YES_NO);

  // YESボタンが押された場合のみスクリプトを実行
  if (result == "yes") {
    let buttonPressed = true;
    importCsv(buttonPressed);
  } else {
    // NOボタンが押された場合は実行しない
    Logger.log("家計簿出力処理は実行されませんでした");
  }
}

// ボタン実行は帳簿管理の選択年、自動実行は前月の年を使う。
function getFileId(isButtonPressed) {
  const spreadsheet = getSpreadsheet();
  const year = isButtonPressed
    ? spreadsheet.getSheetByName(SHEET_THOUBOKANRI)
      .getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].yearCell).getValue()
    : getPreviousMonth_(spreadsheet.getSpreadsheetTimeZone()).year;
  const file = findOneBankCsv_(year);
  if (!file) throw new Error("指定されたCSVファイルが見つかりませんでした");
  return file.getId();
}
