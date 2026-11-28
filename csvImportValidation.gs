function checkSheetsDate(imButtonPressed) {
   // 結果を格納する配列
  let result = [];
  // imButtonPressed に基づいてターゲット日付を作成
  let targetDate;
  let SPREADSHEET = getSpreadsheet();
  if (imButtonPressed) {
    // G15 と H15 から年と月を取得して日付を生成 (帳簿管理を基準)
    const ledger = SPREADSHEET.getSheetByName(SHEET_THOUBOKANRI);
    let year = ledger.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].yearCell).getValue();
    let month = ledger.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].monthCell).getValue();
    month = String(month).padStart(2, "0"); // 前ゼロを付与
    targetDate = `${year}/${month}`;
  } else {
    // 本日の先月の日付を計算して yyyy/MM 形式で取得
    const { year, month } = getPreviousMonth_(SPREADSHEET.getSpreadsheetTimeZone());
    targetDate = `${year}/${String(month).padStart(2, '0')}`;
  }

  // 対象シートをループ処理
  [SHEET_THOUBOKANRI, SHEET_SISYUTUKANRI, SHEET_ONLINE].forEach(sheetName => {
    let sheet = SPREADSHEET.getSheetByName(sheetName);
    if (!sheet) {
      console.warn(`シート "${sheetName}" が見つかりませんでした。`);
      return; // または throw new Error(...)
    }

    let positions = CELL_POSITIONS[sheetName];
    let dateRange = sheet.getRange(positions.dateRange).getValues()[0]; // 日付行を取得

    const timezone = SPREADSHEET.getSpreadsheetTimeZone();
    const targetKey = normalizeMonthKey_(targetDate, timezone, true);
    // ターゲット日付と一致する列を探索
    dateRange.forEach((date, columnIndex) => {
      if (targetKey && normalizeMonthKey_(date, timezone, true) === targetKey) {
        // 一致する列が見つかった場合、その列と行を配列に追加
        result.push({
          sheetName: sheetName,
          row: sheet.getRange(positions.dateRange).getRow(),
          column: columnIndex + 2 // A列から始まるため +2
        });
      }
    });
  });

  return result; // 結果を返す
}


// CSVインポートの成否チェック
function checkImportValue(basicCells) {
  let SPREADSHEET = getSpreadsheet();
  const { year, month } = getPreviousMonth_(SPREADSHEET.getSpreadsheetTimeZone());
  const targetyyyyMM = `${year}/${String(month).padStart(2, '0')}`;

  let checkFlag = false; // ① チェックフラグをfalseに初期化

  // ② basicCells 内の対象シートを処理
  let targetSheets = [SHEET_SISYUTUKANRI, SHEET_ONLINE];
  
  for (let cellInfo of basicCells) {
    if (!targetSheets.includes(cellInfo.sheetName)) {
      continue; // 対象外のシートはスキップ
    }

    let sheet = SPREADSHEET.getSheetByName(cellInfo.sheetName);
    if (!sheet) continue;

    let dateRange = sheet.getRange(CELL_POSITIONS[cellInfo.sheetName].dateRange).getValues()[0]; // 1行分取得
    let targetColumn = findMonthIndex_(dateRange, targetyyyyMM, SPREADSHEET.getSpreadsheetTimeZone()); // ③ 一致する列を検索

    if (targetColumn === -1) {
      checkFlag = true; // ④ 一致する列がない場合、フラグを false に
      continue;
    }

    let dataRange;
    if (sheet.getName().toLowerCase() === SHEET_SISYUTUKANRI.toLowerCase()) {
      dataRange = sheet.getRange(4, targetColumn + 2, 22, 1);
    } else if (sheet.getName().toLowerCase() === SHEET_ONLINE.toLowerCase()) {
      dataRange = sheet.getRange(4, targetColumn + 2, 5, 1);
    }

    let values = dataRange.getValues();

    for (let row of values) {
      if (row[0] === null || row[0] === "") { // 値がnullまたは空文字なら
        checkFlag = true;
        break;
      }
    }
  }
  return checkFlag; // ⑤ フラグをリターン
}


