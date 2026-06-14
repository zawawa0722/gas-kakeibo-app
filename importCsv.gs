function importCsv(imButtonPressed) {

  let onlineFlug;
  let tempFlug_N;
  let tempFlug_S;
  let fileId = getFileId(imButtonPressed);
  let gotArray = extractTargetLine(fileId, imButtonPressed);
  let basisCells = checkSheetsDate(imButtonPressed);
  let importCells = [];

  if (!imButtonPressed) {
    imButtonPressed = false;
  }

  // ループで gotArray の各行をチェック
  for (let i = 0; i < gotArray.length; i++) {
    onlineFlug = false;
    tempFlug_N = false;
    tempFlug_S = false;
    let isMatched = false; // マッチングフラグ

    // オンラインフラグのロジック:
    if (
      (gotArray[i][7] === CARD_CATEGORY.RAKUTEN_CARD ||
        gotArray[i][7] === CARD_CATEGORY.VISA) &&
      gotArray[i][5] === EXPEND_CATEGORY.SHOPPING
    ) {
      onlineFlug = true;
    }

    // メモに "直人一時負担" または "沙羅一時負担" の記載があった場合は一時負担取引とする
    if (gotArray[i][3] === EXPEND_CATEGORY.TEMP_N) {
      tempFlug_N = true;
      isMatched = true;
    } else if (gotArray[i][3] === EXPEND_CATEGORY.TEMP_S) {
      tempFlug_S = true;
      isMatched = true;
    }

    // オンラインフラグが立っている場合
    if (onlineFlug) {
      // ONLINE_CATEGORY の値に部分一致する場合は対応する変数に加算
      for (const [key, value] of Object.entries(ONLINE_CATEGORY)) {
        // RAKUTENの値が「楽天」または「ﾗｸﾃﾝｲﾁﾊﾞ」に一致するか確認
        if (Array.isArray(value) ? value.some(v => gotArray[i][1].includes(v)) : gotArray[i][1].includes(value)) {
          let onlineKey = ONLINE_MAP[key];
          if (onlineKey) {
            if (!isNaN(gotArray[i][4])) {
              let positiveValue = Math.abs(Number(gotArray[i][4])); // 正の整数に変換
              // CANCEL が含まれている場合は減算、それ以外の場合は加算
              if (gotArray[i][3].includes(CANCEL)) {
                  calculateOnline[onlineKey] -= positiveValue;  // 減算
              } else {
                  calculateOnline[onlineKey] += positiveValue;  // 加算
              }
              isMatched = true; // マッチング成功
            }
          }
          break;
        }
      }
    } else {
      // EXPEND_CATEGORY の値に部分一致する場合は対応する変数に加算
      for (const [key, value] of Object.entries(EXPEND_CATEGORY)) {
        if (gotArray[i][5] == value) {
          let expendKey = EXPEND_MAP[key];
          if (expendKey) {
            if (!isNaN(gotArray[i][4])) {
              let positiveValue = Math.abs(Number(gotArray[i][4]));
              // CANCEL が含まれている場合は減算、それ以外の場合は加算
              if (gotArray[i][3].includes(CANCEL)) {
                  calculateExpend[expendKey] -= positiveValue;  // 減算
              } else {
                  calculateExpend[expendKey] += positiveValue;  // 加算
              }
              isMatched = true; // マッチング成功
            }
            if (tempFlug_N) {
              if (gotArray[i][3].includes(CANCEL)) {
                  calculateExpend.tempNaoto -= positiveValue;  // 減算
              } else {
                  calculateExpend.tempNaoto += positiveValue;  // 加算
              }
            }
            if (tempFlug_S) {
              if (gotArray[i][3].includes(CANCEL)) {
                  calculateExpend.tempNaoto -= positiveValue;  // 減算
              } else {
                  calculateExpend.tempNaoto += positiveValue;  // 加算
              }
            }
          }
          break;
        }
      }
    }

    // オンラインフラグが立っていて部分一致しない場合はOTHER_ONLINEに加算
    if (onlineFlug && !isMatched) {
      if (!isNaN(gotArray[i][4])) {
        let positiveValue = Math.abs(Number(gotArray[i][4]));
        calculateOnline[ONLINE_MAP.OTHER_ONLINE] += positiveValue; // OTHER_ONLINEに加算
        isMatched = true; // マッチ済みとする
      }
      //console.warn(`警告: オンライン取引の未分類データ - 行 ${i + 1}, 内容: ${JSON.stringify(gotArray[i])}`);
    }

    // マッチしない場合はOTHER（その他）に加算
    if (!isMatched) {
      if (!isNaN(gotArray[i][4])) {
        let positiveValue = Math.abs(Number(gotArray[i][4]));
        calculateExpend[EXPEND_MAP.OTHER] += positiveValue; // 雑費に加算
      }
      //console.warn(`警告: その他の未分類データ - 行 ${i + 1}, 内容: ${JSON.stringify(gotArray[i])}`);
    }
  }

  // basisCellsをループして条件に基づいて処理
  basisCells.forEach(cell => {
    // 帳簿管理の場合はスキップ
    if (cell.sheetName === SHEET_THOUBOKANRI) return;

    // importCellsに追加（columnはbasisCellsのものを引き継ぐ）
    importCells.push({
      sheetName: cell.sheetName,
      row: cell.row + 1,
      column: cell.column  // basisCellsのcolumnを引き継ぐ
    });
  });

  importCells.forEach(importCell => {
    let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(importCell.sheetName);

    // A列の最下行を取得（空白行を無視して最下行を取得）
    let data = sheet.getRange("A:A").getValues(); // A列全体を取得
    let lastRow = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      if (data[i][0] !== "") {
        lastRow = i + 1;
        break;
      }
    }

    // シート名に基づいてカテゴリを選択
    let categoryData, mapData;
    if (sheet.getName() === SHEET_SISYUTUKANRI) {
      categoryData = EXPEND_CATEGORY;
      mapData = EXPEND_MAP;  // EXPEND_MAPを使用
      calculateData = calculateExpend
    } else if (sheet.getName() === SHEET_ONLINE) {
      categoryData = ONLINE_MEISAI_CATEGORY;
      mapData = ONLINE_MAP;  // ONLINE_MAPを使用
      calculateData = calculateOnline
    } else {
      return; // その他のシートは処理をスキップ
    }

    // ① categoryCells の値を使って処理
    for (let row = importCell.row; row < lastRow; row++) {
      // categoryCellの値を取得
      let categoryValue = sheet.getRange(row, 1).getValue(); 

      // ② categoryValue がカテゴリに含まれているかチェック
      for (let key in categoryData) {
        if (categoryValue === categoryData[key]) {

          // ③ mapDataを経由して該当のキーを取得
          let calculateKey = mapData[key];

          // ④ calculatedataの値を取得
          let calculatedValue = calculateData[calculateKey];

          // セルの値が空であれば calculatedValue をセット
          if (!sheet.getRange(row, importCell.column).getValue()) {
            // importCells に入力
            sheet.getRange(row, importCell.column).setValue(calculatedValue); 
          }

          // 一度マッチしたら次のカテゴリーへ
          break; 
        }
      }
    }
  });
  // CSVインポートの成否チェック
  let emptyFlug = checkImportValue(basisCells);
  // インポートの成否通知
  compNotify(imButtonPressed,emptyFlug);
}


function checkImportValue(basicCells) {
  let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let today = new Date();
  let getlastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  let targetyyyyMM = Utilities.formatDate(getlastMonth, Session.getScriptTimeZone(), "yyyy/MM");

  let checkFlag = true; // ① チェックフラグをtrueに初期化

  // ② basicCells 内の対象シートを処理
  let targetSheets = [SHEET_SISYUTUKANRI, SHEET_ONLINE];
  
  for (let cellInfo of basicCells) {
    if (!targetSheets.includes(cellInfo.sheetName)) {
      continue; // 対象外のシートはスキップ
    }

    let sheet = spreadsheet.getSheetByName(cellInfo.sheetName);
    if (!sheet) continue;

    let dateRange = sheet.getRange(CELL_POSITIONS[cellInfo.sheetName].dateRange).getValues()[0]; // 1行分取得
    let targetColumn = dateRange.indexOf(targetyyyyMM); // ③ 一致する列を検索

    if (targetColumn === -1) {
      checkFlag = false; // ④ 一致する列がない場合、フラグを false に
      continue;
    }

    let dataRange = sheet.getRange(4, targetColumn + 2, 22, 1); // ④ 該当列の4~25行目を取得
    let values = dataRange.getValues();

    for (let row of values) {
      if (row[0] === null || row[0] === "") { // 値がnullまたは空文字なら
        checkFlag = false;
        break;
      }
    }
  }
  return checkFlag; // ⑤ フラグをリターン
}

