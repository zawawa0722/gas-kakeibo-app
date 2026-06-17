function importCsv(imButtonPressed) {

  // トリガ実行の場合はボタン実行フラグをオフにしておく
  if (typeof imButtonPressed !== "boolean") {
    imButtonPressed = false;
  }

  let SPREADSHEET, fileId, gotArray, basisCells;

  // スプレッドシート取得
  try {
    outputLog("INFO", "getSpreadsheet 開始", imButtonPressed);
    SPREADSHEET = getSpreadsheet();
    outputLog("INFO", "getSpreadsheet 成功", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "getSpreadsheet 失敗: " + e.stack, imButtonPressed);
    throw e;
  }

  // 取り込み対象のCSVファイルを取得
  try {
    outputLog("INFO", "getFileId 開始", imButtonPressed);
    fileId = getFileId(imButtonPressed);
    outputLog("INFO", "getFileId 成功", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "getFileId 失敗: " + e.stack, imButtonPressed);
    throw e;
  }

  // CSVファイルから今月分の明細を取得
  try {
    outputLog("INFO", "extractTargetLine 開始", imButtonPressed);
    gotArray = extractTargetLine(fileId, imButtonPressed);
    outputLog("INFO", "extractTargetLine 成功: " + gotArray.length + "件", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "extractTargetLine 失敗: " + e.stack, imButtonPressed);
    throw e;
  }

  // 対象のシート内から、出力対象となるセルを特定
  try {
    outputLog("INFO", "checkSheetsDate 開始", imButtonPressed);
    basisCells = checkSheetsDate(imButtonPressed);
    if (!basisCells) throw new Error("basisCells が undefined");
    outputLog("INFO", "checkSheetsDate 成功: " + basisCells.length + "件", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "checkSheetsDate 失敗: " + e.stack, imButtonPressed);
    throw e;
  }

  // gotArrayに格納した今月分の明細を、basisCellsで取得したセルに出力する
  try {
    outputLog("INFO", "取引データ処理開始", imButtonPressed);

    for (let i = 0; i < gotArray.length; i++) {
      let onlineFlug = false;
      let tempFlug_N = false;
      let tempFlug_S = false;
      let isMatched = false;

      // 支出に含まれない明細はスキップする
      if (gotArray[i][COL_VALID] === 0) {
        continue;
      }

      // 楽天カード・Amazonカードで決済している場合はオンラインフラグを付与
      if (
        (gotArray[i][COL_CARD] === CARD_CATEGORY.RAKUTEN_CARD ||
          gotArray[i][COL_CARD] === CARD_CATEGORY.AMAZON) &&
        gotArray[i][COL_CATEGORY] === EXPEND_CATEGORY.SHOPPING
      ) {
        onlineFlug = true;
      }

      // 一時負担のロジック
      if (gotArray[i][COL_TYPE] === EXPEND_CATEGORY.TEMP_N) {
        tempFlug_N = true;
        isMatched = true;
      } else if (gotArray[i][COL_TYPE] === EXPEND_CATEGORY.TEMP_S) {
        tempFlug_S = true;
        isMatched = true;
      }

      // 明細の解析と計算
      if (onlineFlug) {
        isMatched = onlineTransaction(gotArray[i], calculateOnline);
      } else {
        isMatched = expendTransaction(gotArray[i], calculateExpend, tempFlug_N, tempFlug_S);
      }

      // どのカテゴリともマッチしなかった場合は"その他"カテゴリへ仕分ける
      isMatched = unmatchedTransaction(
        gotArray[i],
        onlineFlug,
        isMatched,
        calculateOnline,
        calculateExpend
      );
    }

    outputLog("INFO", "取引データ処理終了", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "取引データ処理中に失敗: " + e.stack, imButtonPressed);
    throw e;
  }

  // シートに結果を出力
  applyBasisCells(SPREADSHEET, basisCells, calculateExpend, calculateOnline, imButtonPressed);

  // CSVインポート結果確認
  try {
    outputLog("INFO", "checkImportValue 開始", imButtonPressed);
    let emptyFlug = checkImportValue(basisCells);
    outputLog("INFO", "checkImportValue 成功", imButtonPressed);

    // CSVインポート結果通知
    compNotify(imButtonPressed, emptyFlug);
    outputLog("INFO", "compNotify 成功", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "checkImportValue / compNotify 失敗: " + e.stack, imButtonPressed);
    throw e;
  }

  outputLog("INFO", "importCsv 全体 正常終了", imButtonPressed);
}

// オンライン明細の計算ロジック
function onlineTransaction(row, calculateOnline) {
  let isMatched = false;

  for (const [key, value] of Object.entries(ONLINE_CATEGORY)) {
    const isHit = Array.isArray(value)
      ? value.some(v => row[COL_DESC].includes(v))
      : row[COL_DESC].includes(value);

    if (isHit) {
      let onlineKey = ONLINE_MAP[key];
      if (onlineKey && !isNaN(row[COL_AMOUNT])) {
        let positiveValue = Math.abs(Number(row[COL_AMOUNT]));
        if (isCancel(row)) {
          calculateOnline[onlineKey] -= positiveValue;
        } else {
          calculateOnline[onlineKey] += positiveValue;
        }
        isMatched = true;
      }
      break;
    }
  }
  return isMatched;
}

// オフライン明細の計算ロジック
function expendTransaction(row, calculateExpend, tempFlug_N, tempFlug_S) {
  let isMatched = false;

  for (const [key, value] of Object.entries(EXPEND_CATEGORY)) {
    if (row[COL_CATEGORY] == value) {
      // ワンバンクへの入金明細は一旦スキップ
      if (row[COL_DESC] === NYUUKIN || row[COL_DESC] === WANBANK) {
        isMatched = true;
        break;
      }

      let expendKey = EXPEND_MAP[key];
      if (expendKey && !isNaN(row[COL_AMOUNT])) {
        let positiveValue = Math.abs(Number(row[COL_AMOUNT]));

        if (isCancel(row)) {
          calculateExpend[expendKey] -= positiveValue;
        } else {
          calculateExpend[expendKey] += positiveValue;
        }
        isMatched = true;

        // 一時負担処理
        if (tempFlug_N) {
          if (isCancel(row)) {
            calculateExpend.tempNaoto -= positiveValue;
          } else {
            calculateExpend.tempNaoto += positiveValue;
          }
        }
        if (tempFlug_S) {
          if (isCancel(row)) {
            calculateExpend.tempNaoto -= positiveValue;
          } else {
            calculateExpend.tempNaoto += positiveValue;
          }
        }
      }
      break;
    }
  }
  return isMatched;
}

// CSVで集計した結果を実際のシートに書き戻す
function applyBasisCells(SPREADSHEET, basisCells, calculateExpend, calculateOnline, imButtonPressed) {
  try {
    outputLog("INFO", "basisCells 処理開始", imButtonPressed);

    let importCells = [];

    // basisCells から importCells を生成
    basisCells.forEach(cell => {
      if (cell.sheetName === SHEET_THOUBOKANRI) return;
      importCells.push({
        sheetName: cell.sheetName,
        row: cell.row + 1,
        column: cell.column
      });
    });

    // 各 importCell に対してシートを更新
    importCells.forEach(importCell => {
      let sheet = SPREADSHEET.getSheetByName(importCell.sheetName);
      let data = sheet.getRange("A:A").getValues();

      // 最終行を探索
      let lastRow = 0;
      for (let i = data.length - 1; i >= 0; i--) {
        if (data[i][0] !== "") {
          lastRow = i + 1;
          break;
        }
      }

      // シート種別ごとのカテゴリ定義
      let categoryData, mapData, calculateData;
      if (sheet.getName() === SHEET_SISYUTUKANRI) {
        categoryData = EXPEND_CATEGORY;
        mapData = EXPEND_MAP;
        calculateData = calculateExpend;
      } else if (sheet.getName() === SHEET_ONLINE) {
        categoryData = ONLINE_MEISAI_CATEGORY;
        mapData = ONLINE_MAP;
        calculateData = calculateOnline;
      } else {
        return;
      }

      // 各行のカテゴリに対応する集計値をセット
      for (let row = importCell.row; row < lastRow; row++) {
        let categoryValue = sheet.getRange(row, 1).getValue();
        for (let key in categoryData) {
          if (categoryValue === categoryData[key]) {
            let calculateKey = mapData[key];
            let calculatedValue = calculateData[calculateKey];
            // 基本的に上書きする
            sheet.getRange(row, importCell.column).setValue(calculatedValue);
            break;
          }
        }
      }
    });

    outputLog("INFO", "basisCells 処理終了", imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "basisCells 処理中に失敗: " + e.stack, imButtonPressed);
    throw e;
  }
}

// どのカテゴリにもマッチしなかった場合の処理
function unmatchedTransaction(row, onlineFlug, isMatched, calculateOnline, calculateExpend) {
  let matched = isMatched;

  // オンラインフラグありだが未分類の場合 → OTHER_ONLINE に加算
  if (onlineFlug && !matched) {
    if (!isNaN(row[COL_AMOUNT])) {
      let positiveValue = Math.abs(Number(row[COL_AMOUNT]));
      calculateOnline[ONLINE_MAP.OTHER_ONLINE] += positiveValue;
      matched = true;
    }
  }

  // それ以外の未分類 → OTHER に加算
  if (!matched) {
    if (!isNaN(row[COL_AMOUNT])) {
      let positiveValue = Math.abs(Number(row[COL_AMOUNT]));
      calculateExpend[EXPEND_MAP.OTHER] += positiveValue;
    }
  }

  return matched;
}

// キャンセル判定関数
function isCancel(row) {
  return row[COL_TYPE] && row[COL_TYPE].includes(CANCEL);
}
