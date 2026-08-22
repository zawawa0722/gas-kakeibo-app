function importCsv(imButtonPressed) {

  // トリガ実行の場合はボタン実行フラグをオフにしておく
  if (typeof imButtonPressed !== "boolean") {
    imButtonPressed = false;
  }

  let SPREADSHEET, fileId, gotArray, basisCells, settingContents;
  let flagObj = { hasWarnOrError: false };
  outputLog("INFO", "importCsv処理を開始します", imButtonPressed, flagObj);

  // スプレッドシート取得
  try {
    SPREADSHEET = getSpreadsheet();
  } catch (e) {
    outputLog("ERROR", "スプレッドシート取得処理でエラーが発生しました。: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // 設定シートの値を取得
  try {
    settingContents = loadSheetSetting(SPREADSHEET);
    SHEET_SETTING_CONTENTS = settingContents.settingContents
    SHOP_LIST_RAKUTEN = settingContents.shopListRakuten
    SHOP_LIST_OTHER = settingContents.shopListOther
  } catch (e) {
    outputLog("ERROR", "設定シート読み込み処理でエラーが発生しました。: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // プログラム内部で保持するユーザ名に設定シートの情報を反映
  try {
    updateExpendCategoryNames();
  } catch (e) {
    outputLog("ERROR", "支出項目名の更新処理でエラーが発生しました。: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // 取り込み対象のCSVファイルを取得
  try {
    fileId = getFileId(imButtonPressed);
  } catch (e) {
    outputLog("ERROR", "CSVファイル取得処理でエラーが発生しました。: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // CSVファイルから今月分の明細を取得
  try {
    gotArray = extractTargetLine(fileId, imButtonPressed);
    outputLog("INFO", "今月分の明細カウント数: " + gotArray.length + "件", imButtonPressed, flagObj);
  } catch (e) {
    outputLog("ERROR", "今月の明細取得処理でエラ-が発生しました: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // 対象のシート内から、出力対象となるセルを特定
  try {
    basisCells = checkSheetsDate(imButtonPressed);
    if (!basisCells) throw new Error("basisCells が undefined");
  } catch (e) {
    outputLog("ERROR", "明細出力先セルの特定処理でエラーが発生しました: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // gotArrayに格納した今月分の明細を、basisCellsで取得したセルに出力する
  try {

    for (let i = 0; i < gotArray.length; i++) {
      let onlineCardFlug = false;
      let tempFlug_A = false;
      let tempFlug_B = false;
      let isMatched = false;
      let knownCards = Object.values(CARD_CATEGORY);
      let usedCard = gotArray[i][COL_CARD];

      // 支出に含まれない明細はスキップする
      if (gotArray[i][COL_VALID] === 0) {
        continue;
      }

      // 該当するカードがない場合はエラー扱いとする(ローンは除く)
      if (gotArray[i][COL_CATEGORY] !== EXPEND_CATEGORY.LOAN && !knownCards.includes(usedCard)) {
        throw new Error(
          `未知のカード名です（設定と不一致）\nカード名: ${usedCard}\n明細: ${gotArray[i][COL_DESC]}`
        );
      }

      // 楽天カード・Amazonカードで決済している場合はオンラインフラグを、ワンバンクカードの場合はワンバンクカードフラグを付与
      if (
        (gotArray[i][COL_CARD] === CARD_CATEGORY.RAKUTEN_CARD ||
          gotArray[i][COL_CARD] === CARD_CATEGORY.AMAZON) &&
        gotArray[i][COL_CATEGORY] === EXPEND_CATEGORY.SHOPPING
      ) {
        onlineCardFlug = true;
      } 

      // 一時負担のロジック
      if (gotArray[i][COL_TYPE] === EXPEND_CATEGORY.TEMP_USER_A) {
        tempFlug_A = true;
        isMatched = true;
      } else if (gotArray[i][COL_TYPE] === EXPEND_CATEGORY.TEMP_USER_B) {
        tempFlug_B = true;
        isMatched = true;
      }

      // 明細の解析と計算
      if (onlineCardFlug) {
        isMatched = onlineTransaction(gotArray[i], calculateOnline, flagObj);
      } else {
        isMatched = expendTransaction(gotArray[i], calculateExpend, tempFlug_A, tempFlug_B);
      }

      // どのカテゴリともマッチしなかった場合は"その他"カテゴリへ仕分ける
      isMatched = unmatchedTransaction(
        gotArray[i],
        onlineCardFlug,
        isMatched,
        calculateOnline,
        calculateExpend,
        flagObj
      );
    }

  } catch (e) {
    outputLog("ERROR", "取引データ反映処理でエラーが発生しました: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  // オンライン出力が 0 件かどうか判定（停止はしない）
  let hasOnline = false;
  for (let key in calculateOnline) {
    if (calculateOnline[key] !== 0) {
      hasOnline = true;
      break;
    }
  }
  if (!hasOnline) {
    outputLog(
      "WARN",
      "オンライン明細の出力が 0 件でした（CSV内のカード名が変更された可能性あり）",
      imButtonPressed,
      flagObj
    );
  }

  // 一般支出の「その他」カテゴリが 10000円以上の場合は警告扱いとする
  if (calculateExpend.otherexp > 10000){
    outputLog(
      "WARN",
      "支出カテゴリ「その他」が10000円以上です。\n「その他」の明細を確認してください。",
      imButtonPressed,
      flagObj
    );
  }

  // シートに結果を出力
  applyBasisCells(SPREADSHEET, basisCells, calculateExpend, calculateOnline, imButtonPressed, flagObj);

  // CSVインポート結果確認
  try {
    let emptyFlug = checkImportValue(basisCells);
    // CSVインポート結果通知
    compNotify(imButtonPressed, emptyFlug, flagObj.hasWarnOrError);
  } catch (e) {
    outputLog("ERROR", "CSVインポート処理/CSVインポート結果通知処理でエラーが発生しました: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }

  outputLog("INFO", "importCsv処理が正常終了しました", imButtonPressed, flagObj);
}

// オンライン明細の計算ロジック
function onlineTransaction(row, calculateOnline, flagObj) {
  let isMatched = false;
  let categoryKey = null;

  // 1. ONLINE_CATEGORY の判定（OTHER_ONLINE / UNKNOWN_ONLINE 以外）
  for (const [key, value] of Object.entries(ONLINE_CATEGORY)) {
    if (key === 'OTHER_ONLINE' || key === 'UNKNOWN_ONLINE') continue;

    const isHit = Array.isArray(value)
      ? value.some(v => row[COL_DESC].includes(v))
      : row[COL_DESC].includes(value);

    if (isHit) {
      categoryKey = key;
      break;
    }
  }

  // 2. 判定①：ONLINE_CATEGORY でヒットしない場合、SHOP_LIST_RAKUTEN 内を検索
  if (!categoryKey && Array.isArray(SHOP_LIST_RAKUTEN)) {
    if (SHOP_LIST_RAKUTEN.some(shop => row[COL_DESC].includes(shop))) {
      categoryKey = 'RAKUTEN';
    }
  }

  // 3. 判定②：①でもヒットしない場合、SHOP_LIST_OTHER 内を検索
  if (!categoryKey) {
    if (Array.isArray(SHOP_LIST_OTHER) && SHOP_LIST_OTHER.some(shop => row[COL_DESC].includes(shop))) {
      categoryKey = 'OTHER_ONLINE';
    } else {
      categoryKey = 'UNKNOWN_ONLINE';
    }
  }

  // 4. カテゴリに応じた金額集計処理
  if (categoryKey && !isNaN(row[COL_AMOUNT])) {
    // UNKNOWN_ONLINE の場合は「その他」のマップキー（ONLINE_MAP['OTHER_ONLINE']）に合算
    let mapLookupKey = (categoryKey === 'UNKNOWN_ONLINE') ? 'OTHER_ONLINE' : categoryKey;
    let onlineKey = ONLINE_MAP[mapLookupKey];

    if (onlineKey) {
      let positiveValue = Math.abs(Number(row[COL_AMOUNT]));

      // 集計用キーの初期化
      if (calculateOnline[onlineKey] === undefined) {
        calculateOnline[onlineKey] = 0;
      }

      if (isCancel(row)) {
        calculateOnline[onlineKey] -= positiveValue;
      } else {
        calculateOnline[onlineKey] += positiveValue;
      }
      isMatched = true;

      // 不明カテゴリの明細はログ出力する
      if (categoryKey === 'UNKNOWN_ONLINE' && typeof outputLog === 'function') {
        outputLog(
          "WARN",
          `不明カテゴリ明細："${row[COL_DESC]}" / ${row[COL_AMOUNT]}`,
          false,
          flagObj // ※ flagObj を onlineTransaction の引数に追加するか、グローバル変数として参照
        );
      }
    }
  }
  return isMatched;
}

// オフライン明細の計算ロジック
function expendTransaction(row, calculateExpend, tempFlug_A, tempFlug_B) {
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
        if (tempFlug_A) {
          if (isCancel(row)) {
            calculateExpend.temp_user_A -= positiveValue;
          } else {
            calculateExpend.temp_user_A += positiveValue;
          }
        }
        if (tempFlug_B) {
          if (isCancel(row)) {
            calculateExpend.temp_user_B -= positiveValue;
          } else {
            calculateExpend.temp_user_B += positiveValue;
          }
        }
      }
      break;
    }
  }
  return isMatched;
}

// CSVで集計した結果を支出管理シートに書き落とす
function applyBasisCells(SPREADSHEET, basisCells, calculateExpend, calculateOnline, imButtonPressed, flagObj) {
  try {
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
        // 書き込み対象のセルをあらかじめ取得
        let targetRange = sheet.getRange(row, importCell.column);
        
        // セルに設定されている数式を取得
        let formula = targetRange.getFormula();
        
        // 数式の場合はそのカテゴリの書き込み処理をスキップする
        if (formula !== "") {
          continue; 
        }

        let categoryValue = sheet.getRange(row, 1).getValue();
        for (let key in categoryData) {
          if (categoryValue === categoryData[key]) {
            let calculateKey = mapData[key];
            let calculatedValue = calculateData[calculateKey];
            
            // 数式が入っていない場合のみ、上書き入力する
            targetRange.setValue(calculatedValue);
            break;
          }
        }
      }
    });
  } catch (e) {
    outputLog("ERROR", "basisCells 処理中にエラーが発生しました: " + e.stack, imButtonPressed, flagObj);
    throw e;
  }
}

// どのカテゴリにもマッチしなかった場合の処理（修正後）
function unmatchedTransaction(row, onlineCardFlug, isMatched, calculateOnline, calculateExpend, flagObj) {
  let matched = isMatched;

  // それ以外の未分類（オンライン以外のカード利用等） → OTHER に加算
  if (!matched) {
    if (!isNaN(row[COL_AMOUNT])) {
      let positiveValue = Math.abs(Number(row[COL_AMOUNT]));
      let key = EXPEND_MAP.OTHER;

      if (calculateExpend[key] === undefined) {
        calculateExpend[key] = 0;
      }

      if (typeof isCancel === 'function' && isCancel(row)) {
        calculateExpend[key] -= positiveValue;
      } else {
        calculateExpend[key] += positiveValue;
      }
    }
  }

  return matched;
}

// キャンセル判定関数
function isCancel(row) {
  return row[COL_TYPE] && row[COL_TYPE].includes(CANCEL);
}
