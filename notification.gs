// CSVファイル配置依頼
function sendEmailRequestCsvStorage() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_CSV_STORAGE, BODY_REQUEST_CSV_STORAGE);
}

// インポート結果通知
function compNotify(buttonFlug, emptyCheckFlug, hasWarnOrError) {
  // ボタンフラグと支出シートの空チェックフラグでインポート成否を判定
  if (buttonFlug == false){
    if (emptyCheckFlug == false){
        if (hasWarnOrError) {
          GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT_INCLUDE_ERROR);
        }else{
          GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT);
        }
    }else{
      GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_FAILED_IMPORT, BODY_FAILED_IMPORT);
    }
  }
}

// バックアップリクエスト
function sendEmailRequestBuckup() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_BUCKUP, BODY_REQUEST_BUCKUP);
}

// 振込依頼通知
function sendFamilyTransferRequest() {

  let imButtonPressed = false
  let message
  SPREADSHEET = getSpreadsheet();
  
  // getFileIdのエラー確認
  try {
    message = createMessage()
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }
  let options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify({
      "messages": [{
        "type": "text",
        "text": message
      }]
    })
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", options);
}

// 収支結果を取得
function getResaltBalance(sheet, dateRange, lastMonthText) {

  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);
  let resultTable = ""

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }
  
  // 項目名を取得 (A8:A12)
  let labels = sheet.getRange("A8:A12").getValues().map(row => row[0]);

  // 8〜12行目の値を取得
  let values = sheet.getRange(8, lastMonthIndex + 2, 5, 1).getValues();

  // 表形式にするため、項目名と値を結びつける
  let formattedValues = {}; // 項目名と金額の格納用

  // 8〜12行目の値を整形
  for (let i = 0; i < labels.length; i++) {
    let label = labels[i];
    let value = values[i][0];

    // 「貯金率」の場合は100を掛けてパーセント表記にする
    if (label === "貯金率") {
      value = (isNaN(parseFloat(value)) || value === "") ? 0 : parseFloat(value) * 100; // 100を掛ける
      formattedValues[label] = value.toFixed(2) + "%"; // 小数点以下2桁でパーセント表示
      resultTable += `${label}: ${formattedValues[label]}`;
    } else {
      // 数値として変換可能な場合のみ、金額を表示
      value = (isNaN(parseInt(value)) || value === "") ? 0 : parseInt(value); // 数値に変換できない場合は0に
      formattedValues[labels[i]] = value;
      resultTable += `${label}: ${value.toLocaleString()}円\n`; // 表形式に追加
    }
  }
  return resultTable
}

// 支出結果を取得
function getExpenses(sheet, dateRange, lastMonthText) {
  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);
  let resultTable = ""

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }
  
  // 項目名を取得 A4:A25)
  let labels = sheet.getRange("A4:A25").getValues().map(row => row[0]);

  // 4〜25行目の値を取得
  let values = sheet.getRange(4, lastMonthIndex + 2, 22, 1).getValues();

  // 表形式にするため、項目名と値を結びつける
  let formattedValues = {}; // 項目名と金額の格納用

  // 8〜12行目の値を整形
  for (let i = 0; i < labels.length; i++) {
    let label = labels[i];
    let value = values[i][0];

    // 数値として変換可能な場合のみ、金額を表示
    value = (isNaN(parseInt(value)) || value === "") ? 0 : parseInt(value); // 数値に変換できない場合は0に
    formattedValues[labels[i]] = value;
    resultTable += `${label}: ${value.toLocaleString()}円\n`; // 表形式に追加
  }
  return resultTable
}

// LINE通知用メッセージ作成
function createMessage() {

  let imButtonPressed = false
  let tyoubosheet = SPREADSHEET.getSheetByName(SHEET_THOUBOKANRI);
  let sisyutusheet = SPREADSHEET.getSheetByName(SHEET_SISYUTUKANRI);
  let tyoubodateRange = CELL_POSITIONS[SHEET_THOUBOKANRI].dateRange;
  let sisyutudateRange = CELL_POSITIONS[SHEET_SISYUTUKANRI].dateRange;
  let incomeAndExpenditure
  let spending
  let feedBack
  let TODAY = new Date();

  // 先月のyyyy/MM表記を生成 (ゼロ埋め対応)
  let lastMonth = new Date(TODAY.getFullYear(), TODAY.getMonth() - 1, 1);
  let lastMonthText = `${lastMonth.getFullYear()}/${String(lastMonth.getMonth() + 1).padStart(2, "0")}`;
  let lastMonthTextNotZero = `${String(lastMonth.getMonth() + 1)}`
  let dateValues = tyoubosheet.getRange(tyoubodateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);
  let headComment = `以下が${lastMonthTextNotZero}月の結果です。\nお疲れ様でした!`;

  // 帳簿管理のA8~A12の項目を取得
  try {
    incomeAndExpenditure = getResaltBalance(tyoubosheet, tyoubodateRange, lastMonthText)
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }
  
  // 支出管理の全項目を取得
  try {
    spending = getExpenses(sisyutusheet, sisyutudateRange, lastMonthText)
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }

  // 帳簿管理の6行目と7行目の値を取得
  let colIndex = lastMonthIndex + 2; // 列番号（1始まり）
  let row6Value = tyoubosheet.getRange(6, colIndex).getValue(); // 6行目の値
  let row7Value = tyoubosheet.getRange(7, colIndex).getValue(); // 7行目の値
  
  // 収支結果と支出明細をもとにAI分析したコメントを取得
  try {
    feedBack = analyzeHouseholdBudget(incomeAndExpenditure, spending)
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }

  // メッセージ作成
  let messageBody = `${headComment}\n\n${incomeAndExpenditure}\n\n${feedBack}\n\n★今月の2人の振込金額\n直人：${row6Value.toLocaleString()}円\n沙羅：${row7Value.toLocaleString()}円`;

  return messageBody
}
// CSVファイル配置依頼
function sendEmailRequestCsvStorage() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_CSV_STORAGE, BODY_REQUEST_CSV_STORAGE);
}

// インポート結果通知
function compNotify(buttonFlug, emptyCheckFlug, hasWarnOrError) {
  // ボタンフラグと支出シートの空チェックフラグでインポート成否を判定
  if (buttonFlug == false){
    if (emptyCheckFlug == false){
        if (hasWarnOrError) {
          GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT_INCLUDE_ERROR);
        }else{
          GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT);
        }
    }else{
      GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_FAILED_IMPORT, BODY_FAILED_IMPORT);
    }
  }
}

// バックアップリクエスト
function sendEmailRequestBuckup() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_BUCKUP, BODY_REQUEST_BUCKUP);
}

// 振込依頼通知
function sendFamilyTransferRequest() {

  let imButtonPressed = false
  let message
  SPREADSHEET = getSpreadsheet();
  
  // getFileIdのエラー確認
  try {
    message = createMessage()
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }
  let options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify({
      "messages": [{
        "type": "text",
        "text": message
      }]
    })
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", options);
}

// 収支結果を取得
function getResaltBalance(sheet, dateRange, lastMonthText) {

  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);
  let resultTable = ""

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }
  
  // 項目名を取得 (A8:A12)
  let labels = sheet.getRange("A8:A12").getValues().map(row => row[0]);

  // 8〜12行目の値を取得
  let values = sheet.getRange(8, lastMonthIndex + 2, 5, 1).getValues();

  // 表形式にするため、項目名と値を結びつける
  let formattedValues = {}; // 項目名と金額の格納用

  // 8〜12行目の値を整形
  for (let i = 0; i < labels.length; i++) {
    let label = labels[i];
    let value = values[i][0];

    // 「貯金率」の場合は100を掛けてパーセント表記にする
    if (label === "貯金率") {
      value = (isNaN(parseFloat(value)) || value === "") ? 0 : parseFloat(value) * 100; // 100を掛ける
      formattedValues[label] = value.toFixed(2) + "%"; // 小数点以下2桁でパーセント表示
      resultTable += `${label}: ${formattedValues[label]}`;
    } else {
      // 数値として変換可能な場合のみ、金額を表示
      value = (isNaN(parseInt(value)) || value === "") ? 0 : parseInt(value); // 数値に変換できない場合は0に
      formattedValues[labels[i]] = value;
      resultTable += `${label}: ${value.toLocaleString()}円\n`; // 表形式に追加
    }
  }
  return resultTable
}

// 支出結果を取得
function getExpenses(sheet, dateRange, lastMonthText) {
  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);
  let resultTable = ""

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }
  
  // 項目名を取得 A4:A25)
  let labels = sheet.getRange("A4:A25").getValues().map(row => row[0]);

  // 4〜25行目の値を取得
  let values = sheet.getRange(4, lastMonthIndex + 2, 22, 1).getValues();

  // 表形式にするため、項目名と値を結びつける
  let formattedValues = {}; // 項目名と金額の格納用

  // 8〜12行目の値を整形
  for (let i = 0; i < labels.length; i++) {
    let label = labels[i];
    let value = values[i][0];

    // 数値として変換可能な場合のみ、金額を表示
    value = (isNaN(parseInt(value)) || value === "") ? 0 : parseInt(value); // 数値に変換できない場合は0に
    formattedValues[labels[i]] = value;
    resultTable += `${label}: ${value.toLocaleString()}円\n`; // 表形式に追加
  }
  return resultTable
}

// LINE通知用メッセージ作成
function createMessage() {

  let imButtonPressed = false
  let tyoubosheet = SPREADSHEET.getSheetByName(SHEET_THOUBOKANRI);
  let sisyutusheet = SPREADSHEET.getSheetByName(SHEET_SISYUTUKANRI);
  let tyoubodateRange = CELL_POSITIONS[SHEET_THOUBOKANRI].dateRange;
  let sisyutudateRange = CELL_POSITIONS[SHEET_SISYUTUKANRI].dateRange;
  let incomeAndExpenditure
  let spending
  let feedBack
  let TODAY = new Date();

  // 先月のyyyy/MM表記を生成 (ゼロ埋め対応)
  let lastMonth = new Date(TODAY.getFullYear(), TODAY.getMonth() - 1, 1);
  let lastMonthText = `${lastMonth.getFullYear()}/${String(lastMonth.getMonth() + 1).padStart(2, "0")}`;
  let lastMonthTextNotZero = `${String(lastMonth.getMonth() + 1)}`
  let dateValues = tyoubosheet.getRange(tyoubodateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);
  let headComment = `以下が${lastMonthTextNotZero}月の結果です。\nお疲れ様でした!`;

  // 帳簿管理のA8~A12の項目を取得
  try {
    incomeAndExpenditure = getResaltBalance(tyoubosheet, tyoubodateRange, lastMonthText)
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }
  
  // 支出管理の全項目を取得
  try {
    spending = getExpenses(sisyutusheet, sisyutudateRange, lastMonthText)
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }

  // 帳簿管理の6行目と7行目の値を取得
  let colIndex = lastMonthIndex + 2; // 列番号（1始まり）
  let row6Value = tyoubosheet.getRange(6, colIndex).getValue(); // 6行目の値
  let row7Value = tyoubosheet.getRange(7, colIndex).getValue(); // 7行目の値
  
  // 収支結果と支出明細をもとにAI分析したコメントを取得
  try {
    feedBack = analyzeHouseholdBudget(incomeAndExpenditure, spending)
  } catch (e) {
    outputErrorLog(e, imButtonPressed);
  }

  // メッセージ作成
  let messageBody = `${headComment}\n\n${incomeAndExpenditure}\n\n${feedBack}\n\n★今月の2人の振込金額\n直人：${row6Value.toLocaleString()}円\n沙羅：${row7Value.toLocaleString()}円`;

  return messageBody
}
