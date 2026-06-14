// CSVファイル配置依頼
function sendEmailRequestCsvStorage() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_CSV_STORAGE, BODY_REQUEST_CSV_STORAGE);
}

function compNotify(buttonFlug, emptyCheckFlug) {
  // メールを送信
  if (buttonFlug == false){
    if (emptyCheckFlug == false){
      GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT);
    }else{
      GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_FAILED_IMPORT, BODY_FAILED_IMPORT);
    }
  }
}

function sendEmailTwentySixDay() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_BUCKUP, BODY_REQUEST_BUCKUP);
}

// 振込依頼
function sendFamilyTransferRequest() {
  let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_THOUBOKANRI);
  let dateRange = CELL_POSITIONS[SHEET_THOUBOKANRI].dateRange;
  let ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');

  // 先月のyyyy/MM表記を生成 (ゼロ埋め対応)
  let today = new Date();
  let lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  let lastMonthText = `${lastMonth.getFullYear()}/${String(lastMonth.getMonth() + 1).padStart(2, "0")}`;
  let lastMonthTextNotZero = `${String(lastMonth.getMonth() + 1)}`

  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }

  // 6行目と7行目の値を取得
  let colIndex = lastMonthIndex + 2; // 列番号（1始まり）
  let row6Value = sheet.getRange(6, colIndex).getValue(); // 6行目の値
  let row7Value = sheet.getRange(7, colIndex).getValue(); // 7行目の値

  // 項目名を取得 (A8:A12)
  let labels = sheet.getRange("A8:A12").getValues().map(row => row[0]);

  // 8〜12行目の値を取得
  let values = sheet.getRange(8, lastMonthIndex + 2, 5, 1).getValues();

  // 表形式にするため、項目名と値を結びつける
  let resultTable = `以下が${lastMonthTextNotZero}月の結果です。\nお疲れ様でした!\n\n`;
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

  // 収入総額と収支結果を取得し、計算前に有効な数値かチェック
  let income = formattedValues["収入総額"] || 0; // 収入総額
  let balance = formattedValues["収支結果"] || 0; // 収支結果

  // 貯金率の計算 (収支結果 / 収入総額)
  let savingsRate = (income > 0) ? (balance / income) * 100 : 0;

  // 貯金率に基づくメッセージ分岐
  let savingsMessage = "";
  if (savingsRate >= 30) {
    savingsMessage = SAVINGS_MESSAGES.high;
  } else if (savingsRate >= 20) {
    savingsMessage = SAVINGS_MESSAGES.midHigh;
  } else if (savingsRate >= 10) {
    savingsMessage = SAVINGS_MESSAGES.midLow;
  } else if (savingsRate > 0 && savingsRate < 10) {
    savingsMessage = SAVINGS_MESSAGES.low;
  } else if (savingsRate <= 0) {
    savingsMessage = SAVINGS_MESSAGES.death;
  } else {
    savingsMessage = SAVINGS_MESSAGES.none;
  }

  // メッセージ作成
  let messageBody = resultTable + `\n\n${savingsMessage}\n\n★今月の2人の振込金額\n直人：${row6Value.toLocaleString()}円\n沙羅：${row7Value.toLocaleString()}円`;

  let options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify({
      "messages": [{
        "type": "text",
        "text": messageBody
      }]
    })
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", options);

}
// CSVファイル配置依頼
function sendEmailRequestCsvStorage() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_CSV_STORAGE, BODY_REQUEST_CSV_STORAGE);
}

function compNotify(buttonFlug, emptyCheckFlug) {
  // メールを送信
  if (buttonFlug == false){
    if (emptyCheckFlug == false){
      GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT);
    }else{
      GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_FAILED_IMPORT, BODY_FAILED_IMPORT);
    }
  }
}

function sendEmailTwentySixDay() {
  // メールを送信
  GmailApp.sendEmail(RECIPIENTS.naoto, SUBJECT_REQUEST_BUCKUP, BODY_REQUEST_BUCKUP);
}

// 振込依頼
function sendFamilyTransferRequest() {
  let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_THOUBOKANRI);
  let dateRange = CELL_POSITIONS[SHEET_THOUBOKANRI].dateRange;
  let ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty('LINE_ACCESS_TOKEN');

  // 先月のyyyy/MM表記を生成 (ゼロ埋め対応)
  let today = new Date();
  let lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  let lastMonthText = `${lastMonth.getFullYear()}/${String(lastMonth.getMonth() + 1).padStart(2, "0")}`;
  let lastMonthTextNotZero = `${String(lastMonth.getMonth() + 1)}`

  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = dateValues.indexOf(lastMonthText);

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }

  // 6行目と7行目の値を取得
  let colIndex = lastMonthIndex + 2; // 列番号（1始まり）
  let row6Value = sheet.getRange(6, colIndex).getValue(); // 6行目の値
  let row7Value = sheet.getRange(7, colIndex).getValue(); // 7行目の値

  // 項目名を取得 (A8:A12)
  let labels = sheet.getRange("A8:A12").getValues().map(row => row[0]);

  // 8〜12行目の値を取得
  let values = sheet.getRange(8, lastMonthIndex + 2, 5, 1).getValues();

  // 表形式にするため、項目名と値を結びつける
  let resultTable = `以下が${lastMonthTextNotZero}月の結果です。\nお疲れ様でした!\n\n`;
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

  // 収入総額と収支結果を取得し、計算前に有効な数値かチェック
  let income = formattedValues["収入総額"] || 0; // 収入総額
  let balance = formattedValues["収支結果"] || 0; // 収支結果

  // 貯金率の計算 (収支結果 / 収入総額)
  let savingsRate = (income > 0) ? (balance / income) * 100 : 0;

  // 貯金率に基づくメッセージ分岐
  let savingsMessage = "";
  if (savingsRate >= 30) {
    savingsMessage = SAVINGS_MESSAGES.high;
  } else if (savingsRate >= 20) {
    savingsMessage = SAVINGS_MESSAGES.midHigh;
  } else if (savingsRate >= 10) {
    savingsMessage = SAVINGS_MESSAGES.midLow;
  } else if (savingsRate > 0 && savingsRate < 10) {
    savingsMessage = SAVINGS_MESSAGES.low;
  } else if (savingsRate <= 0) {
    savingsMessage = SAVINGS_MESSAGES.death;
  } else {
    savingsMessage = SAVINGS_MESSAGES.none;
  }

  // メッセージ作成
  let messageBody = resultTable + `\n\n${savingsMessage}\n\n★今月の2人の振込金額\n直人：${row6Value.toLocaleString()}円\n沙羅：${row7Value.toLocaleString()}円`;

  let options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + ACCESS_TOKEN,
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify({
      "messages": [{
        "type": "text",
        "text": messageBody
      }]
    })
  };

  UrlFetchApp.fetch("https://api.line.me/v2/bot/message/broadcast", options);

}
