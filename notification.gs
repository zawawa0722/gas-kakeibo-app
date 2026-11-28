let flagObj = { hasWarnOrError: false };

// CSVファイル配置依頼
function sendEmailRequestCsvStorage() {
  // メールを送信
  GmailApp.sendEmail(EMAIL_PRIMARY, SUBJECT_REQUEST_CSV_STORAGE, BODY_REQUEST_CSV_STORAGE);
}

// インポート結果通知
function compNotify(buttonFlug, emptyCheckFlug, hasWarnOrError) {
  // ボタンフラグと支出シートの空チェックフラグでインポート成否を判定
  if (buttonFlug == false){
    if (emptyCheckFlug == false){
        if (hasWarnOrError) {
          GmailApp.sendEmail(EMAIL_PRIMARY, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT_INCLUDE_ERROR);
        }else{
          GmailApp.sendEmail(EMAIL_PRIMARY, SUBJECT_SCCEED_IMPORT, BODY_SCCEED_IMPORT);
        }
    }else{
      GmailApp.sendEmail(EMAIL_PRIMARY, SUBJECT_FAILED_IMPORT, BODY_FAILED_IMPORT);
    }
  }
}

// バックアップリクエスト
function sendEmailRequestBuckup() {
  // メールを送信
  GmailApp.sendEmail(EMAIL_PRIMARY, SUBJECT_REQUEST_BUCKUP, BODY_REQUEST_BUCKUP);
}

// 振込依頼通知
function sendFamilyTransferRequest() {

  const flagObj = { hasWarnOrError: false };
  let message;
  try {
    message = createMessage(getSpreadsheet(), flagObj);
  } catch (error) {
    outputLog("ERROR", "月次LINE通知の本文を作成できませんでした。: " + error.message, false, flagObj);
    throw error;
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
  let lastMonthIndex = findMonthIndex_(dateValues, lastMonthText, sheet.getParent().getSpreadsheetTimeZone());

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }
  
  const labelRange = sheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].totalResultRange);
  const labels = labelRange.getValues();
  const values = sheet.getRange(labelRange.getRow(), lastMonthIndex + 2, labels.length, 1).getValues();
  return labels.map(([label], index) => {
    const value = values[index][0];
    if (!label || !Number.isFinite(value)) {
      throw new Error(lastMonthText + " の収支項目「" + (label || "項目名なし") + "」を確認してください。");
    }
    return label === "貯蓄率"
      ? `${label}: ${(value * 100).toFixed(2)}%`
      : `${label}: ${Math.trunc(value).toLocaleString()}円`;
  }).join("\n");
}

// 支出結果を取得
function getExpenses(sheet, dateRange, lastMonthText) {
  // dateRange から先月の列を特定
  let dateValues = sheet.getRange(dateRange).getValues()[0];
  let lastMonthIndex = findMonthIndex_(dateValues, lastMonthText, sheet.getParent().getSpreadsheetTimeZone());
  let resultTable = ""

  if (lastMonthIndex === -1) {
    Logger.log("先月のyyyy/MM表記が見つかりません。");
    return;
  }
  
  const labelRange = sheet.getRange(CELL_POSITIONS[SHEET_SISYUTUKANRI].labelRangeCustom);
  const labels = labelRange.getValues();
  const values = sheet.getRange(labelRange.getRow(), lastMonthIndex + 2, labels.length, 1).getValues();
  for (const [value] of values) {
    if (value !== "" && !Number.isFinite(value)) {
      throw new Error(lastMonthText + " の支出内訳に数値以外の金額があります。");
    }
  }
  return labels.flatMap(([label], index) => {
    const value = values[index][0];
    return label && Number.isFinite(value) ? [`${label}: ${Math.trunc(value).toLocaleString()}円`] : [];
  }).join("\n");
}

// LINE通知用メッセージ作成
function createMessage(spreadsheet = getSpreadsheet(), flagObj = { hasWarnOrError: false }) {
  const ledger = spreadsheet.getSheetByName(SHEET_THOUBOKANRI);
  if (!ledger) throw new Error("帳簿管理シートが見つかりません。");
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const { year, month } = getPreviousMonth_(timezone);
  const targetMonth = `${year}/${String(month).padStart(2, "0")}`;
  const ledgerRange = CELL_POSITIONS[SHEET_THOUBOKANRI].dateRange;
  const headers = ledger.getRange(ledgerRange).getValues()[0];
  const monthIndex = findMonthIndex_(headers, targetMonth, timezone);
  if (monthIndex === -1) throw new Error("帳簿管理に " + targetMonth + " の見出しがありません。");

  const balance = getResaltBalance(ledger, ledgerRange, targetMonth);
  if (!balance?.trim()) throw new Error(targetMonth + " の収支集計がありません。");

  const [[amountA], [amountB]] = ledger.getRange(6, monthIndex + 2, 2, 1).getValues();
  if (![amountA, amountB].every(Number.isFinite)) {
    throw new Error(targetMonth + " の振込金額を確認してください。");
  }

  const { USER_A_NAME, USER_B_NAME } = loadSheetSetting(spreadsheet).settingContents;
  if (![USER_A_NAME, USER_B_NAME].every(name => typeof name === "string" && name.trim())) {
    throw new Error("設定シートの振込先の表示名を確認してください。");
  }

  let spending = "";
  let missingReason = "支出内訳が空欄か、対象月がありません。";
  try {
    const expenseSheet = spreadsheet.getSheetByName(SHEET_SISYUTUKANRI);
    if (!expenseSheet) throw new Error("支出管理シートが見つかりません。");
    spending = getExpenses(expenseSheet, CELL_POSITIONS[SHEET_SISYUTUKANRI].dateRange, targetMonth);
  } catch (error) {
    missingReason = error.message;
  }

  let feedback = "AI分析コメントは取得できませんでした。";
  if (!spending?.trim()) {
    feedback = "集計データが不足しているため、AI分析を省略しました。";
    outputLog("WARN", targetMonth + " のAI分析を省略しました。: " + missingReason, false, flagObj);
  } else {
    try {
      feedback = analyzeHouseholdBudget(balance, spending);
    } catch (error) {
      outputLog("ERROR", targetMonth + " のAI分析コメントを取得できませんでした。: " + error.message, false, flagObj);
    }
  }

  return [
    `以下が${month}月の結果です。\nお疲れ様でした!`,
    balance,
    feedback,
    `★今月の振込金額\n${USER_A_NAME}：${amountA.toLocaleString()}円\n${USER_B_NAME}：${amountB.toLocaleString()}円`
  ].join("\n\n");
}

function sendMonthlyMigrationFailure_(monthKey, error) {
  GmailApp.sendEmail(EMAIL_PRIMARY, "【家計簿】月次転記に失敗",
    "対象月: " + (monthKey || "取得できませんでした") + "\n原因: " + error.message);
}

function sendMonthlyMigrationSuccess_(result) {
  GmailApp.sendEmail(EMAIL_PRIMARY, "【家計簿】月次転記が完了",
    ["対象月: " + result.targetMonth,
      "月次入力: " + result.monthlyInputRows + "行追加",
      "月次実績: " + result.monthlyActualRows + "行追加"].join("\n"));
}
