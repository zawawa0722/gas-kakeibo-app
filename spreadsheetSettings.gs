// スプレッドシート取得
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 設定シートの値を取得
function loadSheetSetting(SPREADSHEET) {
  let settingsheet = SPREADSHEET.getSheetByName(SHEET_SETTING);
  let values = settingsheet.getDataRange().getValues();

  let settingContents = {};
  let shopListRakuten = [];
  let shopListOther = [];

  // ヘッダー（1行目）を除外して処理
  values.slice(1).forEach(row => {
    // A・B列: 設定値 (row[0]=A列, row[1]=B列)
    if (row[0]) {
      settingContents[row[0]] = row[1];
    }
    // D列: 楽天市場オンラインショップ名 (row[3]=D列)
    if (row[3]) {
      shopListRakuten.push(row[3]);
    }
    // E列: その他オンラインショップ名 (row[4]=E列)
    if (row[4]) {
      shopListOther.push(row[4]);
    }
  });

  return {
    settingContents: settingContents,
    shopListRakuten: shopListRakuten,
    shopListOther: shopListOther
  };
}

// プログラム内部で保持するユーザ名に設定シートの情報を反映
function updateExpendCategoryNames() {
  // 設定シートの値を取り出す
  const nameA = SHEET_SETTING_CONTENTS.USER_A_NAME;
  const nameB = SHEET_SETTING_CONTENTS.USER_B_NAME;

  // 空チェック（どちらか一方でも空ならエラーを投げる）
  if (!nameA || !nameB) {
    throw new Error("設定シートの『USER_A_NAME』または『USER_B_NAME』が空欄です。設定を確認してください。");
  }

  // 安全が確認できたら上書き
  EXPEND_CATEGORY.TEMP_USER_A = nameA + "一時負担";
  EXPEND_CATEGORY.TEMP_USER_B = nameB + "一時負担";
}
