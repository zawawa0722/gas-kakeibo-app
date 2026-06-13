function extractTargetLine(fileId, imButtonPressed) {
  
  // 日付文字列の作成
  let targetDate;
  if (imButtonPressed) {
    // yearCell と monthCell の値を取得
    let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = spreadsheet.getSheetByName(SHEET_THOUBOKANRI);
    let year = sheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].yearCell).getValue();
    let month = sheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].monthCell).getValue();
    targetDate = `${year}/${String(month).padStart(2, '0')}`;
  } else {
    // imButtonPressed が FALSE の場合、先月の日付文字列を作成
    let now = new Date();
    now.setMonth(now.getMonth() - 1); // 先月に設定
    let year = now.getFullYear();
    let month = now.getMonth() + 1; // 月は0始まりのため+1
    targetDate = `${year}/${String(month).padStart(2, '0')}`;
  }

  // ファイルを取得し内容を読み取る
  let file = DriveApp.getFileById(fileId);
  if (!file) {
    throw new Error("指定されたファイルが見つかりません。");
  }

  let content = file.getBlob().getDataAsString(); // CSVファイルの内容を文字列として取得
  let lines = content.split("\n"); // 各行を分割

  let filteredArray = []; // 条件を満たす行を格納する配列

  // 1行ずつ処理
  for (let i = 0; i < lines.length; i++) {
    let line = parseCSVLine(lines[i]); // カスタムパーサーで解析

    // 必要条件を満たす行を格納
    if (
      line.length === 8 && // 要素数が8であることを確認
      line[0] && line[0].includes(targetDate) && // 日付文字列が含まれる
      line[6] === "1" // インデックス6が "1" であること
    ) {
      filteredArray.push(line);
    }
  }

  return filteredArray; // フィルタリングされた配列を返す
}

// CSVの1行を正確に解析する関数
function parseCSVLine(line) {
  let regex = /("([^"]|(""))*"|[^,]+|(?<=,)(?=,))/g; // CSVのルールに従って分割
  let matches = [...line.matchAll(regex)]; // 正規表現で全てのマッチを取得
  return matches.map(match => {
    let value = match[0];
    if (value.startsWith('"') && value.endsWith('"')) {
      value = value.slice(1, -1).replace(/""/g, '"'); // ダブルクオートを解除
    }
    return value.trim(); // 前後の空白を削除
  });
}

function checkSheetsDate(imButtonPressed) {

  let result = []; // 結果を格納する配列
  let today = new Date();
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // imButtonPressed に基づいてターゲット日付を作成
  let targetDate;
  if (imButtonPressed) {
    // G15 と H15 から年と月を取得して日付を生成 (帳簿管理を基準)
    let year = spreadsheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].yearCell).getValue();
    let month = spreadsheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].monthCell).getValue();
    month = month < 10 ? `0${month}` : month; // 前ゼロを付与
    targetDate = `${year}/${month}`;
  } else {
    // 本日の先月の日付を計算して yyyy/MM 形式で取得
    let lastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    targetDate = Utilities.formatDate(lastMonth, Session.getScriptTimeZone(), "yyyy/MM");
  }

  // 対象シートをループ処理
  [SHEET_THOUBOKANRI, SHEET_SISYUTUKANRI, SHEET_ONLINE].forEach(sheetName => {
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      console.warn(`シート "${sheetName}" が見つかりませんでした。`);
      return; // または throw new Error(...)
    }

    let positions = CELL_POSITIONS[sheetName];
    let dateRange = sheet.getRange(positions.dateRange).getValues()[0]; // 日付行を取得

    // ターゲット日付と一致する列を探索
    dateRange.forEach((date, columnIndex) => {
      if (date === targetDate) {
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
