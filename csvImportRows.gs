function extractTargetLine(fileId, imButtonPressed) {

  // 日付文字列の作成
  let targetDate;
  let SPREADSHEET = getSpreadsheet();
  if (imButtonPressed) {
    // yearCell と monthCell の値を取得
    let tyoubosheet = SPREADSHEET.getSheetByName(SHEET_THOUBOKANRI);
    let year = tyoubosheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].yearCell).getValue();
    let month = tyoubosheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].monthCell).getValue();
    targetDate = `${year}/${String(month).padStart(2, '0')}`;
  } else {
    // imButtonPressed が FALSE の場合、先月の日付文字列を作成
    const { year, month } = getPreviousMonth_(SPREADSHEET.getSpreadsheetTimeZone());
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
