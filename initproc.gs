// ボタンフラグ取得
function onButtonClick() {
  // 実行前に確認メッセージを表示
  let result = Browser.msgBox("確認", "家計簿出力処理を実行してもよろしいですか？", Browser.Buttons.YES_NO);

  // YESボタンが押された場合のみスクリプトを実行
  if (result == "yes") {
    let buttonPressed = true;
    importCsv(buttonPressed);
  } else {
    // NOボタンが押された場合は実行しない
    Logger.log("家計簿出力処理は実行されませんでした");
  }
}

// スプレッドシート取得
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// 対象ファイルID取得関数（ボタンが押されたかどうかで処理を分岐）
function getFileId(isButtonPressed) {
  // シートの取得
  let SPREADSHEET = getSpreadsheet();
  let tyoubosheet = SPREADSHEET.getSheetByName(SHEET_THOUBOKANRI);
  let targetDateget;
  let targetFileName;
  let targetFileYear;
  let oneMonthAgo;
  let TODAY = new Date();

  // ボタンが押された場合、G15とH15セルの日付をそのまま使用
  if (isButtonPressed) {
    targetDateget = tyoubosheet.getRange(CELL_POSITIONS[SHEET_THOUBOKANRI].yearCell).getValue();
    // G15から年を取得してファイル名を作成
    targetFileYear = targetDateget;
    targetFileName = `ワンバンク支出_${targetFileYear}.csv`;
  } else { 
    // ボタンが押されなかった場合、1ヶ月前の日付を使用
    oneMonthAgo = new Date(TODAY);
    oneMonthAgo.setMonth(TODAY.getMonth() - 1);
    targetFileYear = Utilities.formatDate(oneMonthAgo, 'JST', 'yyyy');
    targetFileName = `ワンバンク支出_${targetFileYear}.csv`;
  }

  // 対象ファイルIDを取得
  let { fileNameArray, fileObjects } = getFilenamesInFolder();
  let fileExistFlag = false;
  let foundFilePath = null;
  let fileId = null;

  // 取得ファイル名の中に対象ファイルがあるかチェック
  for (let i = 0; i < fileNameArray.length; i++) {
    if (normalizeString(fileNameArray[i]) === normalizeString(targetFileName)) {  
      fileExistFlag = true;
      // ファイルのURLを取得
      foundFilePath = fileObjects[i].getUrl();
      // URL から ID を抽出
      fileId = extractFileId(foundFilePath);
      break;
    }
  }

  // ファイルIDをリターン
  if (fileExistFlag === true && fileId !== null) {
    return fileId;
  } else {
    throw new Error("指定されたCSVファイルが見つかりませんでした");
  }
}


// Unicodeの正規化を実施
function normalizeString(str) {
  return str.normalize('NFKC');
}

// ワンバンクフォルダ内の全ファイル名とファイルオブジェクトを取得
function getFilenamesInFolder() {
  let folderName = "ワンバンク"; // 取得したいフォルダの名前
  let folder = DriveApp.getFoldersByName(folderName).next();
  let files = folder.getFiles();
  let filenames = [];
  let fileObjects = [];

  while (files.hasNext()) {
    let file = files.next();
    filenames.push(file.getName());
    fileObjects.push(file);
  }

  return { fileNameArray: filenames, fileObjects: fileObjects };
}

// URLからファイルIDを抽出
function extractFileId(fileUrl) {
  let idPattern = /\/d\/([a-zA-Z0-9_-]+)/;
  let match = fileUrl.match(idPattern);
  return match ? match[1] : null;
}


