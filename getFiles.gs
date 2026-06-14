// 対象ファイルID取得関数（ボタンが押されたかどうかで処理を分岐）
function getFileId(isButtonPressed) {
  // シートの取得
  let spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_THOUBOKANRI);

  let targetFileName; // ファイル名を格納する変数

  // ボタンが押された場合、G15とH15セルの日付をそのまま使用
  if (isButtonPressed) {
    let targetDateget = sheet.getRange('G15').getValue();

    // G15から年を取得してファイル名を作成
    let targetFileYear = targetDateget;
    targetFileName = `B43支出_${targetFileYear}.csv`;

  } else { 
    // ボタンが押されなかった場合、1ヶ月前の日付を使用
    let today = new Date();
    let oneMonthAgo = new Date(today);
    oneMonthAgo.setMonth(today.getMonth() - 1);

    let targetFileYear = Utilities.formatDate(oneMonthAgo, 'JST', 'yyyy');
    targetFileName = `B43支出_${targetFileYear}.csv`;
  }

  // 対象ファイルIDを取得
  let { fileNameArray, fileObjects } = getFilenamesInFolder();
  let fileExistFlag = false;
  let foundFilePath = null;
  let fileId = null;

  // 取得ファイル名の中に対象ファイルがあるかチェック
  for (let i = 0; i < fileNameArray.length; i++) {
    if (fileNameArray[i] === targetFileName) {
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

// B43フォルダ内の全ファイル名とファイルオブジェクトを取得
function getFilenamesInFolder() {
  let folderName = "B43"; // 取得したいフォルダの名前
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

