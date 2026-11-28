function findOneBankCsv_(year) {
  const folders = DriveApp.getFoldersByName("ワンバンク");
  if (!folders.hasNext()) return null;
  const filename = `ワンバンク支出_${year}.csv`.normalize("NFKC");
  const files = folders.next().getFiles();
  let match = null;
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().normalize("NFKC") !== filename) continue;
    if (match) throw new Error("同じ名前のワンバンクCSVが複数あります。");
    match = file;
  }
  return match;
}
