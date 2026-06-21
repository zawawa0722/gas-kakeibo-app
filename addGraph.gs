// 月選択をトリガーにグラフ出力関数を呼び出す
function onEdit(e) {
  let sheet = e.source.getActiveSheet();
  let range = e.range;
  let sheetName = sheet.getName();

  // 対象シートでなければ処理を中断
  if (!(sheetName in CELL_POSITIONS)) return;

  let { yearCell, monthCell, labelRange, dateRange, chartPosition } = CELL_POSITIONS[sheetName];

  // 編集されたセルが対象の月セルでない場合は処理を中断
  if (range.getA1Notation() !== monthCell) return;

  createDynamicPieChart(sheet, yearCell, monthCell, labelRange, dateRange, chartPosition);
}

// 動的円グラフを作成する関数
function createDynamicPieChart(sheet, yearCell, monthCell, labelRange, dateRange, chartPosition) {
  let year = sheet.getRange(yearCell).getValue();
  let month = sheet.getRange(monthCell).getValue();

  // 年と月が空白の場合は処理を中断
  if (!year || !month) {
    SpreadsheetApp.getUi().alert("年月の入力先セルの定義を確認してください");
    return
  };

  // 年月を生成 (月を文字列化して補正)
  let targetDate = `${year}/${month.toString().padStart(2, "0")}`;

  let headers = sheet.getRange(dateRange).getValues()[0]; // 月の年月 (B3:M3など)
  let columnIndex = headers.indexOf(targetDate) + 2; // 列インデックスを取得

  if (columnIndex === 1) {
    SpreadsheetApp.getUi().alert("指定された年月が見つかりません。");
    return;
  }

  let labelsRange = sheet.getRange(labelRange); // ラベルの範囲 (A4:A23など)
  let valuesRange = sheet.getRange(4, columnIndex, labelsRange.getNumRows(), 1); // 対象年月の列データ

  let labels = labelsRange.getValues().flat();
  let values = valuesRange.getValues().flat();

  let chartData = labels.map((label, i) => [label, values[i]])
                          .filter(row => row[1] && row[1] > 0);

  if (chartData.length === 0) {
    SpreadsheetApp.getUi().alert("有効なデータがありません。");
    return;
  }

  // グラフを作成・表示
  let chart = sheet.newChart()
    .setChartType(Charts.ChartType.PIE) // 円グラフを指定
    .addRange(labelsRange)
    .addRange(valuesRange)
    .setPosition(sheet.getRange(chartPosition).getRow(), sheet.getRange(chartPosition).getColumn(), 0, 0) // グラフ表示先を設定
    .setOption("title", `内訳 (${targetDate})`)
    .setOption("pieHole", 0.4) // ドーナツ型にするためのオプション
    .setOption("height", 395) // グラフの高さを指定 (ピクセル単位)
    .setOption("width", 705) // グラフの幅を指定 (ピクセル単位)
    .build();

  // 既存のグラフを削除
  let charts = sheet.getCharts();
  charts.forEach(existingChart => sheet.removeChart(existingChart));

  sheet.insertChart(chart);
}
