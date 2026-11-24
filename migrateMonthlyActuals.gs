function onOpen() {
  SpreadsheetApp.getUi().createMenu("月次転記")
    .addItem("プレビューして転記", "confirmMonthlyMigration")
    .addItem("候補年を更新", "refreshMonthlyMigrationYears")
    .addToUi();
}

function refreshMonthlyMigrationYears() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const [year, month] = Utilities.formatDate(new Date(), timezone, "yyyy-MM").split("-").map(Number);
  const years = monthlyMigrationYears_(readMonthlyMigrationSources_(spreadsheet, timezone), year, month);
  const selector = getMonthlyMigrationYearSelector_(spreadsheet);
  if (!years.length) {
    selector.clearDataValidations().clearContent();
    SpreadsheetApp.getUi().alert("転記元4シートに対象月の揃った年がありません。");
    return;
  }
  selector.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(years.map(String), true).setAllowInvalid(false).build());
  if (!years.includes(Number(selector.getValue()))) selector.setValue(years[0]);
}

function confirmMonthlyMigration() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  try {
    const year = getMonthlyMigrationYear_(spreadsheet);
    const preview = migrateMonthlyActuals_(spreadsheet, year);
    const message = [
      year + "年1月〜" + preview.lastMonth + "月分",
      "月次入力 " + preview.monthlyInputRows + "行／月次実績 " + preview.monthlyActualRows + "行",
      ...(preview.includesCurrentMonth ? ["当月分には現時点の値が含まれます。"] : [])
    ].join("\n");
    if (!preview.monthlyInputRows && !preview.monthlyActualRows) {
      ui.alert("月次転記", message + "\n追加する行はありません。", ui.ButtonSet.OK);
      return;
    }
    if (ui.alert("月次転記の確認", message + "\n転記しますか？", ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
    migrateMonthlyActuals_(spreadsheet, year, preview);
    ui.alert("月次転記が完了しました。", message, ui.ButtonSet.OK);
  } catch (error) {
    ui.alert("月次転記", error.message, ui.ButtonSet.OK);
  }
}

function readMonthlyMigrationSources_(spreadsheet, timezone) {
  return [
    readMigrationSource_(spreadsheet, SHEET_THOUBOKANRI, 7, timezone),
    readMigrationSource_(spreadsheet, SHEET_SYUNYUKANRI, 10, timezone),
    readMigrationSource_(spreadsheet, SHEET_SISYUTUKANRI, 28, timezone),
    readMigrationSource_(spreadsheet, SHEET_ONLINE, 9, timezone)
  ];
}

function monthlyMigrationYears_(sources, currentYear, currentMonth) {
  const years = new Set(Array.from(sources[0].months.keys(), key => Number(key.slice(0, 4))));
  return Array.from(years).filter(year => {
    if (year > currentYear) return false;
    const lastMonth = year === currentYear ? currentMonth : 12;
    const months = Array.from({ length: lastMonth }, (_, index) =>
      year + "-" + String(index + 1).padStart(2, "0"));
    return sources.every(source => months.every(month => source.months.has(month)) &&
      months.some(month => source.hasData(month)));
  }).sort((a, b) => b - a);
}

function getMonthlyMigrationYearSelector_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SHEET_THOUBOKANRI);
  if (!sheet) throw new Error(SHEET_THOUBOKANRI + "シートが見つかりません。");
  return sheet.getRange(MONTHLY_MIGRATION_YEAR_CELL);
}

function getMonthlyMigrationYear_(spreadsheet) {
  const value = getMonthlyMigrationYearSelector_(spreadsheet).getValue();
  if ((typeof value !== "number" && typeof value !== "string") ||
      !/^\d{4}$/.test(String(value).trim()) || Number(value) < 1000) {
    throw new Error("帳簿管理の「転記対象年」を選択してください。");
  }
  return Number(value);
}

function previewMonthlyActuals() {
  const spreadsheet = getSpreadsheet();
  return migrateMonthlyActuals_(spreadsheet, getMonthlyMigrationYear_(spreadsheet));
}

function migrateMonthlyActuals() {
  const spreadsheet = getSpreadsheet();
  const year = getMonthlyMigrationYear_(spreadsheet);
  const preview = migrateMonthlyActuals_(spreadsheet, year);
  return migrateMonthlyActuals_(spreadsheet, year, preview);
}

function migrateMonthlyActuals_(spreadsheet, year, approvedPreview) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const timezone = spreadsheet.getSpreadsheetTimeZone();
    const [currentYear, currentMonth] = Utilities.formatDate(new Date(), timezone, "yyyy-MM")
      .split("-").map(Number);
    const lastMonth = currentYear === year ? currentMonth : 12;
    const [ledger, income, expenses, online] = readMonthlyMigrationSources_(spreadsheet, timezone);
    if (!monthlyMigrationYears_([ledger, income, expenses, online], currentYear, currentMonth).includes(year)) {
      throw new Error("選択年の対象月が転記元の4シートに揃っていません。候補年を更新してください。");
    }
    const settings = loadSheetSetting(spreadsheet).settingContents;
    const users = [
      ["USER_A", settings.USER_A_NAME],
      ["USER_B", settings.USER_B_NAME]
    ];
    if (users.some(([, name]) => typeof name !== "string" || !name.trim()) ||
        users[0][1] === users[1][1]) {
      throw new Error("設定シートのユーザー名を確認してください。");
    }

    const inputRows = [];
    const actualRows = [];
    for (let month = 1; month <= lastMonth; month++) {
      const monthKey = year + "-" + String(month).padStart(2, "0");
      const date = Utilities.parseDate(monthKey + "-01", timezone, "yyyy-MM-dd");
      for (const [userId, name] of users) {
        const provisional = ledger.amount(name + "仮入金額", monthKey);
        const actual = ledger.amount(name + "実入金額", monthKey);
        if (provisional !== "" || actual !== "") {
          inputRows.push([date, userId, provisional, actual]);
        }
      }

      let incomeSum = 0;
      let expenseSum = 0;
      let onlineSum = 0;
      for (const category of income.labels.filter(label => label !== "収入総額")) {
        const amount = income.amount(category, monthKey);
        if (amount === "") continue;
        actualRows.push([date, "収入", category, "", amount]);
        incomeSum += amount;
      }
      for (const category of expenses.labels.filter(label => !["オンライン", "支出総額"].includes(label))) {
        const amount = expenses.amount(category, monthKey);
        if (amount === "") continue;
        actualRows.push([date, "支出", category, "", amount]);
        expenseSum += amount;
      }
      for (const store of online.labels) {
        const amount = online.amount(store, monthKey);
        if (amount === "") continue;
        actualRows.push([date, "支出", "オンライン", store, amount]);
        onlineSum += amount;
      }
      if (incomeSum !== (income.amount("収入総額", monthKey) || 0) ||
          onlineSum !== (expenses.amount("オンライン", monthKey) || 0) ||
          expenseSum + onlineSum !== (expenses.amount("支出総額", monthKey) || 0)) {
        throw new Error(monthKey + " の内訳と合計が一致しません。");
      }
    }

    const input = planMigrationAppend_(spreadsheet, "月次入力",
      ["対象月", "ユーザーID", "仮入金額", "実入金額"], 2, inputRows, timezone, year);
    const actuals = planMigrationAppend_(spreadsheet, "月次実績",
      ["対象月", "区分", "カテゴリ", "内訳", "金額"], 4, actualRows, timezone, year);
    const summary = {
      year,
      lastMonth,
      includesCurrentMonth: currentYear === year,
      monthlyInputRows: input.rows.length,
      monthlyActualRows: actuals.rows.length,
      digest: Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
        JSON.stringify([input.rows, actuals.rows])).join(",")
    };
    if (approvedPreview) {
      if (JSON.stringify(summary) !== JSON.stringify(approvedPreview)) {
        throw new Error("転記予定が確認時から変わりました。もう一度プレビューしてください。");
      }
      appendMigrationRows_(input);
      appendMigrationRows_(actuals);
      SpreadsheetApp.flush();
    }
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function readMigrationSource_(spreadsheet, name, lastRow, timezone) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error(name + "シートが見つかりません。");
  const range = sheet.getRange(3, 1, lastRow - 2, sheet.getLastColumn());
  const values = range.getValues();
  const months = new Map();
  const duplicates = new Set();
  values[0].slice(1).forEach((value, index) => {
    const key = migrationMonthKey_(value, timezone, true);
    if (!key) return;
    if (months.has(key)) duplicates.add(key);
    months.set(key, index + 1);
  });
  duplicates.forEach(key => months.delete(key));

  const rows = new Map();
  values.slice(1).forEach(row => {
    const label = String(row[0]).trim();
    if (!label || rows.has(label)) throw new Error(name + "の項目名を確認してください。");
    rows.set(label, row);
  });
  return {
    labels: Array.from(rows.keys()),
    months,
    hasData(month) {
      return Array.from(rows.values()).some(row => Number.isFinite(row[months.get(month)]));
    },
    amount(label, month) {
      if (!months.has(month) || !rows.has(label)) {
        throw new Error(name + "の年月または項目が見つかりません。 " + month + " " + label);
      }
      const value = rows.get(label)[months.get(month)];
      if (value !== "" && !Number.isFinite(value)) {
        throw new Error(name + "の金額を確認してください。 " + month + " " + label);
      }
      return value;
    }
  };
}

function migrationMonthKey_(value, timezone, optional = false) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, "yyyy-MM");
  const match = String(value).trim().match(/^(\d{4})[/-](0?[1-9]|1[0-2])$/);
  if (!match && optional) return null;
  if (!match) throw new Error("対象月の形式を確認してください。");
  return match[1] + "-" + match[2].padStart(2, "0");
}

function planMigrationAppend_(spreadsheet, name, headers, keyColumns, candidates, timezone, year) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error(name + "シートが見つかりません。");
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, Math.max(lastRow, 1), headers.length).getValues();
  if (headers.some((header, index) => values[0][index] !== header)) {
    throw new Error(name + "の見出しを確認してください。");
  }
  const keyFor = row => JSON.stringify([
    migrationMonthKey_(row[0], timezone), ...row.slice(1, keyColumns)
  ]);
  const existing = new Map();
  values.slice(1).forEach(row => {
    if (row.every(value => value === "")) return;
    const month = migrationMonthKey_(row[0], timezone);
    if (!month.startsWith(year + "-")) return;
    const key = keyFor(row);
    if (existing.has(key)) throw new Error(name + "に" + year + "年の重複行があります。 " + key);
    existing.set(key, row);
  });

  const rows = candidates.filter(row => {
    const key = keyFor(row);
    const prior = existing.get(key);
    if (!prior) return true;
    if (row.slice(keyColumns).some((value, index) => value !== prior[keyColumns + index])) {
      throw new Error(name + "の既存行と転記元の金額が異なります。 " + key);
    }
    return false;
  });
  return { sheet: sheet, startRow: lastRow + 1, rows: rows, width: headers.length };
}

function appendMigrationRows_(plan) {
  if (!plan.rows.length) return;
  const needed = plan.startRow + plan.rows.length - 1 - plan.sheet.getMaxRows();
  if (needed > 0) plan.sheet.insertRowsAfter(plan.sheet.getMaxRows(), needed);
  plan.sheet.getRange(plan.startRow, 1, plan.rows.length, 1).setNumberFormat("yyyy/mm");
  plan.sheet.getRange(plan.startRow, 1, plan.rows.length, plan.width).setValues(plan.rows);
}
