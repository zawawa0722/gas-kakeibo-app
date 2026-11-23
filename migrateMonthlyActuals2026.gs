// Apps Script エディタから実行する。プレビューは件数だけをログに出す。
function previewMonthlyActuals2026() {
  return migrateMonthlyActuals2026_(false);
}

function migrateMonthlyActuals2026() {
  return migrateMonthlyActuals2026_(true);
}

function migrateMonthlyActuals2026_(write) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = getSpreadsheet();
    const timezone = spreadsheet.getSpreadsheetTimeZone();
    const year = 2026;
    const [currentYear, currentMonth] = Utilities.formatDate(new Date(), timezone, "yyyy-MM")
      .split("-").map(Number);
    const lastMonth = currentYear < year ? 0 : currentYear === year ? currentMonth : 12;
    const settings = loadSheetSetting(spreadsheet).settingContents;
    const users = [
      ["USER_A", settings.USER_A_NAME],
      ["USER_B", settings.USER_B_NAME]
    ];
    if (users.some(([, name]) => typeof name !== "string" || !name.trim()) ||
        users[0][1] === users[1][1]) {
      throw new Error("設定シートのユーザー名を確認してください。");
    }

    const ledger = readMigrationSource2026_(spreadsheet, SHEET_THOUBOKANRI, 7);
    const income = readMigrationSource2026_(spreadsheet, SHEET_SYUNYUKANRI, 10);
    const expenses = readMigrationSource2026_(spreadsheet, SHEET_SISYUTUKANRI, 28);
    const online = readMigrationSource2026_(spreadsheet, SHEET_ONLINE, 9);
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

    const input = planMigrationAppend2026_(spreadsheet, "月次入力",
      ["対象月", "ユーザーID", "仮入金額", "実入金額"], 2, inputRows, timezone);
    const actuals = planMigrationAppend2026_(spreadsheet, "月次実績",
      ["対象月", "区分", "カテゴリ", "内訳", "金額"], 4, actualRows, timezone);
    if (write) {
      appendMigrationRows2026_(input);
      appendMigrationRows2026_(actuals);
      SpreadsheetApp.flush();
    }

    const summary = {
      year: year,
      lastMonth: lastMonth,
      preview: !write,
      monthlyInputRows: input.rows.length,
      monthlyActualRows: actuals.rows.length
    };
    Logger.log(JSON.stringify(summary));
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function readMigrationSource2026_(spreadsheet, name, lastRow) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error(name + "シートが見つかりません。");
  const range = sheet.getRange(3, 1, lastRow - 2, 13);
  const values = range.getValues();
  const months = new Map(values[0].slice(1).map((value, index) =>
    [migrationMonthKey2026_(value, spreadsheet.getSpreadsheetTimeZone()), index + 1]));
  if (months.size !== 12) throw new Error(name + "の年月見出しが重複しています。");

  const rows = new Map();
  values.slice(1).forEach(row => {
    const label = String(row[0]).trim();
    if (!label || rows.has(label)) throw new Error(name + "の項目名を確認してください。");
    rows.set(label, row);
  });
  return {
    labels: Array.from(rows.keys()),
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

function migrationMonthKey2026_(value, timezone) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, "yyyy-MM");
  const match = String(value).trim().match(/^(\d{4})[/-](0?[1-9]|1[0-2])$/);
  if (!match) throw new Error("対象月の形式を確認してください。");
  return match[1] + "-" + match[2].padStart(2, "0");
}

function planMigrationAppend2026_(spreadsheet, name, headers, keyColumns, candidates, timezone) {
  const sheet = spreadsheet.getSheetByName(name);
  if (!sheet) throw new Error(name + "シートが見つかりません。");
  const lastRow = sheet.getLastRow();
  const values = sheet.getRange(1, 1, Math.max(lastRow, 1), headers.length).getValues();
  if (headers.some((header, index) => values[0][index] !== header)) {
    throw new Error(name + "の見出しを確認してください。");
  }
  const keyFor = row => JSON.stringify([
    migrationMonthKey2026_(row[0], timezone), ...row.slice(1, keyColumns)
  ]);
  const existing = new Map();
  values.slice(1).forEach(row => {
    if (row.every(value => value === "")) return;
    const month = migrationMonthKey2026_(row[0], timezone);
    if (!month.startsWith("2026-")) return;
    const key = keyFor(row);
    if (existing.has(key)) throw new Error(name + "に2026年の重複行があります。 " + key);
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

function appendMigrationRows2026_(plan) {
  if (!plan.rows.length) return;
  const needed = plan.startRow + plan.rows.length - 1 - plan.sheet.getMaxRows();
  if (needed > 0) plan.sheet.insertRowsAfter(plan.sheet.getMaxRows(), needed);
  plan.sheet.getRange(plan.startRow, 1, plan.rows.length, 1).setNumberFormat("yyyy/mm");
  plan.sheet.getRange(plan.startRow, 1, plan.rows.length, plan.width).setValues(plan.rows);
}
