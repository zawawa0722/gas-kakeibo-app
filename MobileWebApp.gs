function doGet() {
  return HtmlService.createHtmlOutputFromFile("mobile/MobileDashboard")
    .setTitle("家計簿")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover");
}

// 実データはウェブアプリをデプロイした本人にだけ返す。
function getMobileDashboardMonth(month) {
  validateMobileDashboardRequest_(month);

  const spreadsheet = getSpreadsheet();
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const currentYear = Utilities.formatDate(new Date(), timezone, "yyyy");
  const rows = [];
  const addRow = (kind, category, detail, amount) => {
    if (amount === "" || amount === 0) return;
    if (!Number.isFinite(amount) || !["収入", "支出"].includes(kind)) {
      throw new Error("月別内訳の区分または金額を確認してください。");
    }
    rows.push({ kind, category, name: detail || category, amount: kind === "支出" ? -amount : amount });
  };

  if (month.startsWith(currentYear + "-")) {
    // 帳簿管理の収入総額・支出総額は8〜9行目にある。
    const ledger = readMigrationSource_(spreadsheet, SHEET_THOUBOKANRI, 9, timezone);
    if (!ledger.months.has(month)) return { rows, available: false };
    const income = readMigrationSource_(spreadsheet, SHEET_SYUNYUKANRI, 10, timezone);
    const expense = readMigrationSource_(spreadsheet, SHEET_SISYUTUKANRI, 28, timezone);
    const online = readMigrationSource_(spreadsheet, SHEET_ONLINE, 9, timezone);
    for (const label of income.labels.filter(label => label !== "収入総額")) {
      addRow("収入", label, "", income.amount(label, month));
    }
    for (const label of expense.labels.filter(label => !["支出総額", "オンライン"].includes(label))) {
      addRow("支出", label, "", expense.amount(label, month));
    }
    for (const label of online.labels) {
      addRow("支出", "オンライン", label, online.amount(label, month));
    }
    const total = kind => rows.filter(row => row.kind === kind)
      .reduce((sum, row) => sum + row.amount, 0);
    if (total("収入") !== (ledger.amount("収入総額", month) || 0) ||
        -total("支出") !== (ledger.amount("支出総額", month) || 0)) {
      throw new Error("帳簿管理の合計と月別内訳が一致しません。");
    }
    return { rows, available: true };
  }

  const archive = spreadsheet.getSheetByName("月次実績");
  if (!archive) throw new Error("月次実績シートが見つかりません。");
  const [headers, ...actuals] = archive.getDataRange().getValues();
  if (!["対象月", "区分", "カテゴリ", "内訳", "金額"].every((label, index) => headers[index] === label)) {
    throw new Error("月次実績シートの見出しを確認してください。");
  }
  let available = false;
  for (const [date, kind, category, detail, amount] of actuals) {
    if (normalizeMonthKey_(date, timezone, true) !== month) continue;
    available = true;
    addRow(kind, String(category), String(detail || ""), amount);
  }
  return { rows, available };
}

function validateMobileDashboardRequest_(month) {
  const activeEmail = Session.getActiveUser().getEmail();
  if (!activeEmail || activeEmail !== Session.getEffectiveUser().getEmail()) {
    throw new Error("家計簿の所有者アカウントで開いてください。");
  }
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) throw new Error("対象月が無効です。");
}

function getMobileDashboardTransactions(month) {
  validateMobileDashboardRequest_(month);
  const file = findOneBankCsv_(month.slice(0, 4));
  if (!file) return [];

  const rows = Utilities.parseCsv(file.getBlob().getDataAsString().replace(/^\uFEFF/, ""));
  const transactions = [];
  for (const row of rows) {
    if (row.length !== 8 || row[COL_VALID].trim() !== "1") continue;
    const date = row[0].trim().match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?(?:\s|$)/);
    if (!date || date[1] + "-" + date[2].padStart(2, "0") !== month) continue;
    const day = Number(date[3]);
    const daysInMonth = new Date(Date.UTC(Number(date[1]), Number(date[2]), 0)).getUTCDate();
    const name = row[COL_DESC].trim();
    const rawAmount = row[COL_AMOUNT].trim();
    const amount = Number(rawAmount);
    if (day < 1 || day > daysInMonth || !name || !rawAmount || !Number.isFinite(amount)) continue;

    const time = Number(date[4] || 0) * 3600 + Number(date[5] || 0) * 60 + Number(date[6] || 0);
    transactions.push({
      day, time, name, amount: -Math.abs(amount), kind: "支出", category: row[COL_CATEGORY].trim()
    });
  }
  transactions.sort((a, b) => b.day - a.day || b.time - a.time);
  return transactions.map(({ time, ...transaction }) => transaction);
}
