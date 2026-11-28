const MOBILE_FIXED_EXPENSE_CATEGORIES = ["家賃", "家賃代", "住居", "住居費", "住宅費",
  "通信費", "通信料金", "分割費", "分割払い", "分割支払", "分割支払い",
  "ローン", "ローン返済", "ローン返済額", "住宅ローン"];

function doGet() {
  return HtmlService.createHtmlOutputFromFile("mobile/MobileDashboard")
    .setTitle("家計簿")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover");
}

// 実データはウェブアプリをデプロイした本人にだけ返す。
function getMobileDashboardMonth(month) {
  validateMobileDashboardRequest_(month);

  return readMobileDashboardMonth_(getSpreadsheet(), month);
}

function readMobileDashboardMonth_(spreadsheet, month, sources = new Map()) {
  const source = (name, lastRow) => {
    if (!sources.has(name)) sources.set(name, readMigrationSource_(spreadsheet, name, lastRow, timezone));
    return sources.get(name);
  };
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const currentYear = Utilities.formatDate(new Date(), timezone, "yyyy");
  const rows = [];
  const addRow = (kind, category, detail, amount) => {
    if (amount === "" || amount === 0) return;
    if (!Number.isFinite(amount) || !["収入", "支出"].includes(kind)) {
      throw new Error("月別内訳の区分または金額を確認してください。");
    }
    rows.push({ kind, category, name: detail || category, fixedExpense: kind === "支出" &&
      isMobileFixedExpense_(category), amount: kind === "支出" ? -amount : amount });
  };

  if (month.startsWith(currentYear + "-")) {
    // 帳簿管理の収入総額・支出総額は8〜9行目にある。
    const ledger = source(SHEET_THOUBOKANRI, 9);
    if (!ledger.months.has(month)) return { rows, available: false };
    const income = source(SHEET_SYUNYUKANRI, 10);
    const expense = source(SHEET_SISYUTUKANRI, 28);
    const online = source(SHEET_ONLINE, 9);
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
  if (!archive || archive.getLastRow() === 0) return { rows, available: false };
  if (!sources.has("月次実績")) sources.set("月次実績", archive.getDataRange().getValues());
  const [headers, ...actuals] = sources.get("月次実績");
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

function requireMobileOwner_() {
  const activeEmail = Session.getActiveUser().getEmail();
  if (!activeEmail || activeEmail !== Session.getEffectiveUser().getEmail()) {
    throw new Error("家計簿の所有者アカウントで開いてください。");
  }
}

function validateMobileDashboardRequest_(month) {
  requireMobileOwner_();
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

function getMobileDashboardOverview(month) {
  validateMobileDashboardRequest_(month);
  const spreadsheet = getSpreadsheet();
  const timezone = spreadsheet.getSpreadsheetTimeZone();
  const today = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd");
  const currentMonth = today.slice(0, 7);
  const sources = new Map();
  const data = readMobileDashboardMonth_(spreadsheet, month, sources);
  const previousMonth = shiftMobileMonth_(month, -1);
  const actuals = Array.from({ length: 6 }, (_, index) => {
    const key = shiftMobileMonth_(month, index - 5);
    try {
      const result = key === month ? data : readMobileDashboardMonth_(spreadsheet, key, sources);
      return {
        month: key,
        balance: key <= currentMonth && result.available
          ? result.rows.reduce((total, row) => total + row.amount, 0) : null,
        partial: key === currentMonth
      };
    } catch (error) {
      console.error(error);
      return { month: key, balance: null, partial: false };
    }
  });

  let previousRows = null;
  let budget = null;
  let installments = null;
  const errors = {};
  try {
    if (month < currentMonth && actuals[4].balance !== null) {
      previousRows = readMobileDashboardMonth_(spreadsheet, previousMonth, sources).rows;
    }
  } catch (error) {
    console.error(error);
  }
  try {
    if (data.available) budget = readMobileBudget_(spreadsheet, month, data.rows, timezone);
  } catch (error) {
    console.error(error);
    errors.budget = "変動費予算の設定を確認してください。";
  }
  try {
    installments = readMobileInstallments_(spreadsheet, today, timezone);
  } catch (error) {
    console.error(error);
    errors.installments = "分割払い管理シートを確認してください。";
  }
  return { ...data, actuals, previousRows, budget, installments, errors };
}

function isMobileFixedExpense_(category) {
  return MOBILE_FIXED_EXPENSE_CATEGORIES.includes(String(category).normalize("NFKC").replace(/\s/g, ""));
}

function shiftMobileMonth_(key, offset) {
  const [year, month] = key.split("-").map(Number);
  const index = year * 12 + month - 1 + offset;
  return String(Math.floor(index / 12)).padStart(4, "0") + "-" +
    String(index % 12 + 1).padStart(2, "0");
}

function readMobileBudget_(spreadsheet, month, rows, timezone) {
  const plan = spreadsheet.getSheetByName("変動費予算");
  let plannedIncome = null;
  if (plan && plan.getLastRow() > 0) {
    const [headers, ...entries] = plan.getRange(1, 1, plan.getLastRow(), 2).getValues();
    if (headers[0] !== "対象月" || headers[1] !== "手取り収入見込み") {
      throw new Error("変動費予算シートの見出しを確認してください。");
    }
    for (const [date, amount] of entries) {
      if (normalizeMonthKey_(date, timezone, true) !== month || amount === "") continue;
      if (!Number.isFinite(amount) || amount < 0 || plannedIncome !== null) {
        throw new Error("手取り収入見込みの金額または年月の重複を確認してください。");
      }
      plannedIncome = amount;
    }
  }
  const sum = entries => entries.reduce((total, row) => total + row.amount, 0);
  const income = plannedIncome ?? sum(rows.filter(row => row.kind === "収入"));
  if (plannedIncome === null && income <= 0) return null;

  const expenses = rows.filter(row => row.kind === "支出");
  const fixed = -sum(expenses.filter(row => isMobileFixedExpense_(row.category)));
  const used = -sum(expenses.filter(row => !isMobileFixedExpense_(row.category)));
  return { limit: income - fixed, used, planned: plannedIncome !== null };
}

function readMobileInstallments_(spreadsheet, today, timezone) {
  const sheet = spreadsheet.getSheetByName("分割払い管理");
  if (!sheet) return null;
  const [headers, ...entries] = sheet.getDataRange().getValues();
  const columns = [
    ["商品名"], ["支払総額", "支払い総額"], ["分割回数"], ["初回引落日", "初回引き落とし日", "初回支払日"]
  ].map(aliases => headers.findIndex(value => aliases.includes(String(value).trim())));
  if (columns.includes(-1)) throw new Error("分割払い管理の見出しを確認してください。");

  const currentMonth = today.slice(0, 7);
  const nextMonth = shiftMobileMonth_(currentMonth, 1);
  const daysInCurrentMonth = new Date(Date.UTC(Number(today.slice(0, 4)), Number(today.slice(5, 7)), 0)).getUTCDate();
  const items = [];
  for (const row of entries) {
    const [name, total, count, firstDate] = columns.map(index => row[index]);
    if ([name, total, count, firstDate].every(value => value === "")) continue;
    if (!name || !Number.isInteger(total) || total < 0 || !Number.isInteger(count) || count < 1 ||
        !(firstDate instanceof Date) || !Number.isFinite(firstDate.getTime())) {
      throw new Error("分割払い管理の入力内容を確認してください。");
    }
    const firstDateKey = Utilities.formatDate(firstDate, timezone, "yyyy-MM-dd");
    const firstMonth = firstDateKey.slice(0, 7);
    const offset = key => (Number(key.slice(0, 4)) - Number(firstMonth.slice(0, 4))) * 12 +
      Number(key.slice(5, 7)) - Number(firstMonth.slice(5, 7));
    const dueDay = Math.min(Number(firstDateKey.slice(8, 10)), daysInCurrentMonth);
    const paid = Math.max(0, Math.min(count, offset(currentMonth) + (Number(today.slice(8, 10)) >= dueDay ? 1 : 0)));
    if (paid === count) continue;
    const regular = Math.floor(total / count);
    const nextIndex = offset(nextMonth);
    const next = nextIndex < 0 || nextIndex >= count ? 0 :
      nextIndex === count - 1 ? total - regular * (count - 1) : regular;
    items.push({ name: String(name), next, remaining: total - regular * paid });
  }
  return {
    nextMonth,
    next: items.reduce((sum, item) => sum + item.next, 0),
    remaining: items.reduce((sum, item) => sum + item.remaining, 0),
    items
  };
}

// 同一JPEGの再送は既存ファイルを使い、登録済みならAPIを呼ばない。
function registerMobileReceipt(jpeg) {
  requireMobileOwner_();
  const maxBytes = 4 * 1024 * 1024;
  if (typeof jpeg !== "string" || jpeg.length > Math.ceil(maxBytes / 3) * 4 ||
      !/^[A-Za-z0-9+/]+={0,2}$/.test(jpeg)) {
    throw new Error("4MB以下のレシート画像を選んでください。");
  }
  const bytes = Utilities.base64Decode(jpeg);
  if (bytes.length < 4 || bytes.length > maxBytes ||
      (bytes[0] & 255) !== 255 || (bytes[1] & 255) !== 216 || (bytes[2] & 255) !== 255) {
    throw new Error("レシート画像をJPEGに変換できませんでした。");
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error("別のレシートを処理中です。少し待ってから再送してください。");
  try {
    const context = getReceiptImportContext_();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
    const hash = digest.map(byte => (byte & 255).toString(16).padStart(2, "0")).join("");
    const name = MOBILE_RECEIPT_PREFIX + hash + ".jpg";
    const done = context.done.getFilesByName(name);
    if (done.hasNext()) {
      const saved = receiptRowsById_(context.receipts, done.next().getId(), RECEIPT_HEADERS.length);
      if (saved.length !== 1 || saved[0][8] !== "読取済み") {
        throw new Error("読取完了フォルダの画像と登録済みレシートを確認してください。");
      }
      return { store: saved[0][2], total: saved[0][5], alreadyRegistered: true };
    }
    const pending = context.inbox.getFilesByName(name);
    const file = pending.hasNext() ? pending.next() :
      context.inbox.createFile(Utilities.newBlob(bytes, "image/jpeg", name));
    return importReceiptFile_(context, file);
  } finally {
    lock.releaseLock();
  }
}
