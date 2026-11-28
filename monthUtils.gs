// 日付の繰り越しを避け、指定タイムゾーンの暦で前月を求める。
function getPreviousMonth_(timezone, date = new Date()) {
  const [year, month] = Utilities.formatDate(date, timezone, "yyyy-MM").split("-").map(Number);
  return {
    year: month === 1 ? year - 1 : year,
    month: month === 1 ? 12 : month - 1
  };
}

function normalizeMonthKey_(value, timezone, optional = false) {
  if (value instanceof Date) return Utilities.formatDate(value, timezone, "yyyy-MM");
  const match = String(value).trim().match(/^(\d{4})[/-](0?[1-9]|1[0-2])$/);
  if (!match && optional) return null;
  if (!match) throw new Error("対象月の形式を確認してください。");
  return match[1] + "-" + match[2].padStart(2, "0");
}

function findMonthIndex_(headers, target, timezone) {
  const key = normalizeMonthKey_(target, timezone, true);
  if (!key) return -1;
  return headers.findIndex(value => normalizeMonthKey_(value, timezone, true) === key);
}
