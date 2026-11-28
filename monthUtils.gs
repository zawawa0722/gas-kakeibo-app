// 日付の繰り越しを避け、指定タイムゾーンの暦で前月を求める。
function getPreviousMonth_(timezone, date = new Date()) {
  const [year, month] = Utilities.formatDate(date, timezone, "yyyy-MM").split("-").map(Number);
  return {
    year: month === 1 ? year - 1 : year,
    month: month === 1 ? 12 : month - 1
  };
}
