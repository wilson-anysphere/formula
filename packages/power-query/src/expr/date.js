/**
 * @param {string} text
 * @returns {Date}
 */
export function parseDateLiteral(text) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(text);
  if (!match) {
    throw new Error(`Invalid date literal '${text}' (expected YYYY-MM-DD)`);
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  // `Date.UTC` normalizes out-of-range values (e.g. month 13), so we validate
  // by round-tripping the produced date components.
  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    !Number.isFinite(date.getTime()) ||
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    throw new Error(`Invalid date literal '${text}'`);
  }

  return date;
}

