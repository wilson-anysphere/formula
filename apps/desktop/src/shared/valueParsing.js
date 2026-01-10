/**
 * @typedef {"empty" | "number" | "boolean" | "datetime" | "string"} InferredValueType
 * @typedef {{ value: string | number | boolean | null, type: InferredValueType, numberFormat: string | undefined }} ParsedScalar
 */

/**
 * Infer a scalar's type using conservative, locale-agnostic heuristics.
 *
 * We intentionally avoid locale-dependent parsing (e.g. `1/2/2024`) because it
 * introduces ambiguity between MM/DD and DD/MM formats.
 *
 * @param {string} rawInput
 * @returns {InferredValueType}
 */
export function inferValueType(rawInput) {
  return parseScalar(rawInput).type;
}

/**
 * Parse a scalar value from clipboard/CSV text.
 *
 * - Booleans: TRUE/FALSE (case-insensitive)
 * - Numbers: accepts decimals + exponent, rejects leading-zero integers (e.g. 00123)
 * - Datetimes: ISO 8601 / RFC3339-ish values are converted to Excel serial numbers.
 *
 * @param {string} rawInput
 * @returns {ParsedScalar}
 */
export function parseScalar(rawInput) {
  const trimmed = rawInput.trim();
  if (trimmed === "") return { value: null, type: "empty" };

  if (isBooleanString(trimmed)) {
    return { value: trimmed.toLowerCase() === "true", type: "boolean" };
  }

  if (isNumberString(trimmed)) {
    const num = Number(trimmed);
    if (Number.isFinite(num)) return { value: num, type: "number" };
  }

  if (isIsoLikeDateString(trimmed)) {
    const parsed = parseIsoLikeToUtcDate(trimmed);
    if (parsed) {
      const serial = dateToExcelSerial(parsed);
      const dateOnly = /^\d{4}-\d{2}-\d{2}$/.test(trimmed);
      return {
        value: serial,
        type: "datetime",
        numberFormat: dateOnly ? "yyyy-mm-dd" : "yyyy-mm-dd hh:mm:ss",
      };
    }
  }

  return { value: rawInput, type: "string" };
}

/**
 * Convenience wrapper that returns only the parsed value.
 *
 * @param {string} rawInput
 * @returns {string | number | boolean | null}
 */
export function parseScalarValue(rawInput) {
  return parseScalar(rawInput).value;
}

/**
 * Guarded numeric detection intended for clipboard/CSV ingestion:
 * - Accepts optional sign, decimals, and exponent.
 * - Rejects integers with leading zeros (e.g. `00123`) to preserve common ID codes.
 */
export function isNumberString(rawInput) {
  const input = rawInput.trim();
  if (input === "") return false;

  // Preserve common "ID" patterns like 00123.
  if (/^[+-]?0\d+$/.test(input)) return false;

  return /^[+-]?(?:\d+\.?\d*|\d*\.?\d+)(?:[eE][+-]?\d+)?$/.test(input);
}

export function isBooleanString(rawInput) {
  const input = rawInput.trim().toLowerCase();
  return input === "true" || input === "false";
}

/**
 * Date inference is intentionally conservative to avoid locale ambiguity.
 * Accepts ISO 8601 / RFC3339-ish timestamps as well as `YYYY-MM-DD` dates.
 */
export function isIsoLikeDateString(rawInput) {
  const input = rawInput.trim();

  // Date only: 2024-01-31
  if (/^\d{4}-\d{2}-\d{2}$/.test(input)) return true;

  // Datetime: 2024-01-31T12:34:56Z / 2024-01-31 12:34:56
  if (/^\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(:\d{2}(\.\d{1,9})?)?([zZ]|[+-]\d{2}:\d{2})?$/.test(input)) {
    return true;
  }

  return false;
}

const MS_PER_DAY = 24 * 60 * 60 * 1000;
const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 31);

/**
 * Convert a UTC `Date` into an Excel 1900-date-system serial number.
 * Accounts for Excel's 1900 leap year bug by inserting the fictitious 1900-02-29.
 *
 * @param {Date} dateUtc
 * @returns {number}
 */
export function dateToExcelSerial(dateUtc) {
  const utcMs = dateUtc.getTime();
  let serial = (utcMs - EXCEL_EPOCH_UTC_MS) / MS_PER_DAY;

  // Excel includes 1900-02-29 as day 60; bump everything from 1900-03-01 onward.
  if (serial >= 60) serial += 1;

  return serial;
}

/**
 * Convert an Excel 1900-date-system serial number back to a UTC Date.
 *
 * @param {number} serial
 * @returns {Date}
 */
export function excelSerialToDate(serial) {
  let adjusted = serial;
  if (adjusted >= 61) adjusted -= 1;
  const utcMs = EXCEL_EPOCH_UTC_MS + adjusted * MS_PER_DAY;
  return new Date(utcMs);
}

/**
 * Best-effort parse of ISO-like values into a UTC Date.
 *
 * Supports:
 * - YYYY-MM-DD
 * - YYYY-MM-DDTHH:mm[:ss[.sss]](Z|±HH:MM)?
 * - YYYY-MM-DD HH:mm[:ss[.sss]](Z|±HH:MM)?
 *
 * If no timezone is present, interpret as UTC (not local) for determinism.
 *
 * @param {string} input
 * @returns {Date | null}
 */
export function parseIsoLikeToUtcDate(input) {
  // Date-only.
  const dateOnly = /^(\d{4})-(\d{2})-(\d{2})$/.exec(input);
  if (dateOnly) {
    const [, y, m, d] = dateOnly;
    return new Date(Date.UTC(Number(y), Number(m) - 1, Number(d)));
  }

  const dt =
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2})(\.\d{1,9})?)?([zZ]|([+-])(\d{2}):(\d{2}))?$/.exec(
      input
    );
  if (!dt) return null;

  const [, y, m, d, hh, mm, ssRaw, fracRaw, tzRaw, signRaw, offHRaw, offMRaw] = dt;
  const ss = ssRaw ? Number(ssRaw) : 0;
  const millis = fracRaw ? Math.round(Number(`0${fracRaw}`) * 1000) : 0;

  // Interpret missing timezone as UTC for deterministic parsing.
  let offsetMinutes = 0;
  if (tzRaw && tzRaw.toLowerCase() !== "z") {
    const sign = signRaw === "-" ? -1 : 1;
    const offH = Number(offHRaw);
    const offM = Number(offMRaw);
    offsetMinutes = sign * (offH * 60 + offM);
  }

  const utcMs =
    Date.UTC(Number(y), Number(m) - 1, Number(d), Number(hh), Number(mm), ss, millis) -
    offsetMinutes * 60 * 1000;

  return new Date(utcMs);
}
