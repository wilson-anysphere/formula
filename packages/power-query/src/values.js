/**
 * Scalar value helpers for Power Query "core" types that cannot be represented
 * as plain JS primitives without losing information (e.g. decimal precision,
 * datetime offsets, durations).
 *
 * This file is intentionally JS + JSDoc so it can run in Node without a TS build.
 */

export const MS_PER_DAY = 24 * 60 * 60 * 1000;

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
  const trimmed = input.trim();
  if (trimmed === "") return null;

  // Date-only: 2024-01-31
  const dateOnly = /^(\d{4})-(\d{2})-(\d{2})$/.exec(trimmed);
  if (dateOnly) {
    const [, y, m, d] = dateOnly;
    return new Date(Date.UTC(Number(y), Number(m) - 1, Number(d)));
  }

  // Datetime with optional seconds/fraction and optional timezone.
  const dt =
    /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2})(\.\d{1,9})?)?([zZ]|([+-])(\d{2}):(\d{2}))?$/.exec(
      trimmed,
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

  const out = new Date(utcMs);
  return Number.isNaN(out.getTime()) ? null : out;
}

/**
 * @param {Date} value
 * @returns {boolean}
 */
export function hasUtcTimeComponent(value) {
  return (
    value.getUTCHours() !== 0 ||
    value.getUTCMinutes() !== 0 ||
    value.getUTCSeconds() !== 0 ||
    value.getUTCMilliseconds() !== 0
  );
}

/**
 * @param {number} n
 * @param {number} width
 * @returns {string}
 */
function padNumber(n, width) {
  const s = String(Math.trunc(Math.abs(n)));
  return s.length >= width ? s : `${"0".repeat(width - s.length)}${s}`;
}

/**
 * @param {number} ms
 * @returns {string}
 */
function formatTimeFromMs(ms) {
  const normalized = ((ms % MS_PER_DAY) + MS_PER_DAY) % MS_PER_DAY;
  const hours = Math.floor(normalized / 3_600_000);
  const minutes = Math.floor((normalized % 3_600_000) / 60_000);
  const seconds = Math.floor((normalized % 60_000) / 1000);
  const millis = Math.floor(normalized % 1000);

  const base = `${padNumber(hours, 2)}:${padNumber(minutes, 2)}:${padNumber(seconds, 2)}`;
  return millis === 0 ? base : `${base}.${padNumber(millis, 3)}`;
}

/**
 * @param {string} input
 * @returns {number | null}
 */
function parseTimeToMs(input) {
  const trimmed = input.trim();
  if (trimmed === "") return null;
  const match = /^(\d{2}):(\d{2})(?::(\d{2})(\.\d{1,9})?)?$/.exec(trimmed);
  if (!match) return null;
  const [, hhRaw, mmRaw, ssRaw, fracRaw] = match;
  const hh = Number(hhRaw);
  const mm = Number(mmRaw);
  const ss = ssRaw ? Number(ssRaw) : 0;
  const millis = fracRaw ? Math.round(Number(`0${fracRaw}`) * 1000) : 0;

  if (!Number.isFinite(hh) || hh < 0 || hh > 23) return null;
  if (!Number.isFinite(mm) || mm < 0 || mm > 59) return null;
  if (!Number.isFinite(ss) || ss < 0 || ss > 59) return null;

  return hh * 3_600_000 + mm * 60_000 + ss * 1000 + millis;
}

export class PqDecimal {
  /**
   * @param {string} value
   */
  constructor(value) {
    this.value = String(value);
  }

  toString() {
    return this.value;
  }

  valueOf() {
    const parsed = Number(this.value);
    return Number.isFinite(parsed) ? parsed : Number.NaN;
  }
}

/**
 * @param {unknown} value
 * @returns {value is PqDecimal}
 */
export function isPqDecimal(value) {
  return value instanceof PqDecimal;
}

export class PqTime {
  /**
   * @param {number} milliseconds
   */
  constructor(milliseconds) {
    // Normalize to [0, MS_PER_DAY) to match the semantic "time of day" range.
    this.milliseconds = ((milliseconds % MS_PER_DAY) + MS_PER_DAY) % MS_PER_DAY;
  }

  /**
   * Parse an ISO-like time value (HH:mm[:ss[.sss]]).
   * @param {string} input
   * @returns {PqTime | null}
   */
  static from(input) {
    const ms = parseTimeToMs(input);
    return ms == null ? null : new PqTime(ms);
  }

  toString() {
    return formatTimeFromMs(this.milliseconds);
  }

  valueOf() {
    return this.milliseconds;
  }
}

/**
 * @param {unknown} value
 * @returns {value is PqTime}
 */
export function isPqTime(value) {
  return value instanceof PqTime;
}

/**
 * @param {number} ms
 * @returns {string}
 */
function formatDurationFromMs(ms) {
  const negative = ms < 0;
  let remaining = Math.abs(ms);

  const days = Math.floor(remaining / MS_PER_DAY);
  remaining -= days * MS_PER_DAY;
  const hours = Math.floor(remaining / 3_600_000);
  remaining -= hours * 3_600_000;
  const minutes = Math.floor(remaining / 60_000);
  remaining -= minutes * 60_000;
  const seconds = Math.floor(remaining / 1000);
  remaining -= seconds * 1000;
  const millis = remaining;

  const parts = [];
  if (days) parts.push(`${days}D`);

  const timeParts = [];
  if (hours) timeParts.push(`${hours}H`);
  if (minutes) timeParts.push(`${minutes}M`);
  if (seconds || millis || (days === 0 && hours === 0 && minutes === 0)) {
    const sec = millis ? `${seconds}.${padNumber(millis, 3)}` : String(seconds);
    timeParts.push(`${sec}S`);
  }

  return `${negative ? "-" : ""}P${parts.join("")}${timeParts.length ? `T${timeParts.join("")}` : ""}`;
}

/**
 * Parse a conservative subset of ISO 8601 durations.
 *
 * Supports:
 * - PnDTnHnMnS
 * - PTnS / PTn.nS
 *
 * (Years/months are intentionally unsupported because they are not fixed-length.)
 *
 * @param {string} input
 * @returns {number | null}
 */
function parseDurationToMs(input) {
  const trimmed = input.trim();
  if (trimmed === "") return null;
  const match = /^(-)?P(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?)?$/.exec(trimmed);
  if (!match) return null;

  const [, negRaw, daysRaw, hoursRaw, minutesRaw, secondsRaw] = match;
  const days = daysRaw ? Number(daysRaw) : 0;
  const hours = hoursRaw ? Number(hoursRaw) : 0;
  const minutes = minutesRaw ? Number(minutesRaw) : 0;
  const seconds = secondsRaw ? Number(secondsRaw) : 0;

  if (![days, hours, minutes, seconds].every((n) => Number.isFinite(n))) return null;

  let ms = 0;
  ms += days * MS_PER_DAY;
  ms += hours * 3_600_000;
  ms += minutes * 60_000;
  ms += seconds * 1000;

  if (negRaw) ms = -ms;
  return ms;
}

export class PqDuration {
  /**
   * @param {number} milliseconds
   */
  constructor(milliseconds) {
    this.milliseconds = milliseconds;
  }

  /**
   * @param {string} input
   * @returns {PqDuration | null}
   */
  static from(input) {
    const ms = parseDurationToMs(input);
    return ms == null ? null : new PqDuration(ms);
  }

  toString() {
    return formatDurationFromMs(this.milliseconds);
  }

  valueOf() {
    return this.milliseconds;
  }
}

/**
 * @param {unknown} value
 * @returns {value is PqDuration}
 */
export function isPqDuration(value) {
  return value instanceof PqDuration;
}

/**
 * @typedef {{ date: Date, offsetMinutes: number }} ParsedDateTimeZone
 */

/**
 * @param {string} input
 * @returns {ParsedDateTimeZone | null}
 */
function parseDateTimeZone(input) {
  const trimmed = input.trim();
  if (trimmed === "") return null;

  // Accept:
  // - YYYY-MM-DD
  // - YYYY-MM-DDTHH:mm[:ss[.sss]](Z|±HH:MM|±HHMM)?
  const dt =
    /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2})(\.\d{1,9})?)?)?(?:([zZ])|([+-])(\d{2}):?(\d{2}))?$/.exec(
      trimmed,
    );
  if (!dt) return null;

  const [, y, m, d, hhRaw, mmRaw, ssRaw, fracRaw, zRaw, signRaw, offHRaw, offMRaw] = dt;
  const hh = hhRaw ? Number(hhRaw) : 0;
  const mm = mmRaw ? Number(mmRaw) : 0;
  const ss = ssRaw ? Number(ssRaw) : 0;
  const millis = fracRaw ? Math.round(Number(`0${fracRaw}`) * 1000) : 0;

  if (![hh, mm, ss, millis].every((n) => Number.isFinite(n))) return null;

  let offsetMinutes = 0;
  if (zRaw && zRaw.toLowerCase() === "z") {
    offsetMinutes = 0;
  } else if (signRaw) {
    const sign = signRaw === "-" ? -1 : 1;
    const offH = Number(offHRaw);
    const offM = Number(offMRaw);
    if (!Number.isFinite(offH) || !Number.isFinite(offM)) return null;
    offsetMinutes = sign * (offH * 60 + offM);
  }

  const utcMs =
    Date.UTC(Number(y), Number(m) - 1, Number(d), hh, mm, ss, millis) - offsetMinutes * 60 * 1000;
  const date = new Date(utcMs);
  if (Number.isNaN(date.getTime())) return null;

  return { date, offsetMinutes };
}

/**
 * @param {Date} dateUtc
 * @param {number} offsetMinutes
 * @returns {string}
 */
function formatDateTimeZone(dateUtc, offsetMinutes) {
  const local = new Date(dateUtc.getTime() + offsetMinutes * 60 * 1000);
  const y = local.getUTCFullYear();
  const mo = local.getUTCMonth() + 1;
  const d = local.getUTCDate();
  const hh = local.getUTCHours();
  const mm = local.getUTCMinutes();
  const ss = local.getUTCSeconds();
  const ms = local.getUTCMilliseconds();

  const base = `${padNumber(y, 4)}-${padNumber(mo, 2)}-${padNumber(d, 2)}T${padNumber(hh, 2)}:${padNumber(mm, 2)}:${padNumber(ss, 2)}.${padNumber(ms, 3)}`;

  if (offsetMinutes === 0) return `${base}Z`;

  const sign = offsetMinutes < 0 ? "-" : "+";
  const abs = Math.abs(offsetMinutes);
  const offH = Math.floor(abs / 60);
  const offM = abs % 60;
  return `${base}${sign}${padNumber(offH, 2)}:${padNumber(offM, 2)}`;
}

export class PqDateTimeZone {
  /**
   * @param {Date} dateUtc
   * @param {number} offsetMinutes
   */
  constructor(dateUtc, offsetMinutes) {
    this.date = dateUtc;
    this.offsetMinutes = offsetMinutes;
  }

  /**
   * Parse an ISO-like datetimezone string.
   * @param {string} input
   * @returns {PqDateTimeZone | null}
   */
  static from(input) {
    const parsed = parseDateTimeZone(input);
    return parsed ? new PqDateTimeZone(parsed.date, parsed.offsetMinutes) : null;
  }

  /**
   * @returns {Date}
   */
  toDate() {
    return this.date;
  }

  toString() {
    return formatDateTimeZone(this.date, this.offsetMinutes);
  }

  valueOf() {
    return this.date.getTime();
  }
}

/**
 * @param {unknown} value
 * @returns {value is PqDateTimeZone}
 */
export function isPqDateTimeZone(value) {
  return value instanceof PqDateTimeZone;
}
