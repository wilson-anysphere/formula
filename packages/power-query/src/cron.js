/**
 * Lightweight cron parser + next-run calculator.
 *
 * Supported format: 5 fields (minute hour day-of-month month day-of-week)
 *
 * Supported syntax (per field):
 *  - Wildcard: "*"
 *  - Lists: "1,2,3"
 *  - Ranges: "1-5"
 *  - Steps: "*\/5" or "1-10/2"
 *
 * Day-of-week uses 0-6 where 0 = Sunday. 7 is also accepted as Sunday.
 *
 * Day matching follows common cron semantics:
 *  - If day-of-month is "*" and day-of-week is "*" => match any day.
 *  - If either field is "*" => the other field must match.
 *  - If both fields are restricted => match when *either* field matches.
 */

/**
 * @typedef {"local" | "utc"} CronTimezone
 */

/**
 * @typedef {{
 *   source: string;
 *   minutes: number[];
 *   minutesSet: boolean[];
 *   hours: number[];
 *   hoursSet: boolean[];
 *   daysOfMonth: number[];
 *   daysOfMonthSet: boolean[];
 *   daysOfMonthAny: boolean;
 *   months: number[];
 *   monthsSet: boolean[];
 *   daysOfWeek: number[];
 *   daysOfWeekSet: boolean[];
 *   daysOfWeekAny: boolean;
 * }} CronSchedule
 */

/**
 * @param {string} expression
 * @returns {CronSchedule}
 */
export function parseCronExpression(expression) {
  const parts = expression.trim().split(/\s+/);
  if (parts.length !== 5) {
    throw new Error(`Cron expression must have 5 fields (got ${parts.length}): '${expression}'`);
  }

  const [minuteExpr, hourExpr, domExpr, monthExpr, dowExpr] = parts;

  const minutes = parseCronField(minuteExpr, 0, 59);
  const hours = parseCronField(hourExpr, 0, 23);
  const months = parseCronField(monthExpr, 1, 12);
  const daysOfMonth = parseCronField(domExpr, 1, 31);
  const daysOfWeek = parseCronField(dowExpr, 0, 7, { map: (v) => (v === 7 ? 0 : v) });

  // After mapping 7->0, normalize ranges to 0-6 for membership checks.
  const daysOfWeekSet = new Array(7).fill(false);
  /** @type {number[]} */
  const daysOfWeekValues = [];
  for (const v of daysOfWeek.values) {
    if (v < 0 || v > 6) throw new Error(`Invalid day-of-week '${v}' in '${expression}'`);
    if (!daysOfWeekSet[v]) {
      daysOfWeekSet[v] = true;
      daysOfWeekValues.push(v);
    }
  }
  daysOfWeekValues.sort((a, b) => a - b);

  return {
    source: expression,
    minutes: minutes.values,
    minutesSet: minutes.set,
    hours: hours.values,
    hoursSet: hours.set,
    daysOfMonth: daysOfMonth.values,
    daysOfMonthSet: daysOfMonth.set,
    daysOfMonthAny: daysOfMonth.any,
    months: months.values,
    monthsSet: months.set,
    daysOfWeek: daysOfWeekValues,
    daysOfWeekSet,
    daysOfWeekAny: daysOfWeek.any,
  };
}

/**
 * @typedef {{ values: number[], set: boolean[], any: boolean }} ParsedCronField
 */

/**
 * @param {string} field
 * @param {number} min
 * @param {number} max
 * @param {{ map?: (value: number) => number }} [options]
 * @returns {ParsedCronField}
 */
function parseCronField(field, min, max, options) {
  const trimmed = field.trim();
  const set = new Array(max + 1).fill(false);
  /** @type {number[]} */
  const values = [];

  /**
   * @param {number} raw
   */
  function addValue(raw) {
    const mapped = options?.map ? options.map(raw) : raw;
    if (!Number.isFinite(mapped) || mapped < min || mapped > max) {
      throw new Error(`Invalid cron value '${raw}' (expected ${min}-${max}) in '${field}'`);
    }
    if (!set[mapped]) {
      set[mapped] = true;
      values.push(mapped);
    }
  }

  if (trimmed === "*") {
    for (let v = min; v <= max; v++) addValue(v);
    values.sort((a, b) => a - b);
    return { values, set, any: true };
  }

  for (const part of trimmed.split(",")) {
    parseCronPart(part.trim(), { min, max, addValue });
  }

  if (values.length === 0) {
    throw new Error(`Cron field '${field}' does not select any values`);
  }

  values.sort((a, b) => a - b);
  return { values, set, any: false };
}

/**
 * @param {string} part
 * @param {{ min: number, max: number, addValue: (value: number) => void }} ctx
 */
function parseCronPart(part, ctx) {
  if (!part) return;

  const [rangePart, stepPart] = part.split("/");
  const step = stepPart === undefined ? null : parsePositiveInt(stepPart);

  /** @type {number} */
  let start;
  /** @type {number} */
  let end;

  if (rangePart === "*") {
    start = ctx.min;
    end = ctx.max;
  } else if (rangePart.includes("-")) {
    const [startStr, endStr] = rangePart.split("-");
    start = parseInt(startStr, 10);
    end = parseInt(endStr, 10);
  } else {
    start = parseInt(rangePart, 10);
    end = start;
  }

  if (!Number.isFinite(start) || !Number.isFinite(end)) {
    throw new Error(`Invalid cron field part '${part}'`);
  }

  if (step !== null) {
    if (step <= 0) throw new Error(`Cron step must be > 0 in '${part}'`);
    // When applying steps to a single value, treat it as "start-max/step" (common behavior).
    if (start === end) end = ctx.max;
    for (let v = start; v <= end; v += step) ctx.addValue(v);
  } else {
    if (start > end) throw new Error(`Cron range start must be <= end in '${part}'`);
    for (let v = start; v <= end; v++) ctx.addValue(v);
  }
}

/**
 * @param {string} value
 * @returns {number}
 */
function parsePositiveInt(value) {
  if (!/^\d+$/.test(value)) throw new Error(`Invalid integer '${value}'`);
  return parseInt(value, 10);
}

/**
 * Calculate the next scheduled run time in milliseconds.
 *
 * The returned time is always strictly after `afterMs` (cron scheduling is minute-granular).
 *
 * @param {CronSchedule} schedule
 * @param {number} afterMs
 * @param {CronTimezone} [timezone]
 * @returns {number}
 */
export function nextCronRun(schedule, afterMs, timezone = "local") {
  const date = new Date(afterMs);

  // Move to the start of the *next* minute.
  if (timezone === "utc") {
    date.setUTCSeconds(0, 0);
    date.setUTCMinutes(date.getUTCMinutes() + 1);
  } else {
    date.setSeconds(0, 0);
    date.setMinutes(date.getMinutes() + 1);
  }

  const minMinute = schedule.minutes[0];
  const minHour = schedule.hours[0];
  const minMonth = schedule.months[0];

  const startYear = timezone === "utc" ? date.getUTCFullYear() : date.getFullYear();
  const maxYear = startYear + 10;
  let iterations = 0;

  while ((timezone === "utc" ? date.getUTCFullYear() : date.getFullYear()) <= maxYear) {
    if (++iterations > 250_000) {
      throw new Error(`Cron next-run exceeded iteration limit for '${schedule.source}'`);
    }

    const year = timezone === "utc" ? date.getUTCFullYear() : date.getFullYear();
    const month = timezone === "utc" ? date.getUTCMonth() + 1 : date.getMonth() + 1;

    if (!schedule.monthsSet[month]) {
      const nextMonth = firstGreater(schedule.months, month);
      if (nextMonth !== null) {
        // Set day first to avoid month rollover (e.g. Jan 31 -> Feb becomes Mar 3).
        setDayOfMonth(date, timezone, 1);
        setMonth(date, timezone, nextMonth);
      } else {
        setDayOfMonth(date, timezone, 1);
        setYear(date, timezone, year + 1);
        setMonth(date, timezone, minMonth);
      }
      // Reset to the earliest time within the new month.
      setHour(date, timezone, minHour);
      setMinute(date, timezone, minMinute);
      continue;
    }

    const dom = timezone === "utc" ? date.getUTCDate() : date.getDate();
    const dow = timezone === "utc" ? date.getUTCDay() : date.getDay();
    const matchesDom = schedule.daysOfMonthSet[dom];
    const matchesDow = schedule.daysOfWeekSet[dow];
    const matchesDay = matchesCronDay(schedule, matchesDom, matchesDow);

    if (!matchesDay) {
      addDays(date, timezone, 1);
      setHour(date, timezone, minHour);
      setMinute(date, timezone, minMinute);
      continue;
    }

    const hour = timezone === "utc" ? date.getUTCHours() : date.getHours();
    if (!schedule.hoursSet[hour]) {
      const nextHour = firstGreater(schedule.hours, hour);
      if (nextHour !== null) {
        setHour(date, timezone, nextHour);
        setMinute(date, timezone, minMinute);
      } else {
        addDays(date, timezone, 1);
        setHour(date, timezone, minHour);
        setMinute(date, timezone, minMinute);
      }
      continue;
    }

    const minute = timezone === "utc" ? date.getUTCMinutes() : date.getMinutes();
    if (!schedule.minutesSet[minute]) {
      const nextMinute = firstGreater(schedule.minutes, minute);
      if (nextMinute !== null) {
        setMinute(date, timezone, nextMinute);
      } else {
        addHours(date, timezone, 1);
        setMinute(date, timezone, minMinute);
      }
      continue;
    }

    // Month/day/hour/minute all match.
    return date.getTime();
  }

  throw new Error(`Cron expression '${schedule.source}' did not match any time within 10 years`);
}

/**
 * @param {CronSchedule} schedule
 * @param {boolean} matchesDom
 * @param {boolean} matchesDow
 * @returns {boolean}
 */
function matchesCronDay(schedule, matchesDom, matchesDow) {
  if (schedule.daysOfMonthAny && schedule.daysOfWeekAny) return true;
  if (schedule.daysOfMonthAny) return matchesDow;
  if (schedule.daysOfWeekAny) return matchesDom;
  return matchesDom || matchesDow;
}

/**
 * @param {number[]} values
 * @param {number} current
 * @returns {number | null}
 */
function firstGreater(values, current) {
  for (const v of values) {
    if (v > current) return v;
  }
  return null;
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} year
 */
function setYear(date, timezone, year) {
  if (timezone === "utc") date.setUTCFullYear(year);
  else date.setFullYear(year);
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} month1Based
 */
function setMonth(date, timezone, month1Based) {
  if (timezone === "utc") date.setUTCMonth(month1Based - 1);
  else date.setMonth(month1Based - 1);
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} day
 */
function setDayOfMonth(date, timezone, day) {
  if (timezone === "utc") date.setUTCDate(day);
  else date.setDate(day);
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} hour
 */
function setHour(date, timezone, hour) {
  if (timezone === "utc") date.setUTCHours(hour, 0, 0, 0);
  else date.setHours(hour, 0, 0, 0);
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} minute
 */
function setMinute(date, timezone, minute) {
  if (timezone === "utc") date.setUTCMinutes(minute, 0, 0);
  else date.setMinutes(minute, 0, 0);
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} days
 */
function addDays(date, timezone, days) {
  if (timezone === "utc") date.setUTCDate(date.getUTCDate() + days);
  else date.setDate(date.getDate() + days);
}

/**
 * @param {Date} date
 * @param {CronTimezone} timezone
 * @param {number} hours
 */
function addHours(date, timezone, hours) {
  if (timezone === "utc") date.setUTCHours(date.getUTCHours() + hours);
  else date.setHours(date.getHours() + hours);
}
