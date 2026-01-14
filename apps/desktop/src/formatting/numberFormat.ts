import { excelSerialToDate } from "../shared/valueParsing.js";

function formatNumberGeneral(value: number): string {
  if (!Number.isFinite(value)) return "0";
  // Mirror the heuristics used for chart tick labels (see `charts/scene/format.ts`).
  const rounded = Math.round(value * 1000) / 1000;
  if (Object.is(rounded, -0)) return "0";
  if (Number.isInteger(rounded)) return String(rounded);
  return rounded.toFixed(3).replace(/\.?0+$/, "");
}

function pad2(value: number): string {
  return String(value).padStart(2, "0");
}

function groupThousands(integer: string): string {
  // Assume `integer` contains only digits.
  if (integer.length <= 3) return integer;
  let out = "";
  let i = integer.length;
  while (i > 3) {
    const start = i - 3;
    out = `,${integer.slice(start, i)}${out}`;
    i = start;
  }
  return `${integer.slice(0, i)}${out}`;
}

function parseDecimalPlaces(format: string): number {
  const dot = format.indexOf(".");
  if (dot === -1) return 0;
  let count = 0;
  for (let i = dot + 1; i < format.length; i++) {
    const ch = format[i];
    if (ch === "0" || ch === "#") count += 1;
    else break;
  }
  return count;
}

function hasThousandsSeparators(format: string): boolean {
  const integerSection = format.split(".")[0] ?? format;
  return integerSection.includes(",");
}

function gcd(a: number, b: number): number {
  let x = Math.abs(Math.trunc(a));
  let y = Math.abs(Math.trunc(b));
  while (y !== 0) {
    const t = x % y;
    x = y;
    y = t;
  }
  return x === 0 ? 1 : x;
}

function isSupportedDateFormat(format: string): boolean {
  // v1 only supports a couple of known presets. Avoid treating arbitrary numeric
  // formats containing letters (e.g. `0.0,"M"`) as dates.
  const lower = format.toLowerCase();
  if (lower.includes("m/d/yyyy") || lower.includes("yyyy-mm-dd")) return true;

  // Time-only presets (Excel-style). These are commonly used for Ctrl+Shift+; insertion.
  // Keep the detection intentionally narrow to avoid mis-classifying arbitrary formats.
  const compact = lower.replace(/\s+/g, "");
  return /^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact);
}

function parseScientificFormat(format: string): { decimals: number; expDigits: number } | null {
  const upper = format.toUpperCase();
  const match = /E([+-])([0]+)/.exec(upper);
  if (!match) return null;
  const expDigits = match[2]?.length ?? 0;
  if (expDigits <= 0) return null;
  const base = upper.slice(0, match.index);
  const decimals = parseDecimalPlaces(base);
  return { decimals, expDigits };
}

function formatScientific(value: number, { decimals, expDigits }: { decimals: number; expDigits: number }): string {
  if (!Number.isFinite(value)) return "0";
  const sign = value < 0 && !Object.is(value, -0) ? "-" : "";
  const abs = Math.abs(value);
  if (abs === 0) {
    const mantissa = (0).toFixed(decimals);
    return `${sign}${mantissa}E+${"0".padStart(expDigits, "0")}`;
  }

  // Compute exponent in base 10.
  let exponent = Math.floor(Math.log10(abs));
  let mantissa = abs / Math.pow(10, exponent);

  // Round mantissa to the required decimals.
  const rounded = Number(mantissa.toFixed(decimals));
  mantissa = rounded;
  if (mantissa >= 10) {
    mantissa /= 10;
    exponent += 1;
  }

  const mantissaText = mantissa.toFixed(decimals);
  const expSign = exponent >= 0 ? "+" : "-";
  const expText = Math.abs(exponent).toString().padStart(expDigits, "0");
  return `${sign}${mantissaText}E${expSign}${expText}`;
}

function parseFractionFormat(format: string): { maxDenominator: number } | null {
  // Restrict to classic "?" placeholder-based fraction formats (e.g. "# ?/?", "# ??/??").
  if (!format.includes("?")) return null;
  const match = /\/\?+/.exec(format);
  if (!match) return null;
  const digits = match[0].length - 1; // exclude the "/"
  if (digits <= 0) return null;
  const maxDenominator = Math.pow(10, digits) - 1;
  return { maxDenominator };
}

function formatFraction(value: number, { maxDenominator }: { maxDenominator: number }): string {
  if (!Number.isFinite(value)) return "";

  const isNegative = value < 0 && !Object.is(value, -0);
  const abs = Math.abs(value);
  const whole = Math.floor(abs);
  const frac = abs - whole;

  if (frac < 1e-12) return `${isNegative ? "-" : ""}${whole}`;

  let bestNumerator = 0;
  let bestDenominator = 1;
  let bestError = Infinity;

  for (let d = 1; d <= maxDenominator; d += 1) {
    const n = Math.round(frac * d);
    const approx = n / d;
    const error = Math.abs(frac - approx);
    if (error < bestError) {
      bestError = error;
      bestNumerator = n;
      bestDenominator = d;
      if (error < 1e-12) break;
    }
  }

  if (bestNumerator === 0) return `${isNegative ? "-" : ""}${whole}`;

  if (bestNumerator >= bestDenominator) {
    // Rounding pushed us up to 1.
    const nextWhole = whole + 1;
    return `${isNegative ? "-" : ""}${nextWhole}`;
  }

  const divisor = gcd(bestNumerator, bestDenominator);
  const numerator = bestNumerator / divisor;
  const denominator = bestDenominator / divisor;

  const sign = isNegative ? "-" : "";
  if (whole === 0) return `${sign}${numerator}/${denominator}`;
  return `${sign}${whole} ${numerator}/${denominator}`;
}

function formatExcelDate(serial: number, format: string): string {
  if (!Number.isFinite(serial)) return "";

  const date = excelSerialToDate(serial);

  const y = date.getUTCFullYear();
  const m = date.getUTCMonth() + 1;
  const d = date.getUTCDate();

  const lower = format.toLowerCase();
  const compact = lower.replace(/\s+/g, "");

  // Time-only formats (no date component).
  const timeOnlyMatch = /^h{1,2}:m{1,2}(:s{1,2})?$/.exec(compact);
  if (timeOnlyMatch) {
    const hh = date.getUTCHours();
    const mm = date.getUTCMinutes();
    const ss = date.getUTCSeconds();
    const hasSeconds = compact.includes(":s");
    return hasSeconds ? `${pad2(hh)}:${pad2(mm)}:${pad2(ss)}` : `${pad2(hh)}:${pad2(mm)}`;
  }

  let dateText = "";
  if (lower.includes("m/d/yyyy")) {
    dateText = `${m}/${d}/${String(y).padStart(4, "0")}`;
  } else if (lower.includes("yyyy-mm-dd")) {
    dateText = `${String(y).padStart(4, "0")}-${pad2(m)}-${pad2(d)}`;
  } else {
    // Unknown date pattern: default to ISO date for determinism.
    dateText = `${String(y).padStart(4, "0")}-${pad2(m)}-${pad2(d)}`;
  }

  const includeTime = /[hs]/.test(lower);
  if (!includeTime) return dateText;

  const hh = date.getUTCHours();
  const mm = date.getUTCMinutes();
  const ss = date.getUTCSeconds();
  return `${dateText} ${pad2(hh)}:${pad2(mm)}:${pad2(ss)}`;
}

function formatNumeric(value: number, format: string): string {
  if (!Number.isFinite(value)) return "0";

  const isPercent = format.includes("%");
  const currencyMatch = /[$€£¥]/.exec(format);
  const isCurrency = Boolean(currencyMatch);
  const currencySymbol = currencyMatch?.[0] ?? "";
  const decimals = parseDecimalPlaces(format);
  const useThousands = hasThousandsSeparators(format);

  const scaled = isPercent ? value * 100 : value;

  // Use `toFixed` for deterministic decimal rounding/zero padding.
  const fixed = Math.abs(scaled).toFixed(decimals);
  const [integerPart, decimalPart] = fixed.split(".");

  const groupedInteger = useThousands ? groupThousands(integerPart) : integerPart;
  const numericText = decimalPart !== undefined ? `${groupedInteger}.${decimalPart}` : groupedInteger;

  // Avoid printing "-0.00" for tiny values.
  const isNegative = scaled < 0 && !Object.is(Number(fixed), 0);
  const sign = isNegative ? "-" : "";

  const prefix = isCurrency ? currencySymbol : "";
  const suffix = isPercent ? "%" : "";
  return `${sign}${prefix}${numericText}${suffix}`;
}

/**
 * Minimal number format support used by the shared-grid canvas renderer.
 *
 * This intentionally supports only a small subset of Excel-style codes (currency, percent,
 * and a couple of date presets). It is *not* intended to be a full formatter.
 */
export function formatValueWithNumberFormat(value: number, numberFormat: string): string {
  const raw = typeof numberFormat === "string" ? numberFormat : "";
  const section = (raw.split(";")[0] ?? "").trim();
  if (section === "" || section.toLowerCase() === "general") return formatNumberGeneral(value);

  if (isSupportedDateFormat(section)) {
    return formatExcelDate(value, section);
  }

  const scientific = parseScientificFormat(section);
  if (scientific) {
    return formatScientific(value, scientific);
  }

  const fraction = parseFractionFormat(section);
  if (fraction) {
    return formatFraction(value, fraction);
  }

  return formatNumeric(value, section);
}

/**
 * Best-effort validation for Excel-style custom number format codes.
 *
 * Notes:
 * - This is intentionally **not** a full Excel number format parser.
 * - The goal is to catch obvious syntax errors (unbalanced quotes/brackets, too many sections)
 *   so the custom format prompt can avoid storing clearly-invalid codes.
 * - Valid Excel codes may still be rendered approximately by `formatValueWithNumberFormat` since
 *   our canvas renderer supports only a subset of Excel formatting features.
 */
export function isValidExcelNumberFormatCode(format: string): boolean {
  if (typeof format !== "string") return false;

  let inQuotes = false;
  let bracketDepth = 0;
  let sections = 1;

  for (let i = 0; i < format.length; i += 1) {
    const ch = format[i];
    if (ch === "\\") {
      // Excel uses backslash to escape the next character (treat it as a literal).
      i += 1;
      if (i >= format.length) return false;
      continue;
    }

    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if (inQuotes) continue;

    if (ch === ";") {
      sections += 1;
      if (sections > 4) return false;
      continue;
    }

    // `*` repeats the next character; `_` reserves space for it. Missing the trailing
    // character is a syntax error.
    if (ch === "*" || ch === "_") {
      if (i + 1 >= format.length) return false;
      i += 1;
      continue;
    }

    if (ch === "[") {
      bracketDepth += 1;
      continue;
    }

    if (ch === "]") {
      bracketDepth -= 1;
      if (bracketDepth < 0) return false;
      continue;
    }
  }

  if (inQuotes) return false;
  if (bracketDepth !== 0) return false;
  return true;
}
