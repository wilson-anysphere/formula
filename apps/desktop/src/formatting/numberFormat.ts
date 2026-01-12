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

function looksLikeDateFormat(format: string): boolean {
  // Keep v1 conservative: treat presence of Y/M/D tokens as a date/time format.
  return /[ymdhis]/i.test(format);
}

function formatExcelDate(serial: number, format: string): string {
  if (!Number.isFinite(serial)) return "";

  const date = excelSerialToDate(serial);

  const y = date.getUTCFullYear();
  const m = date.getUTCMonth() + 1;
  const d = date.getUTCDate();

  const lower = format.toLowerCase();

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
  const isCurrency = format.includes("$");
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

  const prefix = isCurrency ? "$" : "";
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

  if (looksLikeDateFormat(section)) {
    return formatExcelDate(value, section);
  }

  return formatNumeric(value, section);
}

