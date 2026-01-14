/**
 * Parse common spreadsheet-formatted numbers.
 *
 * This intentionally supports a conservative subset of formats commonly seen in
 * CSV / spreadsheet exports while avoiding accidental parsing of non-numeric
 * strings (e.g. dates like "2024-01-01").
 *
 * Supported:
 * - numbers (finite)
 * - numeric strings with:
 *   - commas as thousands separators ("1,200")
 *   - optional leading currency symbol ($, €, £) after an optional sign ("-$5", "$-5")
 *   - optional trailing percent ("10%" -> 0.1)
 *   - optional parentheses for negatives ("(1,200)" -> -1200)
 *   - leading-decimal floats (".5")
 */
export function parseSpreadsheetNumber(value: unknown): number | null {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (typeof value !== "string") return null;

  let text = value.trim();
  if (text === "") return null;

  // Preserve existing behavior for "real numbers" (including exponents, hex, etc.)
  // without trying to replicate JS's full numeric grammar.
  const direct = Number(text);
  if (Number.isFinite(direct)) return direct;

  let sign = 1;
  const hasParens = text.startsWith("(") && text.endsWith(")");

  // Accounting-style negatives: "(123)".
  if (hasParens) {
    text = text.slice(1, -1).trim();
    if (text === "") return null;
  }

  let isPercent = false;
  if (text.endsWith("%")) {
    isPercent = true;
    text = text.slice(0, -1).trim();
    if (text === "") return null;
  }

  let consumedLeadingSign = false;
  if (text.startsWith("+") || text.startsWith("-")) {
    consumedLeadingSign = true;
    if (text[0] === "-") sign *= -1;
    text = text.slice(1).trimStart();
    if (text === "") return null;
  }

  const currency = text[0];
  const consumedCurrency = currency === "$" || currency === "€" || currency === "£";
  if (consumedCurrency) {
    text = text.slice(1).trimStart();
    if (text === "") return null;
  }

  // Some exports render negative currency as "$-5" (sign after the symbol).
  if (consumedCurrency && !consumedLeadingSign && (text.startsWith("+") || text.startsWith("-"))) {
    if (text[0] === "-") sign *= -1;
    text = text.slice(1).trimStart();
    if (text === "") return null;
  }

  if (!isValidNumberWithThousandsSeparators(text)) return null;

  const normalized = text.replaceAll(",", "");
  const numeric = Number(normalized);
  if (!Number.isFinite(numeric)) return null;

  let result = numeric * sign;
  if (isPercent) result /= 100;
  // Parentheses typically indicate a negative number in spreadsheets.
  // Only apply the negation when the parsed value is positive to avoid
  // double-negating tokens like "(-5)" or "($-5)".
  if (hasParens && result > 0) result = -result;

  return Number.isFinite(result) ? result : null;
}

function isValidNumberWithThousandsSeparators(text: string): boolean {
  // Leading-decimal floats (".5"). We intentionally do not support commas here.
  if (/^\.\d+$/.test(text)) return true;

  // Integer part:
  // - either plain digits ("1200")
  // - or grouped thousands ("1,200", "12,345,678")
  // Optional fractional part (".5", ".").
  return /^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d*)?$/.test(text);
}
