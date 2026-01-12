function isAsciiLetter(ch: string): boolean {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isAsciiDigit(ch: string): boolean {
  return ch >= "0" && ch <= "9";
}

function isAsciiAlphaNum(ch: string): boolean {
  return isAsciiLetter(ch) || isAsciiDigit(ch);
}

function isReservedUnquotedSheetName(name: string): boolean {
  // Excel boolean literals are tokenized as keywords; quoting avoids ambiguity in formulas.
  return name.toLowerCase() === "true" || name.toLowerCase() === "false";
}

function looksLikeA1CellReference(name: string): boolean {
  // If an unquoted sheet name looks like a cell reference (e.g. "A1" or "XFD1048576"),
  // Excel requires quoting to disambiguate.
  //
  // Match the Rust backend's `looks_like_a1_cell_reference`.
  let i = 0;
  let letters = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiLetter(ch)) break;
    if (letters.length >= 3) return false;
    letters += ch;
    i += 1;
  }

  if (letters.length === 0) return false;

  let digits = "";
  while (i < name.length) {
    const ch = name[i];
    if (!ch || !isAsciiDigit(ch)) break;
    digits += ch;
    i += 1;
  }

  if (digits.length === 0) return false;
  if (i !== name.length) return false;

  const col = letters
    .split("")
    .reduce((acc, c) => acc * 26 + (c.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1), 0);
  return col <= 16_384;
}

function looksLikeR1C1CellReference(name: string): boolean {
  // In R1C1 notation, `R`/`C` are valid relative references. Excel may also treat
  // `R123C456` as a cell reference even when the workbook is in A1 mode.
  //
  // Match the Rust backend's `looks_like_r1c1_cell_reference`.
  const upper = name.toUpperCase();
  if (upper === "R" || upper === "C") return true;
  if (!upper.startsWith("R")) return false;

  let i = 1;
  while (i < upper.length && isAsciiDigit(upper[i] ?? "")) i += 1;
  if (i >= upper.length) return false;
  if (upper[i] !== "C") return false;

  i += 1;
  while (i < upper.length && isAsciiDigit(upper[i] ?? "")) i += 1;
  return i === upper.length;
}

export function escapeSheetNameForA1(sheetName: string): string {
  return String(sheetName ?? "").replace(/'/g, "''");
}

export function isValidUnquotedSheetNameForA1(name: string): boolean {
  if (!name) return false;
  const first = name[0];
  if (!first || isAsciiDigit(first)) return false;
  if (!(first === "_" || isAsciiLetter(first))) return false;

  for (let i = 1; i < name.length; i += 1) {
    const ch = name[i];
    if (!(isAsciiAlphaNum(ch) || ch === "_" || ch === ".")) return false;
  }

  if (isReservedUnquotedSheetName(name)) return false;
  if (looksLikeA1CellReference(name) || looksLikeR1C1CellReference(name)) return false;

  return true;
}

/**
 * Format a sheet name token for A1 references using Excel quoting conventions.
 *
 * Examples:
 * - `Sheet1` -> `Sheet1`
 * - `My Sheet` -> `'My Sheet'`
 * - `TRUE` -> `'TRUE'` (avoid boolean literal ambiguity)
 * - `A1` -> `'A1'` (avoid A1 cell reference ambiguity)
 */
export function formatSheetNameForA1(sheetName: string): string {
  const name = String(sheetName ?? "").trim();
  if (!name) return "";
  if (isValidUnquotedSheetNameForA1(name)) return name;
  return `'${escapeSheetNameForA1(name)}'`;
}

