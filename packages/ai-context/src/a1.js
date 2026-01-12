const A1_CELL_RE = /^([A-Z]+)(\d+)$/;
const A1_RANGE_RE = /^(?<start>[A-Z]+\d+)(?::(?<end>[A-Z]+\d+))?$/;

function isAsciiLetter(ch) {
  return ch >= "A" && ch <= "Z" ? true : ch >= "a" && ch <= "z";
}

function isAsciiDigit(ch) {
  return ch >= "0" && ch <= "9";
}

function isAsciiAlphaNum(ch) {
  return isAsciiLetter(ch) || isAsciiDigit(ch);
}

function isReservedUnquotedSheetName(name) {
  // Excel boolean literals are tokenized as keywords; quoting avoids ambiguity in formulas.
  const lower = String(name ?? "").toLowerCase();
  return lower === "true" || lower === "false";
}

function looksLikeA1CellReference(name) {
  // If an unquoted sheet name looks like a cell reference (e.g. "A1" or "XFD1048576"),
  // Excel requires quoting to disambiguate.
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

function looksLikeR1C1CellReference(name) {
  // In R1C1 notation, `R`/`C` are valid relative references. Excel may also treat
  // `R123C456` as a cell reference even when the workbook is in A1 mode.
  const upper = String(name ?? "").toUpperCase();
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

function isValidUnquotedSheetNameForA1(name) {
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
 * Excel-style sheet names:
 *  - Sheet1
 *  - 'My Sheet'
 *  - 'Bob''s Sheet' (escaped single quotes)
 *
 * @param {string} rawSheet
 */
function unescapeSheetName(rawSheet) {
  const sheet = rawSheet.trim();
  if (sheet.startsWith("'") && sheet.endsWith("'")) {
    return sheet.slice(1, -1).replace(/''/g, "'");
  }
  return sheet;
}

/**
 * Quote sheet names when needed for Excel-compatible A1 references.
 *
 * @param {string} sheetName
 */
function formatSheetName(sheetName) {
  // Accept either a raw sheet name ("My Sheet") or an already-quoted Excel sheet
  // name ("'My Sheet'"). Canonicalize to an unquoted name before formatting so we
  // don't double-quote.
  sheetName = unescapeSheetName(sheetName);
  // Identifier-like sheet names can be used without quoting.
  //
  // Note: avoid emitting ambiguous identifiers like `TRUE!A1`, `A1!A1`, or `R1C1!A1`.
  if (isValidUnquotedSheetNameForA1(sheetName)) return sheetName;
  // Excel style: wrap in single quotes and escape embedded quotes via doubling.
  return `'${sheetName.replace(/'/g, "''")}'`;
}

/**
 * @param {unknown} value
 * @returns {boolean}
 */
export function isCellEmpty(value) {
  return value === null || value === undefined || value === "";
}

/**
 * 0 -> A, 25 -> Z, 26 -> AA
 * @param {number} columnIndex
 */
export function columnIndexToA1(columnIndex) {
  if (!Number.isInteger(columnIndex) || columnIndex < 0) {
    throw new Error(`columnIndex must be a non-negative integer, got: ${columnIndex}`);
  }

  let n = columnIndex + 1;
  let letters = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

/**
 * @param {string} letters
 */
export function a1ToColumnIndex(letters) {
  if (!letters || !/^[A-Z]+$/.test(letters)) {
    throw new Error(`Invalid column letters: ${letters}`);
  }

  let value = 0;
  for (const char of letters) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value - 1;
}

/**
 * @param {{ row: number, col: number }} cell
 */
export function cellRefToA1(cell) {
  if (!cell || !Number.isInteger(cell.row) || !Number.isInteger(cell.col)) {
    throw new Error(`Invalid cell ref: ${JSON.stringify(cell)}`);
  }
  return `${columnIndexToA1(cell.col)}${cell.row + 1}`;
}

/**
 * @param {string} a1Cell
 */
export function a1ToCellRef(a1Cell) {
  const match = A1_CELL_RE.exec(a1Cell);
  if (!match) throw new Error(`Invalid A1 cell reference: ${a1Cell}`);

  const [, colLetters, rowDigits] = match;
  const row = Number(rowDigits) - 1;
  const col = a1ToColumnIndex(colLetters);

  if (!Number.isInteger(row) || row < 0) throw new Error(`Invalid row in A1 ref: ${a1Cell}`);
  return { row, col };
}

/**
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number, sheetName?: string }} range
 */
export function rangeToA1(range) {
  const start = cellRefToA1({ row: range.startRow, col: range.startCol });
  const end = cellRefToA1({ row: range.endRow, col: range.endCol });
  const suffix = start === end ? start : `${start}:${end}`;
  return range.sheetName ? `${formatSheetName(range.sheetName)}!${suffix}` : suffix;
}

/**
 * Parse "Sheet1!A1:B2" or "A1" into a 0-indexed range.
 * @param {string} a1Range
 */
export function parseA1Range(a1Range) {
  const input = String(a1Range).trim();
  const bangIndex = input.lastIndexOf("!");
  const rawSheet = bangIndex === -1 ? "" : input.slice(0, bangIndex);
  const sheetName = bangIndex === -1 ? undefined : unescapeSheetName(rawSheet);
  if (bangIndex !== -1 && !sheetName) {
    throw new Error(`Invalid A1 range: missing sheet name in "${a1Range}"`);
  }
  const rest = bangIndex === -1 ? input : input.slice(bangIndex + 1).trim();

  const match = A1_RANGE_RE.exec(rest);
  if (!match || !match.groups) throw new Error(`Invalid A1 range: ${a1Range}`);

  const start = a1ToCellRef(match.groups.start);
  const end = match.groups.end ? a1ToCellRef(match.groups.end) : start;

  return normalizeRange({
    sheetName,
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  });
}

/**
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number, sheetName?: string }} range
 */
export function normalizeRange(range) {
  const startRow = Math.min(range.startRow, range.endRow);
  const endRow = Math.max(range.startRow, range.endRow);
  const startCol = Math.min(range.startCol, range.endCol);
  const endCol = Math.max(range.startCol, range.endCol);
  return {
    sheetName: range.sheetName,
    startRow,
    startCol,
    endRow,
    endCol,
  };
}
