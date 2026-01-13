export const EXCEL_MAX_ROWS = 1_048_576;
export const EXCEL_MAX_COLS = 16_384;
const EXCEL_MAX_ROW_INDEX = EXCEL_MAX_ROWS - 1;
const EXCEL_MAX_COL_INDEX = EXCEL_MAX_COLS - 1;

// Excel-style A1 cell refs (case-insensitive) with optional `$` absolute markers.
// Examples: A1, $A$1, $A1, A$1.
const A1_CELL_RE = /^\$?([A-Za-z]+)\$?(\d+)$/;
const A1_CELL_RANGE_RE = /^(?<start>\$?[A-Za-z]+\$?\d+)(?:\s*:\s*(?<end>\$?[A-Za-z]+\$?\d+))?$/;
const A1_COL_RANGE_RE = /^(?<start>\$?[A-Za-z]+)\s*:\s*(?<end>\$?[A-Za-z]+)$/;
const A1_ROW_RANGE_RE = /^(?<start>\$?\d+)\s*:\s*(?<end>\$?\d+)$/;

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
  if (!letters || !/^[A-Za-z]+$/.test(letters)) {
    throw new Error(`Invalid column letters: ${letters}`);
  }

  letters = letters.toUpperCase();
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
  const match = A1_CELL_RE.exec(String(a1Cell).trim());
  if (!match) throw new Error(`Invalid A1 cell reference: ${a1Cell}`);

  const [, colLetters, rowDigits] = match;
  const col = a1ToColumnIndex(colLetters);
  const rowNumber = Number(rowDigits);
  const row = rowNumber - 1;

  if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > EXCEL_MAX_ROWS) {
    throw new Error(`Invalid row in A1 ref: ${a1Cell}`);
  }
  if (!Number.isInteger(col) || col < 0 || col >= EXCEL_MAX_COLS) {
    throw new Error(`Invalid column in A1 ref: ${a1Cell}`);
  }
  return { row, col };
}

/**
 * @param {string} colRef
 */
function a1ToColumnRangeIndex(colRef) {
  const letters = String(colRef).trim().replace(/^\$/, "");
  const col = a1ToColumnIndex(letters);
  if (!Number.isInteger(col) || col < 0 || col >= EXCEL_MAX_COLS) {
    throw new Error(`Invalid column in A1 ref: ${colRef}`);
  }
  return col;
}

/**
 * @param {string} rowRef
 */
function a1ToRowRangeIndex(rowRef) {
  const digits = String(rowRef).trim().replace(/^\$/, "");
  const rowNumber = Number(digits);
  if (!Number.isInteger(rowNumber) || rowNumber < 1 || rowNumber > EXCEL_MAX_ROWS) {
    throw new Error(`Invalid row in A1 ref: ${rowRef}`);
  }
  return rowNumber - 1;
}

/**
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number, sheetName?: string }} range
 */
export function rangeToA1(range) {
  const normalized = normalizeRange(range);
  const prefix = normalized.sheetName ? `${formatSheetName(normalized.sheetName)}!` : "";

  // Whole-column ranges (e.g. A:C) cover all rows but only a subset of columns.
  if (normalized.startRow === 0 && normalized.endRow === EXCEL_MAX_ROW_INDEX) {
    const startCol = columnIndexToA1(normalized.startCol);
    const endCol = columnIndexToA1(normalized.endCol);
    return `${prefix}${startCol}:${endCol}`;
  }

  // Whole-row ranges (e.g. 1:10) cover all columns but only a subset of rows.
  if (normalized.startCol === 0 && normalized.endCol === EXCEL_MAX_COL_INDEX) {
    if (!Number.isInteger(normalized.startRow) || normalized.startRow < 0) {
      throw new Error(`Invalid row in range: ${JSON.stringify(range)}`);
    }
    if (!Number.isInteger(normalized.endRow) || normalized.endRow < 0) {
      throw new Error(`Invalid row in range: ${JSON.stringify(range)}`);
    }
    return `${prefix}${normalized.startRow + 1}:${normalized.endRow + 1}`;
  }

  const start = cellRefToA1({ row: normalized.startRow, col: normalized.startCol });
  const end = cellRefToA1({ row: normalized.endRow, col: normalized.endCol });
  const suffix = start === end ? start : `${start}:${end}`;
  return `${prefix}${suffix}`;
}

/**
 * Parse "Sheet1!A1:B2" or "A1" into a 0-indexed range.
 * @param {string} a1Range
 */
export function parseA1Range(a1Range) {
  const input = String(a1Range).trim();
  if (!input) throw new Error(`Invalid A1 range: ${a1Range}`);
  const bangIndex = input.lastIndexOf("!");
  const rawSheet = bangIndex === -1 ? "" : input.slice(0, bangIndex);
  const sheetName = bangIndex === -1 ? undefined : unescapeSheetName(rawSheet);
  if (bangIndex !== -1 && !sheetName) {
    throw new Error(`Invalid A1 range: missing sheet name in "${a1Range}"`);
  }
  const rest = bangIndex === -1 ? input : input.slice(bangIndex + 1).trim();
  if (!rest) throw new Error(`Invalid A1 range: missing range in "${a1Range}"`);

  const cellMatch = A1_CELL_RANGE_RE.exec(rest);
  if (cellMatch?.groups) {
    const start = a1ToCellRef(cellMatch.groups.start);
    const end = cellMatch.groups.end ? a1ToCellRef(cellMatch.groups.end) : start;
    return normalizeRange({
      sheetName,
      startRow: start.row,
      startCol: start.col,
      endRow: end.row,
      endCol: end.col,
    });
  }

  const colMatch = A1_COL_RANGE_RE.exec(rest);
  if (colMatch?.groups) {
    const startCol = a1ToColumnRangeIndex(colMatch.groups.start);
    const endCol = a1ToColumnRangeIndex(colMatch.groups.end);
    return normalizeRange({
      sheetName,
      startRow: 0,
      startCol,
      endRow: EXCEL_MAX_ROW_INDEX,
      endCol,
    });
  }

  const rowMatch = A1_ROW_RANGE_RE.exec(rest);
  if (rowMatch?.groups) {
    const startRow = a1ToRowRangeIndex(rowMatch.groups.start);
    const endRow = a1ToRowRangeIndex(rowMatch.groups.end);
    return normalizeRange({
      sheetName,
      startRow,
      startCol: 0,
      endRow,
      endCol: EXCEL_MAX_COL_INDEX,
    });
  }

  throw new Error(`Invalid A1 range: ${a1Range}`);
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
