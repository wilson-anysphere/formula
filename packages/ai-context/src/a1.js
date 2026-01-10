const A1_CELL_RE = /^([A-Z]+)(\d+)$/;
const A1_RANGE_RE = /^(?:(?<sheet>[^!]+)!)?(?<start>[A-Z]+\d+)(?::(?<end>[A-Z]+\d+))?$/;

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
  return range.sheetName ? `${range.sheetName}!${suffix}` : suffix;
}

/**
 * Parse "Sheet1!A1:B2" or "A1" into a 0-indexed range.
 * @param {string} a1Range
 */
export function parseA1Range(a1Range) {
  const match = A1_RANGE_RE.exec(a1Range);
  if (!match || !match.groups) throw new Error(`Invalid A1 range: ${a1Range}`);

  const sheetName = match.groups.sheet;
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
