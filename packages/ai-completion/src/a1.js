const COLUMN_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

/**
 * @param {number} index 0-based column index.
 * @returns {string}
 */
export function columnIndexToLetter(index) {
  if (!Number.isInteger(index) || index < 0) {
    throw new Error(`columnIndexToLetter: invalid index ${index}`);
  }

  let n = index;
  let letters = "";
  while (n >= 0) {
    letters = COLUMN_LETTERS[n % 26] + letters;
    n = Math.floor(n / 26) - 1;
  }
  return letters;
}

/**
 * @param {string} letters Column letters like A, Z, AA.
 * @returns {number} 0-based column index.
 */
export function columnLetterToIndex(letters) {
  if (typeof letters !== "string" || letters.length === 0) {
    throw new Error(`columnLetterToIndex: invalid letters ${letters}`);
  }

  const normalized = letters.toUpperCase();
  let result = 0;
  for (let i = 0; i < normalized.length; i++) {
    const code = normalized.charCodeAt(i);
    if (code < 65 || code > 90) {
      throw new Error(`columnLetterToIndex: invalid letters ${letters}`);
    }
    result = result * 26 + (code - 64);
  }
  return result - 1;
}

/**
 * @typedef {{row:number,col:number}} CellRefObject
 */

/**
 * @param {CellRefObject | string} cellRef
 * @returns {CellRefObject}
 */
export function normalizeCellRef(cellRef) {
  if (typeof cellRef === "string") {
    const parsed = parseA1(cellRef);
    if (!parsed) {
      throw new Error(`Invalid A1 cell reference: ${cellRef}`);
    }
    return parsed;
  }
  if (
    cellRef &&
    typeof cellRef === "object" &&
    Number.isInteger(cellRef.row) &&
    Number.isInteger(cellRef.col)
  ) {
    return { row: cellRef.row, col: cellRef.col };
  }
  throw new Error(`Invalid cellRef: ${String(cellRef)}`);
}

/**
 * @param {CellRefObject} ref 0-based row/col
 * @returns {string}
 */
export function toA1(ref) {
  if (!ref || !Number.isInteger(ref.row) || !Number.isInteger(ref.col)) {
    throw new Error(`toA1: invalid ref ${String(ref)}`);
  }
  return `${columnIndexToLetter(ref.col)}${ref.row + 1}`;
}

/**
 * @param {string} a1
 * @returns {CellRefObject | null}
 */
export function parseA1(a1) {
  if (typeof a1 !== "string") return null;
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(a1.trim());
  if (!match) return null;
  const colLetters = match[2];
  const rowStr = match[4];
  const col = columnLetterToIndex(colLetters);
  const row = Number(rowStr) - 1;
  if (!Number.isInteger(row) || row < 0) return null;
  return { row, col };
}

/**
 * @param {any} value
 * @returns {boolean}
 */
export function isEmptyCell(value) {
  return value === null || value === undefined || value === "";
}
