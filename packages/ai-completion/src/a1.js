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
  const cellPart = stripSheetPrefix(a1);
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(cellPart.trim());
  if (!match) return null;
  const colLetters = match[2];
  const rowStr = match[4];
  const col = columnLetterToIndex(colLetters);
  const row = Number(rowStr) - 1;
  if (!Number.isInteger(row) || row < 0) return null;
  return { row, col };
}

/**
 * Strip an optional sheet prefix from a sheet-qualified A1 reference.
 *
 * Examples:
 * - Sheet1!A1 -> A1
 * - 'My Sheet'!$B$2 -> $B$2
 * - 'Bob''s Sheet'!C3 -> C3
 *
 * @param {string} text
 * @returns {string}
 */
function stripSheetPrefix(text) {
  const trimmed = text.trim();
  if (!trimmed.includes("!")) return trimmed;

  // Handle quoted sheet names: 'My Sheet'!A1
  if (trimmed.startsWith("'")) {
    let i = 1;
    while (i < trimmed.length) {
      const ch = trimmed[i];
      if (ch === "'") {
        // Escaped apostrophe inside sheet name: '' -> '
        if (trimmed[i + 1] === "'") {
          i += 2;
          continue;
        }
        // Closing quote must be followed by '!' to be a valid sheet prefix.
        if (trimmed[i + 1] === "!") {
          return trimmed.slice(i + 2);
        }
        // Malformed quoting; fall back to returning the original string so parsing fails.
        return trimmed;
      }
      i += 1;
    }
    // Unterminated quote; fall back to returning the original string so parsing fails.
    return trimmed;
  }

  // Unquoted sheet names: Sheet1!A1
  const bang = trimmed.indexOf("!");
  if (bang <= 0) return trimmed;
  return trimmed.slice(bang + 1);
}

/**
 * @param {any} value
 * @returns {boolean}
 */
export function isEmptyCell(value) {
  return value === null || value === undefined || value === "";
}
