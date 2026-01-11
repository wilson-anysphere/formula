/**
 * @typedef {object} CellRef
 * @property {string} sheetId
 * @property {number} row 0-based
 * @property {number} col 0-based
 */

/**
 * @param {CellRef} ref
 */
export function cellKeyFromRef(ref) {
  return `${ref.sheetId}:${ref.row}:${ref.col}`;
}

/**
 * @param {string} key
 * @returns {CellRef}
 */
export function cellRefFromKey(key) {
  // Unit-test convenience encoding (sheet resolved implicitly to "Sheet1").
  // CollabSession's `parseCellKey` uses the same default.
  const rc = key.match(/^r(\d+)c(\d+)$/);
  if (rc) {
    return {
      sheetId: "Sheet1",
      row: Number.parseInt(rc[1], 10),
      col: Number.parseInt(rc[2], 10),
    };
  }

  const parts = key.split(":");
  if (parts.length === 3) {
    const [sheetId, rowStr, colStr] = parts;
    return {
      sheetId,
      row: Number.parseInt(rowStr, 10),
      col: Number.parseInt(colStr, 10)
    };
  }

  // Some internal modules use `${sheetId}:${row},${col}`.
  if (parts.length === 2) {
    const [sheetId, tail] = parts;
    const [rowStr, colStr] = tail.split(",");
    return {
      sheetId,
      row: Number.parseInt(rowStr, 10),
      col: Number.parseInt(colStr, 10)
    };
  }

  throw new Error(`Invalid cell key: ${key}`);
}

/**
 * @param {string} colLetters
 * @returns {number} 0-based
 */
export function colToNumber(colLetters) {
  const upper = colLetters.toUpperCase();
  let num = 0;
  for (let i = 0; i < upper.length; i += 1) {
    num = num * 26 + (upper.charCodeAt(i) - 64);
  }
  return num - 1;
}

/**
 * @param {number} col 0-based
 * @returns {string}
 */
export function numberToCol(col) {
  let n = col + 1;
  let out = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}
