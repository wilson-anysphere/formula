/**
 * @typedef {{ row: number, col: number }} CellCoord
 */

/**
 * @param {number} colIndex Zero-based
 * @returns {string}
 */
export function columnIndexToName(colIndex) {
  if (!Number.isInteger(colIndex) || colIndex < 0) {
    throw new Error(`Invalid column index: ${colIndex}`);
  }

  let n = colIndex + 1; // 1-based
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

/**
 * @param {string} name
 * @returns {number} Zero-based column index
 */
export function columnNameToIndex(name) {
  if (typeof name !== "string" || name.length === 0) {
    throw new Error(`Invalid column name: ${name}`);
  }

  const upper = name.toUpperCase();
  let n = 0;
  for (const ch of upper) {
    if (ch < "A" || ch > "Z") throw new Error(`Invalid column name: ${name}`);
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n - 1;
}

/**
 * @param {string} a1
 * @returns {CellCoord}
 */
export function parseA1(a1) {
  if (typeof a1 !== "string") throw new Error(`Invalid A1 ref: ${a1}`);
  const match = /^([A-Za-z]+)(\d+)$/.exec(a1.trim());
  if (!match) throw new Error(`Invalid A1 ref: ${a1}`);
  const [, colName, rowStr] = match;
  const row = Number.parseInt(rowStr, 10) - 1;
  const col = columnNameToIndex(colName);
  if (!Number.isInteger(row) || row < 0) throw new Error(`Invalid A1 ref: ${a1}`);
  return { row, col };
}

/**
 * @param {CellCoord} coord
 * @returns {string}
 */
export function formatA1(coord) {
  if (!coord || !Number.isInteger(coord.row) || !Number.isInteger(coord.col)) {
    throw new Error(`Invalid cell coord: ${coord}`);
  }
  return `${columnIndexToName(coord.col)}${coord.row + 1}`;
}

/**
 * @typedef {{ start: CellCoord, end: CellCoord }} CellRange
 */

/**
 * @param {CellRange} range
 * @returns {CellRange}
 */
export function normalizeRange(range) {
  const startRow = Math.min(range.start.row, range.end.row);
  const endRow = Math.max(range.start.row, range.end.row);
  const startCol = Math.min(range.start.col, range.end.col);
  const endCol = Math.max(range.start.col, range.end.col);
  return { start: { row: startRow, col: startCol }, end: { row: endRow, col: endCol } };
}

/**
 * @param {string} a1Range
 * @returns {CellRange}
 */
export function parseRangeA1(a1Range) {
  if (typeof a1Range !== "string") throw new Error(`Invalid A1 range: ${a1Range}`);
  const trimmed = a1Range.trim();
  const parts = trimmed.split(":");
  if (parts.length === 1) {
    const c = parseA1(parts[0]);
    return { start: c, end: c };
  }
  if (parts.length !== 2) throw new Error(`Invalid A1 range: ${a1Range}`);
  return normalizeRange({ start: parseA1(parts[0]), end: parseA1(parts[1]) });
}

