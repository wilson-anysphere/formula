/**
 * @typedef {{ sheetId: string, row: number, col: number }} CellAddress
 */

/**
 * @param {CellAddress} cell
 * @returns {string}
 */
export function makeCellKey(cell) {
  return `${cell.sheetId}:${cell.row}:${cell.col}`;
}

/**
 * Parse a spreadsheet cell key. Supports:
 * - `${sheetId}:${row}:${col}` (canonical)
 * - `${sheetId}:${row},${col}` (legacy internal encoding)
 * - `r{row}c{col}` (unit-test convenience, resolved against `defaultSheetId`)
 *
 * @param {string} key
 * @param {{ defaultSheetId?: string }} [options]
 * @returns {CellAddress | null}
 */
export function parseCellKey(key, options = {}) {
  const defaultSheetId = options.defaultSheetId ?? "Sheet1";
  if (typeof key !== "string" || key.length === 0) return null;

  const parts = key.split(":");
  if (parts.length === 3) {
    const sheetId = parts[0] || defaultSheetId;
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) return null;
    return { sheetId, row, col };
  }

  if (parts.length === 2) {
    const sheetId = parts[0] || defaultSheetId;
    const m = parts[1].match(/^(\d+),(\d+)$/);
    if (m) {
      return { sheetId, row: Number(m[1]), col: Number(m[2]) };
    }
  }

  const m = key.match(/^r(\d+)c(\d+)$/);
  if (m) {
    return { sheetId: defaultSheetId, row: Number(m[1]), col: Number(m[2]) };
  }

  return null;
}

/**
 * Normalize a (potentially legacy) cell key into the canonical `${sheetId}:${row}:${col}` form.
 *
 * @param {string} key
 * @param {{ defaultSheetId?: string }} [options]
 * @returns {string | null}
 */
export function normalizeCellKey(key, options = {}) {
  const parsed = parseCellKey(key, options);
  return parsed ? makeCellKey(parsed) : null;
}

