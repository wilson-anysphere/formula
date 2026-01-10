import { getCellRaw, normalizeCell } from "./normalizeCell.js";

/**
 * @param {any} sheet
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @param {{ maxRows?: number, maxCols?: number }} [opts]
 */
export function extractCells(sheet, rect, opts) {
  const maxRows = opts?.maxRows ?? Number.POSITIVE_INFINITY;
  const maxCols = opts?.maxCols ?? Number.POSITIVE_INFINITY;
  const rMax = Math.min(rect.r1, rect.r0 + maxRows - 1);
  const cMax = Math.min(rect.c1, rect.c0 + maxCols - 1);

  const out = [];
  for (let r = rect.r0; r <= rMax; r += 1) {
    const row = [];
    for (let c = rect.c0; c <= cMax; c += 1) {
      row.push(normalizeCell(getCellRaw(sheet, r, c)));
    }
    out.push(row);
  }
  return out;
}
