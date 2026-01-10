/**
 * @param {import('./workbookTypes').Sheet} sheet
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 */
import { getSheetMatrix, normalizeCell } from "./normalizeCell.js";

/**
 * @param {any} sheet
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 */
export function extractCells(sheet, rect) {
  const matrix = getSheetMatrix(sheet);
  const out = [];
  for (let r = rect.r0; r <= rect.r1; r += 1) {
    const row = [];
    const srcRow = matrix[r] || [];
    for (let c = rect.c0; c <= rect.c1; c += 1) {
      row.push(normalizeCell(srcRow[c]));
    }
    out.push(row);
  }
  return out;
}
