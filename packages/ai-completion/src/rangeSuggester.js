import { columnLetterToIndex, isEmptyCell, normalizeCellRef } from "./a1.js";

/**
 * @typedef {{range: string, confidence: number, reason: string}} RangeSuggestion
 */

/**
 * @typedef {{
 *   getCellValue: (row: number, col: number) => any
 * }} CellContext
 */

/**
 * Suggests likely ranges based on contiguous non-empty cells near the current cell.
 *
 * The algorithm is intentionally simple and fast:
 * - If the user typed a column letter (e.g. "A"), suggest the contiguous block of
 *   non-empty cells above the current row in that column (A1:A10).
 * - Also suggest the entire column (A:A) at lower confidence.
 *
 * @param {{
 *   currentArgText: string,
 *   cellRef: {row:number,col:number} | string,
 *   surroundingCells: CellContext,
 *   maxScanRows?: number
 * }} params
 * @returns {RangeSuggestion[]}
 */
export function suggestRanges(params) {
  const { currentArgText, surroundingCells } = params;
  const cellRef = normalizeCellRef(params.cellRef);
  const maxScanRows = params.maxScanRows ?? 500;

  if (!surroundingCells || typeof surroundingCells.getCellValue !== "function") {
    return [];
  }

  const arg = (currentArgText ?? "").trim();
  if (arg.length === 0) return [];

  // Only handle simple column/cell prefixes for now (A, A1, $A$1, etc).
  const match = /^\$?([A-Za-z]{1,3})(?:\$?(\d+))?$/.exec(arg);
  if (!match) return [];

  const colLetters = match[1].toUpperCase();
  const colIndex = safeColumnLetterToIndex(colLetters);
  if (colIndex === null) return [];

  const explicitRow = match[2] ? Number(match[2]) : null;
  if (explicitRow !== null && (!Number.isInteger(explicitRow) || explicitRow <= 0)) return [];

  /** @type {RangeSuggestion[]} */
  const suggestions = [];

  const contiguous = explicitRow
    ? findContiguousBlockDown(surroundingCells, colIndex, explicitRow - 1, maxScanRows)
    : findContiguousBlockAbove(surroundingCells, colIndex, cellRef.row - 1, maxScanRows);

  if (contiguous) {
    const { startRow, endRow, numericRatio } = contiguous;
    const startA1 = `${colLetters}${startRow + 1}`;
    const endA1 = `${colLetters}${endRow + 1}`;
    const range = `${startA1}:${endA1}`;

    // Confidence heuristic:
    // - Longer contiguous numeric blocks are more likely to be "the" range for SUM, etc.
    const length = endRow - startRow + 1;
    const base = 0.7;
    const lengthBoost = Math.min(0.2, length / 50);
    const numericBoost = 0.1 * numericRatio;
    suggestions.push({
      range,
      confidence: clamp01(base + lengthBoost + numericBoost),
      reason: explicitRow ? "contiguous_down_from_start" : "contiguous_above_current_cell",
    });
  }

  suggestions.push({
    range: `${colLetters}:${colLetters}`,
    confidence: 0.3,
    reason: "entire_column",
  });

  return suggestions;
}

function safeColumnLetterToIndex(letters) {
  try {
    return columnLetterToIndex(letters);
  } catch {
    return null;
  }
}

/**
 * @param {CellContext} ctx
 * @param {number} col
 * @param {number} fromRow start scanning from this row upwards (inclusive)
 * @param {number} maxScanRows
 */
function findContiguousBlockAbove(ctx, col, fromRow, maxScanRows) {
  if (fromRow < 0) return null;

  // Find the nearest non-empty cell above (skip blank separators).
  let endRow = fromRow;
  let scanned = 0;
  while (endRow >= 0 && scanned < maxScanRows && isEmptyCell(ctx.getCellValue(endRow, col))) {
    endRow--;
    scanned++;
  }
  if (endRow < 0) return null;

  let startRow = endRow;
  scanned = 0;
  while (startRow >= 0 && scanned < maxScanRows && !isEmptyCell(ctx.getCellValue(startRow, col))) {
    startRow--;
    scanned++;
  }
  startRow++;

  const metrics = computeNumericRatio(ctx, col, startRow, endRow);
  return { startRow, endRow, numericRatio: metrics.numericRatio };
}

/**
 * @param {CellContext} ctx
 * @param {number} col
 * @param {number} startRow
 * @param {number} maxScanRows
 */
function findContiguousBlockDown(ctx, col, startRow, maxScanRows) {
  if (startRow < 0) return null;

  // If the explicitly provided start cell is empty, we don't have a good signal.
  if (isEmptyCell(ctx.getCellValue(startRow, col))) return null;

  let endRow = startRow;
  let scanned = 0;
  while (scanned < maxScanRows && !isEmptyCell(ctx.getCellValue(endRow + 1, col))) {
    endRow++;
    scanned++;
  }

  const metrics = computeNumericRatio(ctx, col, startRow, endRow);
  return { startRow, endRow, numericRatio: metrics.numericRatio };
}

function computeNumericRatio(ctx, col, startRow, endRow) {
  let numeric = 0;
  let total = 0;
  for (let r = startRow; r <= endRow; r++) {
    const v = ctx.getCellValue(r, col);
    if (isEmptyCell(v)) continue;
    total++;
    if (typeof v === "number" && Number.isFinite(v)) numeric++;
    else if (typeof v === "string" && v.trim() !== "" && Number.isFinite(Number(v))) numeric++;
  }
  return { numericRatio: total === 0 ? 0 : numeric / total };
}

function clamp01(v) {
  return Math.max(0, Math.min(1, v));
}
