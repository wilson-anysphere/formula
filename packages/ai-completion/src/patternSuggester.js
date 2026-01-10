import { normalizeCellRef, isEmptyCell } from "./a1.js";

/**
 * @typedef {{
 *   text: string,
 *   confidence: number
 * }} PatternSuggestion
 */

/**
 * @typedef {{
 *   getCellValue: (row: number, col: number) => any
 * }} CellContext
 */

/**
 * Suggest value completions by looking for repeated strings in nearby cells.
 *
 * This is intentionally conservative: it's meant to feel instant and helpful,
 * not clever. If we can't find strong evidence, we return nothing.
 *
 * @param {{
 *   currentInput: string,
 *   cursorPosition: number,
 *   cellRef: {row:number,col:number} | string,
 *   surroundingCells: CellContext,
 *   maxScanRows?: number
 * }} params
 * @returns {PatternSuggestion[]}
 */
export function suggestPatternValues(params) {
  const cellRef = normalizeCellRef(params.cellRef);
  const { surroundingCells } = params;
  const maxScanRows = params.maxScanRows ?? 50;

  if (!surroundingCells || typeof surroundingCells.getCellValue !== "function") {
    return [];
  }

  const cursor = clampCursor(params.currentInput ?? "", params.cursorPosition);
  const prefix = (params.currentInput ?? "").slice(0, cursor);
  if (prefix.length === 0) return [];
  if (prefix.startsWith("=")) return [];

  const normalizedPrefix = prefix.toLowerCase();

  /** @type {Map<string, number>} */
  const counts = new Map();

  // Scan up/down the current column for matching text values.
  for (let offset = 1; offset <= maxScanRows; offset++) {
    const up = cellRef.row - offset;
    const down = cellRef.row + offset;
    if (up >= 0) {
      const v = surroundingCells.getCellValue(up, cellRef.col);
      maybeCount(v, normalizedPrefix, counts);
    }
    const v = surroundingCells.getCellValue(down, cellRef.col);
    maybeCount(v, normalizedPrefix, counts);
  }

  const candidates = [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].length - b[0].length || a[0].localeCompare(b[0]))
    .slice(0, 5)
    .map(([text, count]) => ({
      text,
      confidence: clamp01(0.5 + Math.min(0.4, count / 10)),
    }));

  return candidates;
}

function maybeCount(value, normalizedPrefix, counts) {
  if (isEmptyCell(value)) return;
  if (typeof value !== "string") return;
  const text = value;
  if (!text.toLowerCase().startsWith(normalizedPrefix)) return;
  counts.set(text, (counts.get(text) ?? 0) + 1);
}

function clampCursor(input, cursorPosition) {
  if (!Number.isInteger(cursorPosition)) return input.length;
  if (cursorPosition < 0) return 0;
  if (cursorPosition > input.length) return input.length;
  return cursorPosition;
}

function clamp01(v) {
  return Math.max(0, Math.min(1, v));
}
