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
 *   maxScanRows?: number,
 *   maxScanCols?: number
 * }} params
 * @returns {PatternSuggestion[]}
 */
export function suggestPatternValues(params) {
  const cellRef = normalizeCellRef(params.cellRef);
  const { surroundingCells } = params;
  const maxScanRows = params.maxScanRows ?? 50;
  const maxScanCols = params.maxScanCols ?? 50;

  if (!surroundingCells || typeof surroundingCells.getCellValue !== "function") {
    return [];
  }

  const input = params.currentInput ?? "";
  const cursor = clampCursor(input, params.cursorPosition);
  const prefix = input.slice(0, cursor);
  if (prefix.startsWith("=")) return [];

  const normalizedPrefix = prefix.toLowerCase();
  const allowEmptyNumericPrefix = input.length === 0 && normalizedPrefix.length === 0;

  /** @type {Map<string, number>} */
  const scores = new Map();

  if (normalizedPrefix.length > 0) {
    // Scan up/down the current column for matching text values.
    for (let offset = 1; offset <= maxScanRows; offset++) {
      const weight = distanceWeight(offset);
      const up = cellRef.row - offset;
      const down = cellRef.row + offset;
      if (up >= 0) {
        const v = surroundingCells.getCellValue(up, cellRef.col);
        maybeScoreTextMatch(v, normalizedPrefix, scores, weight);
      }
      const v = surroundingCells.getCellValue(down, cellRef.col);
      maybeScoreTextMatch(v, normalizedPrefix, scores, weight);
    }

    // Scan left/right in the current row as well. Nearby patterns in the row are
    // often as useful as the current column (e.g. categorical labels).
    for (let offset = 1; offset <= maxScanCols; offset++) {
      const weight = distanceWeight(offset);
      const left = cellRef.col - offset;
      const right = cellRef.col + offset;
      if (left >= 0) {
        const v = surroundingCells.getCellValue(cellRef.row, left);
        maybeScoreTextMatch(v, normalizedPrefix, scores, weight);
      }
      const v = surroundingCells.getCellValue(cellRef.row, right);
      maybeScoreTextMatch(v, normalizedPrefix, scores, weight);
    }
  }

  // Basic numeric sequence completion: if the user is typing a number-like
  // prefix and the previous 2-3 numeric cells in the column form a stable step,
  // suggest the next value.
  const numericCandidate = suggestNextNumberInColumn({
    typedPrefix: prefix,
    cellRef,
    surroundingCells,
    maxScanRows,
    allowEmptyPrefix: allowEmptyNumericPrefix,
  });
  if (numericCandidate !== null) {
    addScore(numericCandidate, scores, 2);
  }

  const candidates = [...scores.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].length - b[0].length || a[0].localeCompare(b[0]))
    .slice(0, 5)
    .map(([text, score]) => ({
      text,
      confidence: scoreToConfidence(score),
    }));

  return candidates;
}

function maybeScoreTextMatch(value, normalizedPrefix, scores, weight) {
  if (isEmptyCell(value)) return;
  if (typeof value !== "string") return;
  const text = value;
  if (!text.toLowerCase().startsWith(normalizedPrefix)) return;
  addScore(text, scores, weight);
}

function addScore(text, scores, weight) {
  if (!text) return;
  const w = Number.isFinite(weight) ? weight : 0;
  if (w <= 0) return;
  scores.set(text, (scores.get(text) ?? 0) + w);
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

function distanceWeight(distance) {
  if (!Number.isFinite(distance) || distance <= 0) return 0;
  // Favor nearby matches strongly while still letting repeated patterns win.
  return 1 / distance;
}

function scoreToConfidence(score) {
  if (!Number.isFinite(score) || score <= 0) return 0.5;
  // Heuristic: single adjacent match => ~0.62, stronger evidence => up to ~0.9.
  return clamp01(0.5 + Math.min(0.4, score / 8));
}

function isNumericPrefix(text) {
  if (typeof text !== "string") return false;
  const trimmed = text.trim();
  if (!trimmed) return false;
  // Keep conservative: basic integers/decimals only (no scientific notation).
  return /^-?\d+(\.\d*)?$/.test(trimmed);
}

function tryParseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (typeof value !== "string") return null;
  const trimmed = value.trim();
  if (!trimmed) return null;
  if (!/^[-+]?\d+(\.\d+)?$/.test(trimmed)) return null;
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) ? parsed : null;
}

function nearlyEqual(a, b) {
  if (a === b) return true;
  const diff = Math.abs(a - b);
  const scale = Math.max(1, Math.abs(a), Math.abs(b));
  return diff <= 1e-9 * scale;
}

function formatNumberForCompletion(n) {
  if (!Number.isFinite(n)) return null;
  if (Number.isInteger(n)) return String(n);
  // Avoid ugly floating point representations for simple decimals.
  const rounded = Math.round(n * 1e12) / 1e12;
  return String(rounded);
}

function suggestNextNumberInColumn({ typedPrefix, cellRef, surroundingCells, maxScanRows, allowEmptyPrefix = false }) {
  const trimmedPrefix = typeof typedPrefix === "string" ? typedPrefix.trim() : "";
  const hasPrefix = trimmedPrefix.length > 0;
  if (hasPrefix) {
    if (!isNumericPrefix(trimmedPrefix)) return null;
  } else if (!allowEmptyPrefix) {
    return null;
  }

  /** @type {number[]} Nearest-first values above the current cell. */
  const previous = [];

  for (let offset = 1; offset <= maxScanRows; offset++) {
    const row = cellRef.row - offset;
    if (row < 0) break;
    const raw = surroundingCells.getCellValue(row, cellRef.col);
    if (isEmptyCell(raw)) {
      // Keep conservative: require the sequence to be contiguous.
      break;
    }
    const n = tryParseNumber(raw);
    if (n === null) {
      // Stop when we hit non-numeric data (headers, text).
      break;
    }
    previous.push(n);
    if (previous.length >= 3) break;
  }

  if (previous.length < 2) return null;

  const last = previous[0];
  const prev = previous[1];
  const step = last - prev;
  if (!Number.isFinite(step) || step === 0) return null;

  if (previous.length >= 3) {
    const prev2 = previous[2];
    const step2 = prev - prev2;
    if (!nearlyEqual(step, step2)) return null;
  }

  const next = last + step;
  const formatted = formatNumberForCompletion(next);
  if (formatted === null) return null;

  // Only suggest if it matches what the user is already typing.
  if (trimmedPrefix && !formatted.startsWith(trimmedPrefix)) return null;

  return formatted;
}
