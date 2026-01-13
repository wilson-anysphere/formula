import { columnIndexToLetter, columnLetterToIndex, isEmptyCell, normalizeCellRef } from "./a1.js";

/**
 * @typedef {{range: string, confidence: number, reason: string}} RangeSuggestion
 */

/**
 * @typedef {{
 *   getCellValue: (row: number, col: number, sheetName?: string) => any
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
 *   sheetName?: string,
 *   maxScanRows?: number,
 *   maxScanCols?: number
 * }} params
 * @returns {RangeSuggestion[]}
 */
export function suggestRanges(params) {
  const { currentArgText, surroundingCells } = params;
  const cellRef = normalizeCellRef(params.cellRef);
  const sheetName = params.sheetName;
  const maxScanRows = params.maxScanRows ?? 500;
  const maxScanCols = params.maxScanCols ?? 50;

  if (!surroundingCells || typeof surroundingCells.getCellValue !== "function") {
    return [];
  }

  const arg = (currentArgText ?? "").trim();
  if (arg.length === 0) return [];

  // Only handle simple column/cell prefixes for now (A, A1, $A$1, etc).
  const match = /^(\$?)([A-Za-z]{1,3})(?:(\$?)(\d+))?$/.exec(arg);
  if (!match) return [];

  const colPrefix = match[1] === "$" ? "$" : "";
  const colToken = match[2];
  const colLetters = colToken.toUpperCase();
  const colIndex = safeColumnLetterToIndex(colLetters);
  if (colIndex === null) return [];

  const rowPrefix = match[3] === "$" ? "$" : "";
  const explicitRow = match[4] ? Number(match[4]) : null;
  if (explicitRow !== null && (!Number.isInteger(explicitRow) || explicitRow <= 0)) return [];

  /** @type {RangeSuggestion[]} */
  const suggestions = [];

  /** @type {{startRow:number,endRow:number,numericRatio:number} | null} */
  let contiguous = null;
  let contiguousReason = "";
  if (explicitRow !== null) {
    contiguous = findContiguousBlockDown(surroundingCells, colIndex, explicitRow - 1, maxScanRows, sheetName);
    contiguousReason = "contiguous_down_from_start";
  } else {
    contiguous = findContiguousBlockAbove(surroundingCells, colIndex, cellRef.row - 1, maxScanRows, sheetName);
    contiguousReason = "contiguous_above_current_cell";
    if (!contiguous) {
      // When the active cell isn't in the referenced column, it's safe (and often
      // desirable) to include data from the same row (e.g. formula in B2 summing A2:A11).
      const startRow = cellRef.row + (cellRef.col === colIndex ? 1 : 0);
      contiguous = findContiguousBlockBelow(surroundingCells, colIndex, startRow, maxScanRows, sheetName);
      contiguousReason = "contiguous_below_current_cell";
    }
  }

  /** @type {RangeSuggestion | null} */
  let tableSuggestion = null;

  if (contiguous) {
    const { startRow, endRow, numericRatio } = contiguous;
    const startA1 = `${colPrefix}${colToken}${rowPrefix}${startRow + 1}`;
    const endA1 = `${colPrefix}${colToken}${rowPrefix}${endRow + 1}`;
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
      reason: contiguousReason,
    });

    // Also suggest a 2D rectangular range when the column looks like part of a table.
    const tableStartRow = contiguous.rawStartRow ?? startRow;
    const tableEndRow = contiguous.rawEndRow ?? endRow;
    const table = findContiguousTableToRight(
      surroundingCells,
      colIndex,
      tableStartRow,
      tableEndRow,
      maxScanCols,
      sheetName
    );
    if (table) {
      const { endCol, confidence } = table;
      const endLetters = applyColumnCase(columnIndexToLetter(endCol), colToken);
      const startCell = `${colPrefix}${colToken}${rowPrefix}${tableStartRow + 1}`;
      const endCell = `${colPrefix}${endLetters}${rowPrefix}${tableEndRow + 1}`;
      tableSuggestion = {
        range: `${startCell}:${endCell}`,
        confidence,
        reason: explicitRow ? "contiguous_table_down_from_start" : "contiguous_table_above_current_cell",
      };
    }
  }

  suggestions.push({
    range: `${colPrefix}${colToken}:${colPrefix}${colToken}`,
    confidence: 0.3,
    reason: "entire_column",
  });

  // Keep "entire column" as the second suggestion for compatibility with existing
  // tests/callers, then append the optional 2D table range.
  if (tableSuggestion) suggestions.push(tableSuggestion);

  return suggestions;
}

const EXCEL_MAX_COL_INDEX = 16383; // XFD (1-based 16384)

function safeColumnLetterToIndex(letters) {
  try {
    const idx = columnLetterToIndex(letters);
    // Excel's last column is XFD (0-based 16383). Avoid suggesting out-of-bounds
    // ranges if the user types a 3-letter column beyond Excel's limit (e.g. ZZZ).
    if (idx > EXCEL_MAX_COL_INDEX) return null;
    return idx;
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
function findContiguousBlockAbove(ctx, col, fromRow, maxScanRows, sheetName) {
  if (fromRow < 0) return null;

  // Find the nearest non-empty cell above (skip blank separators).
  let endRow = fromRow;
  let scanned = 0;
  while (endRow >= 0 && scanned < maxScanRows && isEmptyCell(ctx.getCellValue(endRow, col, sheetName))) {
    endRow--;
    scanned++;
  }
  if (endRow < 0) return null;

  let startRow = endRow;
  scanned = 0;
  while (startRow >= 0 && scanned < maxScanRows && !isEmptyCell(ctx.getCellValue(startRow, col, sheetName))) {
    startRow--;
    scanned++;
  }
  startRow++;

  const rawStartRow = startRow;
  const rawEndRow = endRow;

  const trimmed = trimNonNumericEdgesIfMostlyNumeric(ctx, col, startRow, endRow, sheetName);
  return {
    startRow: trimmed.startRow,
    endRow: trimmed.endRow,
    numericRatio: trimmed.numericRatio,
    rawStartRow,
    rawEndRow,
  };
}

/**
 * @param {CellContext} ctx
 * @param {number} col
 * @param {number} startRow
 * @param {number} maxScanRows
 */
function findContiguousBlockDown(ctx, col, startRow, maxScanRows, sheetName) {
  if (startRow < 0) return null;

  // If the explicitly provided start cell is empty, we don't have a good signal.
  if (isEmptyCell(ctx.getCellValue(startRow, col, sheetName))) return null;

  let endRow = startRow;
  let scanned = 0;
  while (scanned < maxScanRows && !isEmptyCell(ctx.getCellValue(endRow + 1, col, sheetName))) {
    endRow++;
    scanned++;
  }

  const metrics = computeNumericStats(ctx, col, startRow, endRow, sheetName);
  return { startRow, endRow, numericRatio: metrics.numericRatio, rawStartRow: startRow, rawEndRow: endRow };
}

/**
 * @param {CellContext} ctx
 * @param {number} col
 * @param {number} fromRow start scanning from this row downwards (inclusive)
 * @param {number} maxScanRows
 */
function findContiguousBlockBelow(ctx, col, fromRow, maxScanRows, sheetName) {
  if (fromRow < 0) fromRow = 0;

  // Find the nearest non-empty cell below (skip blank separators).
  let startRow = fromRow;
  let scanned = 0;
  while (scanned < maxScanRows && isEmptyCell(ctx.getCellValue(startRow, col, sheetName))) {
    startRow++;
    scanned++;
  }
  if (isEmptyCell(ctx.getCellValue(startRow, col, sheetName))) return null;

  let endRow = startRow;
  scanned = 0;
  while (scanned < maxScanRows && !isEmptyCell(ctx.getCellValue(endRow + 1, col, sheetName))) {
    endRow++;
    scanned++;
  }

  const rawStartRow = startRow;
  const rawEndRow = endRow;

  const trimmed = trimNonNumericEdgesIfMostlyNumeric(ctx, col, startRow, endRow, sheetName);
  return {
    startRow: trimmed.startRow,
    endRow: trimmed.endRow,
    numericRatio: trimmed.numericRatio,
    rawStartRow,
    rawEndRow,
  };
}

function trimNonNumericEdgesIfMostlyNumeric(ctx, col, startRow, endRow, sheetName) {
  const stats = computeNumericStats(ctx, col, startRow, endRow, sheetName);
  // Heuristic: if the range is almost entirely numeric, treat non-numeric cells
  // at the edges as headers/footers and trim them (common in tables with a text
  // header row).
  if (stats.numericCount < 2) return stats;
  if (stats.numericCount < stats.totalCount - 2) return stats;

  let trimmedStart = startRow;
  let trimmedEnd = endRow;

  while (trimmedStart < trimmedEnd && !isNumericValue(ctx.getCellValue(trimmedStart, col, sheetName))) {
    trimmedStart++;
  }
  while (trimmedEnd > trimmedStart && !isNumericValue(ctx.getCellValue(trimmedEnd, col, sheetName))) {
    trimmedEnd--;
  }

  if (trimmedStart === startRow && trimmedEnd === endRow) return stats;
  return computeNumericStats(ctx, col, trimmedStart, trimmedEnd, sheetName);
}

function computeNumericStats(ctx, col, startRow, endRow, sheetName) {
  let numeric = 0;
  let total = 0;
  for (let r = startRow; r <= endRow; r++) {
    const v = ctx.getCellValue(r, col, sheetName);
    if (isEmptyCell(v)) continue;
    total++;
    if (isNumericValue(v)) numeric++;
  }
  return {
    startRow,
    endRow,
    numericCount: numeric,
    totalCount: total,
    numericRatio: total === 0 ? 0 : numeric / total,
  };
}

function isNumericValue(value) {
  if (typeof value === "number") return Number.isFinite(value);
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed === "") return false;
    return Number.isFinite(Number(trimmed));
  }
  return false;
}

function clamp01(v) {
  return Math.max(0, Math.min(1, v));
}

function applyColumnCase(letters, typedColToken) {
  if (!typedColToken) return letters;
  if (typedColToken.toUpperCase() === typedColToken) return letters.toUpperCase();
  if (typedColToken.toLowerCase() === typedColToken) return letters.toLowerCase();
  return letters;
}

/**
 * @param {CellContext} ctx
 * @param {number} startCol 0-based
 * @param {number} startRow 0-based
 * @param {number} endRow 0-based
 * @param {number} maxScanCols
 * @returns {{endCol:number, confidence:number} | null}
 */
function findContiguousTableToRight(ctx, startCol, startRow, endRow, maxScanCols, sheetName) {
  if (maxScanCols <= 1) return null;
  if (endRow < startRow) return null;
  if (startCol > EXCEL_MAX_COL_INDEX) return null;

  const rowCount = endRow - startRow + 1;
  if (rowCount <= 1) return null;

  const maxCol = Math.min(EXCEL_MAX_COL_INDEX, startCol + Math.max(1, maxScanCols) - 1);

  let endCol = startCol;

  // Confidence helpers.
  let headerNonEmptyCols = 0;
  let numericRatioSum = 0;

  // Include the base column in the stats.
  if (!isEmptyCell(ctx.getCellValue(startRow, startCol, sheetName))) headerNonEmptyCols++;
  numericRatioSum += computeTableColumnNumericRatio(ctx, startCol, startRow, endRow, sheetName);

  for (let col = startCol + 1; col <= maxCol; col++) {
    let nonEmpty = 0;
    for (let r = startRow; r <= endRow; r++) {
      const v = ctx.getCellValue(r, col, sheetName);
      if (!isEmptyCell(v)) nonEmpty++;
    }

    // Stop expanding when we hit an entirely empty column (gap).
    if (nonEmpty === 0) break;

    // Heuristic: require the column to be "table-shaped" (mostly filled) to avoid
    // pulling in stray values far to the right.
    const coverage = nonEmpty / rowCount;
    if (coverage < 0.6) break;

    endCol = col;

    if (!isEmptyCell(ctx.getCellValue(startRow, col, sheetName))) headerNonEmptyCols++;
    numericRatioSum += computeTableColumnNumericRatio(ctx, col, startRow, endRow, sheetName);
  }

  if (endCol === startCol) return null;

  const colCount = endCol - startCol + 1;
  const headerRatio = headerNonEmptyCols / colCount;
  const avgNumericRatio = numericRatioSum / colCount;

  // Confidence heuristic:
  // - Wide blocks are likely "tables" for lookup/filter functions.
  // - Header coverage is a good table signal.
  // - Numeric-heavy columns are common in analytic tables.
  const base = 0.45;
  const widthBoost = Math.min(0.25, (colCount - 1) * 0.05);
  const headerBoost = 0.1 * headerRatio;
  const lengthBoost = Math.min(0.1, rowCount / 200);
  const numericBoost = 0.1 * avgNumericRatio;
  const confidence = clamp01(base + widthBoost + headerBoost + lengthBoost + numericBoost);

  return { endCol, confidence };
}

function computeTableColumnNumericRatio(ctx, col, startRow, endRow, sheetName) {
  // Treat the top row as a potential header and exclude it from numeric stats so
  // numeric columns with a text header don't get penalized.
  if (endRow <= startRow) return 0;
  const stats = computeNumericStats(ctx, col, startRow + 1, endRow, sheetName);
  return stats.numericRatio;
}
