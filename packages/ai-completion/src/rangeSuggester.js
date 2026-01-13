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
 * - If no block exists above the active row (e.g. formula entered above a data block),
 *   fall back to scanning downward for the first contiguous non-empty run.
 * - If the active cell is in a different column (e.g. formula in B5 referencing A),
 *   treat the current row as part of the scan window and (within `maxScanRows`) extend
 *   downward to capture the full contiguous block around that row.
 * - If adjacent columns contain a similarly-filled block over the same row span,
 *   also suggest a 2D rectangular range (A1:D10). Expansion stops at the first
 *   entirely-empty “gap” column and is bounded by `maxScanCols`.
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
  /** @type {{row:number,col:number} | null} */
  let cellRef = null;
  try {
    cellRef = normalizeCellRef(params.cellRef);
  } catch {
    return [];
  }
  const sheetName = params.sheetName;
  const maxScanRows = params.maxScanRows ?? 500;
  const maxScanCols = params.maxScanCols ?? 50;

  if (!surroundingCells || typeof surroundingCells.getCellValue !== "function") {
    return [];
  }

  // If we don't know where the user is editing (or the ref is clearly invalid),
  // avoid emitting range suggestions. The heuristics in this module are based
  // on scanning "near" the active cell.
  if (!Number.isInteger(cellRef.row) || !Number.isInteger(cellRef.col) || cellRef.row < 0 || cellRef.col < 0) {
    return [];
  }

  let arg = (currentArgText ?? "").trim();

  // Empty argument ranges (e.g. "=SUM(") - still provide suggestions. Use the
  // active cell column as the default so the completion remains a pure insertion.
  if (arg.length === 0) {
    if (cellRef.col > EXCEL_MAX_COL_INDEX) return [];
    arg = columnIndexToLetter(cellRef.col);
  }

  // Partial column-range token (A:) -> suggest A:A. Avoid suggesting A1:A10 since
  // that would require inserting characters *before* the typed ':' (not a pure insertion).
  const colRangePrefix = /^(\$?)([A-Za-z]{1,3}):$/.exec(arg);
  if (colRangePrefix) {
    const colPrefix = colRangePrefix[1] === "$" ? "$" : "";
    const colToken = colRangePrefix[2];
    const colIndex = safeColumnLetterToIndex(colToken.toUpperCase());
    if (colIndex === null) return [];
    return [
      {
        range: `${colPrefix}${colToken}:${colPrefix}${colToken}`,
        confidence: 0.35,
        reason: "entire_column",
      },
    ];
  }

  // Only handle conservative A1-style column/cell prefixes, plus partial range syntax:
  // - A / A1 / $A$1
  // - A: / A1:
  // - A:A / A1:A
  const parsed = parseColumnRangePrefix(arg);
  if (!parsed) return [];

  const colIndex = safeColumnLetterToIndex(parsed.colLetters);
  if (colIndex === null) return [];

  const explicitRow = parsed.explicitRow;
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
    // When the active cell isn't in the referenced column, include the current row
    // in the scan. This enables "in-block" formulas like B5 =SUM(A) to capture the
    // full contiguous column block around that row, rather than only the portion above.
    const aboveFromRow = cellRef.row - (cellRef.col === colIndex ? 1 : 0);
    contiguous = findContiguousBlockAbove(surroundingCells, colIndex, aboveFromRow, maxScanRows, sheetName);
    contiguousReason = "contiguous_above_current_cell";
    // If we included the current row in the scan (active cell is in a different column),
    // extend the block downward so we capture the full contiguous run.
    if (contiguous && cellRef.col !== colIndex) {
      const rawStartRow = contiguous.rawStartRow ?? contiguous.startRow;
      const rawEndRow = contiguous.rawEndRow ?? contiguous.endRow;
      const scannedSpan = Math.max(0, rawEndRow - rawStartRow + 1);
      const remainingRows = Math.max(0, maxScanRows - scannedSpan);
      contiguous = extendContiguousBlockDown(
        surroundingCells,
        colIndex,
        rawStartRow,
        rawEndRow,
        remainingRows,
        sheetName,
      );
    }
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
    const startA1 = `${parsed.startColPrefix}${parsed.startColToken}${parsed.rowPrefix}${startRow + 1}`;
    const endA1 = `${parsed.endColPrefix}${parsed.endColToken}${parsed.rowPrefix}${endRow + 1}`;
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
      // Use the end-token casing when expanding the right edge so mixed-case inputs
      // like "AB1:a" keep their casing consistent with the 1D range suggestion.
      const endLetters = applyColumnCase(columnIndexToLetter(endCol), parsed.endColToken);
      const startCell = `${parsed.startColPrefix}${parsed.startColToken}${parsed.rowPrefix}${tableStartRow + 1}`;
      const endCell = `${parsed.endColPrefix}${endLetters}${parsed.rowPrefix}${tableEndRow + 1}`;
      const tableReason = toTableReason(contiguousReason);
      tableSuggestion = {
        range: `${startCell}:${endCell}`,
        confidence,
        reason: tableReason,
      };
    }
  }

  suggestions.push({
    range: `${parsed.startColPrefix}${parsed.startColToken}:${parsed.endColPrefix}${parsed.endColToken}`,
    confidence: 0.3,
    reason: "entire_column",
  });

  // Keep "entire column" as the second suggestion for compatibility with existing
  // tests/callers, then append the optional 2D table range.
  if (tableSuggestion) suggestions.push(tableSuggestion);

  return suggestions;
}

const EXCEL_MAX_COL_INDEX = 16383; // XFD (1-based 16384)

/**
 * Parse a conservative subset of A1 range-prefix syntax.
 *
 * Supported forms:
 * - A / A1 / $A$1
 * - A: / A1:
  * - A:A / A1:A
  *
  * Notes:
  * - We only support ranges within a single column. If the end column prefix can't be
  *   the start column (e.g. "A1:B"), return null to avoid suggesting wrong 2D ranges.
  * - If a partial range has no explicit end token (e.g. "A1:"), the end column inherits
  *   the start column token/prefix (matching the existing single-token behavior).
 *
 * @param {string} arg
 * @returns {{
 *   colLetters: string,
 *   explicitRow: number | null,
 *   rowPrefix: string,
 *   startColPrefix: string,
 *   startColToken: string,
 *   endColPrefix: string,
 *   endColToken: string,
 * } | null}
 */
function parseColumnRangePrefix(arg) {
  if (typeof arg !== "string") return null;
  const text = arg.trim();
  if (text.length === 0) return null;

  const firstColon = text.indexOf(":");
  if (firstColon === -1) {
    const start = parseA1PrefixToken(text);
    if (!start) return null;
    return {
      colLetters: start.colLetters,
      explicitRow: start.explicitRow,
      rowPrefix: start.rowPrefix,
      startColPrefix: start.colPrefix,
      startColToken: start.colToken,
      endColPrefix: start.colPrefix,
      endColToken: start.colToken,
    };
  }

  // Be conservative: only handle a single colon.
  if (text.indexOf(":", firstColon + 1) !== -1) return null;

  const left = text.slice(0, firstColon);
  const right = text.slice(firstColon + 1);
  if (left.length === 0) return null;

  const start = parseA1PrefixToken(left);
  if (!start) return null;

  // "A1:" (no end token yet) - treat as a range anchored at the start column.
  if (right.length === 0) {
    return {
      colLetters: start.colLetters,
      explicitRow: start.explicitRow,
      rowPrefix: start.rowPrefix,
      startColPrefix: start.colPrefix,
      startColToken: start.colToken,
      endColPrefix: start.colPrefix,
      endColToken: start.colToken,
    };
  }

  // End token is column-only (no row digits) so "A1:A" / "$A:$A" works.
  const end = parseColumnToken(right);
  if (!end) return null;

  // Only support single-column ranges. Allow partial end-column prefixes like:
  // - AB1:A  (user has typed the first letter of the end column, but not the full AB yet)
  // - AA1:A  (prefix of AA)
  if (!start.colLetters.startsWith(end.colLetters)) return null;

  const endColToken = completeColumnToken(end.colToken, start.colLetters);
  if (!endColToken) return null;

  return {
    colLetters: start.colLetters,
    explicitRow: start.explicitRow,
    rowPrefix: start.rowPrefix,
    startColPrefix: start.colPrefix,
    startColToken: start.colToken,
    endColPrefix: end.colPrefix,
    endColToken,
  };
}

/**
 * Complete a partially typed end-column token (e.g. "A") to a full column token (e.g. "AB")
 * using the known start column letters.
 *
 * This is intentionally conservative: if the typed token isn't a prefix of the full
 * column letters, return null.
 *
 * @param {string} typedToken Column letters as typed by the user (no $ prefix).
 * @param {string} fullColLetters Canonical uppercase column letters (e.g. "AB").
 * @returns {string | null}
 */
function completeColumnToken(typedToken, fullColLetters) {
  if (typeof typedToken !== "string" || typeof fullColLetters !== "string") return null;
  if (typedToken.length === 0 || fullColLetters.length === 0) return null;

  const typedUpper = typedToken.toUpperCase();
  if (!fullColLetters.startsWith(typedUpper)) return null;
  if (typedToken.length >= fullColLetters.length) return typedToken;

  let remainder = fullColLetters.slice(typedToken.length);
  if (typedToken === typedUpper) remainder = remainder.toUpperCase();
  else if (typedToken === typedToken.toLowerCase()) remainder = remainder.toLowerCase();
  return `${typedToken}${remainder}`;
}

/**
 * Parse an A1 column/cell prefix token (no sheet name; no range colon).
 *
 * Examples:
 * - A
 * - $A
 * - A1
 * - A$1
 * - $A$1
 *
 * @param {string} token
 * @returns {{
 *   colPrefix: string,
 *   colToken: string,
 *   colLetters: string,
 *   rowPrefix: string,
 *   explicitRow: number | null,
 * } | null}
 */
function parseA1PrefixToken(token) {
  if (typeof token !== "string") return null;
  const match = /^(\$?)([A-Za-z]{1,3})(?:(\$?)(\d+))?$/.exec(token.trim());
  if (!match) return null;

  const colPrefix = match[1] === "$" ? "$" : "";
  const colToken = match[2];
  const colLetters = colToken.toUpperCase();

  const rowPrefix = match[3] === "$" ? "$" : "";
  const explicitRow = match[4] ? Number(match[4]) : null;
  if (explicitRow !== null && (!Number.isInteger(explicitRow) || explicitRow <= 0)) return null;

  return { colPrefix, colToken, colLetters, rowPrefix, explicitRow };
}

/**
 * Parse a column-only reference (no row digits).
 *
 * Examples:
 * - A
 * - $A
 *
 * @param {string} token
 * @returns {{ colPrefix: string, colToken: string, colLetters: string } | null}
 */
function parseColumnToken(token) {
  if (typeof token !== "string") return null;
  const match = /^(\$?)([A-Za-z]{1,3})$/.exec(token.trim());
  if (!match) return null;

  const colPrefix = match[1] === "$" ? "$" : "";
  const colToken = match[2];
  const colLetters = colToken.toUpperCase();
  return { colPrefix, colToken, colLetters };
}

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
  // If we hit the scan cap without ever finding a non-empty cell, treat it as "no signal".
  if (isEmptyCell(ctx.getCellValue(endRow, col, sheetName))) return null;

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

function extendContiguousBlockDown(ctx, col, startRow, endRow, maxExtendRows, sheetName) {
  if (startRow < 0) return null;
  if (!Number.isInteger(startRow) || !Number.isInteger(endRow) || endRow < startRow) return null;

  let rawEndRow = endRow;
  let scanned = 0;
  while (scanned < maxExtendRows && !isEmptyCell(ctx.getCellValue(rawEndRow + 1, col, sheetName))) {
    rawEndRow++;
    scanned++;
  }

  const rawStartRow = startRow;
  const trimmed = trimNonNumericEdgesIfMostlyNumeric(ctx, col, rawStartRow, rawEndRow, sheetName);
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

function toTableReason(contiguousReason) {
  switch (contiguousReason) {
    case "contiguous_down_from_start":
      return "contiguous_table_down_from_start";
    case "contiguous_below_current_cell":
      return "contiguous_table_below_current_cell";
    case "contiguous_above_current_cell":
    default:
      return "contiguous_table_above_current_cell";
  }
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
