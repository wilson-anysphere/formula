import { normalizeRange, parseRangeA1 } from "../document/coords.js";
import { showToast } from "../extensions/ui.js";
import { applyStylePatch } from "./styleTable.js";
import { getStyleWrapText } from "./styleFieldAccess.js";

// Excel grid limits (used by the selection model and layered formatting fast paths).
const EXCEL_MAX_ROW = 1_048_576 - 1;
const EXCEL_MAX_COL = 16_384 - 1;
// Keep aligned with `apps/desktop/src/formatting/selectionSizeGuard.ts` default UI limit.
// This guard exists to prevent formatting operations from enumerating or materializing
// extremely large selections in JS.
const MAX_RANGE_FORMATTING_CELLS = 100_000;
const MAX_RANGE_FORMATTING_CELLS_LABEL = MAX_RANGE_FORMATTING_CELLS.toLocaleString();
// Full-width row formatting still requires enumerating each row in the selection
// (DocumentController row formatting layer). Align with `DEFAULT_FORMATTING_BAND_ROW_LIMIT`.
const MAX_RANGE_FORMATTING_BAND_ROWS = 50_000;

function ensureSafeFormattingRange(rangeOrRanges) {
  const showTooLargeToast = () => {
    try {
      showToast(
        `Selection too large to apply formatting (>${MAX_RANGE_FORMATTING_CELLS_LABEL} cells). Select fewer cells and try again.`,
        "warning",
      );
    } catch {
      // `showToast` requires a #toast-root (and DOM globals); ignore in non-UI contexts/tests.
    }
  };

  const ranges = normalizeRanges(rangeOrRanges);
  if (ranges.length === 0) return true;

  let totalCells = 0;
  let allRangesBand = true;

  for (const range of ranges) {
    const r = normalizeCellRange(range);

    const rows = r.end.row - r.start.row + 1;
    const cols = r.end.col - r.start.col + 1;
    totalCells += Math.max(0, rows) * Math.max(0, cols);

    // Full-width row selections are only scalable up to a row-count cap.
    // (This matches `DEFAULT_FORMATTING_BAND_ROW_LIMIT` / `DocumentController.setRangeFormat`.)
    const isFullWidthRows = r.start.col === 0 && r.end.col === EXCEL_MAX_COL;
    const isFullHeightCols = r.start.row === 0 && r.end.row === EXCEL_MAX_ROW;
    const isFullSheet = isFullWidthRows && r.start.row === 0 && r.end.row === EXCEL_MAX_ROW;
    if (isFullWidthRows && !isFullSheet) {
      if (rows > MAX_RANGE_FORMATTING_BAND_ROWS) {
        try {
          showToast("Selection is too large to format. Try selecting fewer rows or formatting the entire sheet.", "warning");
        } catch {
          // ignore (e.g. toast root missing in tests)
        }
        return false;
      }
    }

    const isBandRange = isFullSheet || isFullHeightCols || isFullWidthRows;
    if (!isBandRange) allRangesBand = false;
  }

  if (totalCells > MAX_RANGE_FORMATTING_CELLS && !allRangesBand) {
    showTooLargeToast();
    return false;
  }

  return true;
}

// For small selections, the simplest (and usually fastest) approach is to scan every cell.
// This threshold keeps behavior simple for typical edits while preventing catastrophic
// O(rows*cols) work for full-sheet / full-row / full-column selections.
const CELL_SCAN_THRESHOLD = 10_000;
const AXIS_ENUMERATION_LIMIT = 50_000;

function normalizeRanges(rangeOrRanges) {
  if (Array.isArray(rangeOrRanges)) return rangeOrRanges;
  return [rangeOrRanges];
}

function normalizeCellRange(range) {
  const parsed = typeof range === "string" ? parseRangeA1(range) : range;
  return normalizeRange(parsed);
}

function allCellsMatch(doc, sheetId, rangeOrRanges, predicate) {
  for (const range of normalizeRanges(rangeOrRanges)) {
    const r = normalizeCellRange(range);
    if (!allCellsMatchRange(doc, sheetId, r, predicate)) return false;
  }
  return true;
}

function parseRowColKey(key) {
  const [rowStr, colStr] = key.split(",");
  const row = Number(rowStr);
  const col = Number(colStr);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) return null;
  return { row, col };
}

function axisStyleIdCounts(axisMap, start, end) {
  const total = end - start + 1;
  if (total <= 0) return new Map([[0, 0]]);
  const size = axisMap?.size ?? 0;
  const canEnumerate = total <= AXIS_ENUMERATION_LIMIT && total <= size;

  /** @type {Map<number, number>} */
  const counts = new Map();

  if (canEnumerate) {
    for (let idx = start; idx <= end; idx++) {
      const styleId = axisMap?.get(idx) ?? 0;
      counts.set(styleId, (counts.get(styleId) ?? 0) + 1);
    }
    return counts;
  }

  // Sparse scan: count only explicit overrides, then infer the default count.
  let explicitCount = 0;
  if (axisMap && size > 0) {
    for (const [idx, styleId] of axisMap.entries()) {
      if (idx < start || idx > end) continue;
      if (!styleId) continue;
      counts.set(styleId, (counts.get(styleId) ?? 0) + 1);
      explicitCount += 1;
    }
  }

  const defaultCount = total - explicitCount;
  if (defaultCount > 0) counts.set(0, defaultCount);
  return counts;
}

function styleIdForRowInRuns(runs, row) {
  if (!Array.isArray(runs) || runs.length === 0) return 0;
  let lo = 0;
  let hi = runs.length - 1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const run = runs[mid];
    if (!run) return 0;
    if (row < run.startRow) {
      hi = mid - 1;
    } else if (row >= run.endRowExclusive) {
      lo = mid + 1;
    } else {
      return run.styleId ?? 0;
    }
  }
  return 0;
}

function runsOverlapRange(runs, startRow, endRowExclusive) {
  if (!Array.isArray(runs) || runs.length === 0) return false;
  for (const run of runs) {
    if (!run) continue;
    if (run.endRowExclusive <= startRow) continue;
    if (run.startRow >= endRowExclusive) break;
    return true;
  }
  return false;
}

function lowerBound(sorted, value) {
  let lo = 0;
  let hi = sorted.length;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (sorted[mid] < value) lo = mid + 1;
    else hi = mid;
  }
  return lo;
}

function allCellsMatchRange(doc, sheetId, range, predicate) {
  const rowCount = range.end.row - range.start.row + 1;
  const colCount = range.end.col - range.start.col + 1;
  const cellCount = rowCount * colCount;

  // Avoid materializing sheets for read-only toggle state checks.
  //
  // The DocumentController lazily creates sheets when referenced. Formatting toolbar helpers
  // are often called from UI state updates, so treat a missing sheet as having default
  // formatting instead of creating a "phantom" sheet.
  const modelForExistenceCheck = doc?.model;
  const sheetMapForExistenceCheck = modelForExistenceCheck?.sheets;
  if (sheetMapForExistenceCheck && typeof sheetMapForExistenceCheck.has === "function") {
    if (!sheetMapForExistenceCheck.has(sheetId)) {
      return predicate({});
    }
  }

  // Small rectangles: keep the simple per-cell semantics.
  if (cellCount <= CELL_SCAN_THRESHOLD) {
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) {
        const style = doc.getCellFormat(sheetId, { row, col });
        if (!predicate(style)) return false;
      }
    }
    return true;
  }

  // Layered-format fast path. This avoids iterating every coordinate for:
  // - full sheet selections
  // - full-height column selections
  // - full-width row selections
  //
  // It is also safe to use for other large rectangles: work scales with the number of
  // formatting overrides, not the area.
  const model = doc?.model;
  const styleTable = doc?.styleTable;
  const getCellFormat = doc?.getCellFormat;
  if (!model || !styleTable || typeof getCellFormat !== "function") {
    // Extremely defensive fallback: preserve semantics even if we can't access the model.
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) {
        const style = doc.getCellFormat(sheetId, { row, col });
        if (!predicate(style)) return false;
      }
    }
    return true;
  }

  const sheet = model.sheets?.get(sheetId);
  if (!sheet) {
    // No sheet means everything is default.
    return predicate({});
  }

  // Selection-shape sanity: use the same Excel limits as the controller.
  const isFullSheet =
    range.start.row === 0 &&
    range.end.row === EXCEL_MAX_ROW &&
    range.start.col === 0 &&
    range.end.col === EXCEL_MAX_COL;
  const isFullHeightCols = range.start.row === 0 && range.end.row === EXCEL_MAX_ROW;
  const isFullWidthRows = range.start.col === 0 && range.end.col === EXCEL_MAX_COL;

  // For very large selections that aren't an entire sheet/row/column, fall back to scanning
  // only if the range still isn't too large. (The layered-format scan below is still valid,
  // but this keeps behavior predictable for medium rectangles.)
  if (!isFullSheet && !isFullHeightCols && !isFullWidthRows && cellCount <= AXIS_ENUMERATION_LIMIT) {
    for (let row = range.start.row; row <= range.end.row; row++) {
      for (let col = range.start.col; col <= range.end.col; col++) {
        const style = doc.getCellFormat(sheetId, { row, col });
        if (!predicate(style)) return false;
      }
    }
    return true;
  }

  const sheetStyleId = sheet.defaultStyleId ?? 0;
  const sheetStyle = styleTable.get(sheetStyleId);

  const rowStyleIds = sheet.rowStyleIds ?? new Map();
  const colStyleIds = sheet.colStyleIds ?? new Map();
  const formatRunsByCol = sheet.formatRunsByCol ?? new Map();

  const rowCounts = axisStyleIdCounts(rowStyleIds, range.start.row, range.end.row);
  const colCounts = axisStyleIdCounts(colStyleIds, range.start.col, range.end.col);

  const startRow = range.start.row;
  const endRowExclusive = range.end.row + 1;

  /** @type {{ col: number, runs: any[] }[]} */
  const runCols = [];
  if (formatRunsByCol && typeof formatRunsByCol.get === "function") {
    // Iterate only selected columns (<= 16,384 in Excel space) rather than scanning every
    // column that happens to have runs in the sheet.
    for (let col = range.start.col; col <= range.end.col; col++) {
      const runs = formatRunsByCol.get(col) ?? null;
      if (!runsOverlapRange(runs, startRow, endRowExclusive)) continue;
      runCols.push({ col, runs });
    }
  }

  /** @type {Set<number>} */
  const runColSet = new Set(runCols.map((c) => c.col));
  /** @type {Map<number, number>} */
  const colCountsNoRun = new Map(colCounts);
  for (const { col } of runCols) {
    const colStyleId = colStyleIds.get(col) ?? 0;
    const prev = colCountsNoRun.get(colStyleId) ?? 0;
    if (prev <= 1) colCountsNoRun.delete(colStyleId);
    else colCountsNoRun.set(colStyleId, prev - 1);
  }

  /** @type {Map<number, number>} */
  const rowOverrideStyleByRow = new Map();
  /** @type {number[]} */
  const rowOverrideRows = [];
  if (rowStyleIds && rowStyleIds.size > 0) {
    for (const [row, styleId] of rowStyleIds.entries()) {
      if (row < range.start.row || row > range.end.row) continue;
      rowOverrideStyleByRow.set(row, styleId);
      rowOverrideRows.push(row);
    }
    rowOverrideRows.sort((a, b) => a - b);
  }

  /** @type {Map<string, number>} */
  const overriddenCellCountByNoRunRegion = new Map();
  /** @type {Map<number, number[]>} */
  const cellOverrideRowsByRunCol = new Map();

  const sheetColCache = new Map();
  const sheetColRowCache = new Map();
  const baseStyleCache = new Map();
  const basePredicateCache = new Map();
  const cellPredicateCache = new Map();

  const sheetColStyle = (colStyleId) => {
    const cached = sheetColCache.get(colStyleId);
    if (cached) return cached;
    const merged = applyStylePatch(sheetStyle, styleTable.get(colStyleId));
    sheetColCache.set(colStyleId, merged);
    return merged;
  };

  const sheetColRowStyle = (colStyleId, rowStyleId) => {
    const key = `${colStyleId}|${rowStyleId}`;
    const cached = sheetColRowCache.get(key);
    if (cached) return cached;
    const merged = applyStylePatch(sheetColStyle(colStyleId), styleTable.get(rowStyleId));
    sheetColRowCache.set(key, merged);
    return merged;
  };

  const baseStyle = (colStyleId, rowStyleId, runStyleId) => {
    const key = `${colStyleId}|${rowStyleId}|${runStyleId}`;
    const cached = baseStyleCache.get(key);
    if (cached) return cached;
    const merged = applyStylePatch(sheetColRowStyle(colStyleId, rowStyleId), styleTable.get(runStyleId));
    baseStyleCache.set(key, merged);
    return merged;
  };

  // 1) Check explicit cell-level overrides inside the selection.
  //    These always win over all other layers.
  const boundsAreReliable = !sheet.formatBoundsDirty;
  const storedBounds = boundsAreReliable ? sheet.formatBounds ?? null : null;
  const selectionIntersectsStoredBounds =
    !storedBounds ||
    (storedBounds.endRow >= range.start.row &&
      storedBounds.startRow <= range.end.row &&
      storedBounds.endCol >= range.start.col &&
      storedBounds.startCol <= range.end.col);

  const styledKeys = sheet.styledCells;
  const styledCellsByRow = sheet.styledCellsByRow;
  const styledCellsByCol = sheet.styledCellsByCol;
  const hasAxisIndex =
    styledCellsByRow &&
    typeof styledCellsByRow.get === "function" &&
    styledCellsByCol &&
    typeof styledCellsByCol.get === "function";

  if (selectionIntersectsStoredBounds && hasAxisIndex) {
    const rowCount = range.end.row - range.start.row + 1;
    const colCount = range.end.col - range.start.col + 1;

    // Prefer iterating the smaller axis to keep work proportional to the number of
    // intersecting cell overrides rather than total overrides in the sheet.
    const iterateRows = rowCount <= colCount;

    if (iterateRows && rowCount <= AXIS_ENUMERATION_LIMIT) {
      for (let row = range.start.row; row <= range.end.row; row++) {
        const cols = styledCellsByRow.get(row);
        if (!cols || cols.size === 0) continue;
        for (const col of cols) {
          if (col < range.start.col || col > range.end.col) continue;
          const cell = sheet.cells.get(`${row},${col}`);
          if (!cell || cell.styleId === 0) continue;

          const rowStyleId = rowStyleIds.get(row) ?? 0;
          const colStyleId = colStyleIds.get(col) ?? 0;
          const runStyleId = styleIdForRowInRuns(formatRunsByCol.get(col), row);

          if (runColSet.has(col)) {
            let rows = cellOverrideRowsByRunCol.get(col);
            if (!rows) {
              rows = [];
              cellOverrideRowsByRunCol.set(col, rows);
            }
            rows.push(row);
          } else {
            const regionKey = `${colStyleId}|${rowStyleId}`;
            overriddenCellCountByNoRunRegion.set(regionKey, (overriddenCellCountByNoRunRegion.get(regionKey) ?? 0) + 1);
          }

          const cellKey = `${colStyleId}|${rowStyleId}|${runStyleId}|${cell.styleId}`;
          const cachedMatch = cellPredicateCache.get(cellKey);
          if (cachedMatch === false) return false;
          if (cachedMatch === true) continue;

          const merged = applyStylePatch(baseStyle(colStyleId, rowStyleId, runStyleId), styleTable.get(cell.styleId));
          const matches = Boolean(predicate(merged));
          cellPredicateCache.set(cellKey, matches);
          if (!matches) return false;
        }
      }
    } else {
      // Column iteration is always safe (Excel max 16,384 columns).
      for (let col = range.start.col; col <= range.end.col; col++) {
        const rows = styledCellsByCol.get(col);
        if (!rows || rows.size === 0) continue;
        for (const row of rows) {
          if (row < range.start.row || row > range.end.row) continue;
          const cell = sheet.cells.get(`${row},${col}`);
          if (!cell || cell.styleId === 0) continue;

          const rowStyleId = rowStyleIds.get(row) ?? 0;
          const colStyleId = colStyleIds.get(col) ?? 0;
          const runStyleId = styleIdForRowInRuns(formatRunsByCol.get(col), row);

          if (runColSet.has(col)) {
            let rowList = cellOverrideRowsByRunCol.get(col);
            if (!rowList) {
              rowList = [];
              cellOverrideRowsByRunCol.set(col, rowList);
            }
            rowList.push(row);
          } else {
            const regionKey = `${colStyleId}|${rowStyleId}`;
            overriddenCellCountByNoRunRegion.set(regionKey, (overriddenCellCountByNoRunRegion.get(regionKey) ?? 0) + 1);
          }

          const cellKey = `${colStyleId}|${rowStyleId}|${runStyleId}|${cell.styleId}`;
          const cachedMatch = cellPredicateCache.get(cellKey);
          if (cachedMatch === false) return false;
          if (cachedMatch === true) continue;

          const merged = applyStylePatch(baseStyle(colStyleId, rowStyleId, runStyleId), styleTable.get(cell.styleId));
          const matches = Boolean(predicate(merged));
          cellPredicateCache.set(cellKey, matches);
          if (!matches) return false;
        }
      }
    }
  }
  // Backward-compatible fallback (older sheet encodings without `styledCells`).
  else if (selectionIntersectsStoredBounds && styledKeys && typeof styledKeys[Symbol.iterator] === "function") {
    for (const key of styledKeys) {
      const cell = sheet.cells.get(key);
      if (!cell || cell.styleId === 0) continue;
      const coord = parseRowColKey(key);
      if (!coord) continue;
      const { row, col } = coord;
      if (row < range.start.row || row > range.end.row) continue;
      if (col < range.start.col || col > range.end.col) continue;

      const rowStyleId = rowStyleIds.get(row) ?? 0;
      const colStyleId = colStyleIds.get(col) ?? 0;
      const runStyleId = styleIdForRowInRuns(formatRunsByCol.get(col), row);

      if (runColSet.has(col)) {
        let rows = cellOverrideRowsByRunCol.get(col);
        if (!rows) {
          rows = [];
          cellOverrideRowsByRunCol.set(col, rows);
        }
        rows.push(row);
      } else {
        const regionKey = `${colStyleId}|${rowStyleId}`;
        overriddenCellCountByNoRunRegion.set(regionKey, (overriddenCellCountByNoRunRegion.get(regionKey) ?? 0) + 1);
      }

      const cellKey = `${colStyleId}|${rowStyleId}|${runStyleId}|${cell.styleId}`;
      const cachedMatch = cellPredicateCache.get(cellKey);
      if (cachedMatch === false) return false;
      if (cachedMatch === true) continue;

      const merged = applyStylePatch(baseStyle(colStyleId, rowStyleId, runStyleId), styleTable.get(cell.styleId));
      const matches = Boolean(predicate(merged));
      cellPredicateCache.set(cellKey, matches);
      if (!matches) return false;
    }
  }
  // Backward-compatible fallback (older sheet encodings without `styledCells`).
  else if (selectionIntersectsStoredBounds && sheet.cells && sheet.cells.size > 0) {
    for (const [key, cell] of sheet.cells.entries()) {
      if (!cell || cell.styleId === 0) continue;
      const coord = parseRowColKey(key);
      if (!coord) continue;
      const { row, col } = coord;
      if (row < range.start.row || row > range.end.row) continue;
      if (col < range.start.col || col > range.end.col) continue;

      const rowStyleId = rowStyleIds.get(row) ?? 0;
      const colStyleId = colStyleIds.get(col) ?? 0;
      const runStyleId = styleIdForRowInRuns(formatRunsByCol.get(col), row);

      if (runColSet.has(col)) {
        let rows = cellOverrideRowsByRunCol.get(col);
        if (!rows) {
          rows = [];
          cellOverrideRowsByRunCol.set(col, rows);
        }
        rows.push(row);
      } else {
        const regionKey = `${colStyleId}|${rowStyleId}`;
        overriddenCellCountByNoRunRegion.set(regionKey, (overriddenCellCountByNoRunRegion.get(regionKey) ?? 0) + 1);
      }

      const cellKey = `${colStyleId}|${rowStyleId}|${runStyleId}|${cell.styleId}`;
      const cachedMatch = cellPredicateCache.get(cellKey);
      if (cachedMatch === false) return false;
      if (cachedMatch === true) continue;

      const merged = applyStylePatch(baseStyle(colStyleId, rowStyleId, runStyleId), styleTable.get(cell.styleId));
      const matches = Boolean(predicate(merged));
      cellPredicateCache.set(cellKey, matches);
      if (!matches) return false;
    }
  }

  /** @type {Map<number, Set<number>>} */
  const cellOverrideRowSetByRunCol = new Map();
  for (const [col, rows] of cellOverrideRowsByRunCol.entries()) {
    rows.sort((a, b) => a - b);
    cellOverrideRowSetByRunCol.set(col, new Set(rows));
  }

  // 2) Check base styles for columns WITHOUT applicable range runs in the selection.
  //    Effective precedence for these cells is: sheet < col < row (run=0, cell=0).
  for (const [rowStyleId, rowsWithStyle] of rowCounts.entries()) {
    for (const [colStyleId, colsWithStyle] of colCountsNoRun.entries()) {
      const regionCellCount = rowsWithStyle * colsWithStyle;
      if (regionCellCount <= 0) continue;

      const regionKey = `${colStyleId}|${rowStyleId}`;
      const overriddenCount = overriddenCellCountByNoRunRegion.get(regionKey) ?? 0;
      if (overriddenCount >= regionCellCount) continue;

      const cacheKey = `${colStyleId}|${rowStyleId}|0`;
      const cached = basePredicateCache.get(cacheKey);
      if (cached === false) return false;
      if (cached === true) continue;

      const matches = Boolean(predicate(baseStyle(colStyleId, rowStyleId, 0)));
      basePredicateCache.set(cacheKey, matches);
      if (!matches) return false;
    }
  }

  // 3) Check base styles for columns WITH range runs intersecting the selection.
  //    Effective precedence for these cells is: sheet < col < row < run (cell=0).
  for (const { col, runs } of runCols) {
    const colStyleId = colStyleIds.get(col) ?? 0;
    const overriddenRows = cellOverrideRowsByRunCol.get(col) ?? [];
    const overriddenRowSet = cellOverrideRowSetByRunCol.get(col) ?? new Set();

    let rowOverrideIdx = rowOverrideRows.length > 0 ? lowerBound(rowOverrideRows, startRow) : 0;
    let cellOverrideIdx = overriddenRows.length > 0 ? lowerBound(overriddenRows, startRow) : 0;

    const evalSegment = (segStart, segEnd, runStyleId) => {
      const segLen = segEnd - segStart;
      if (segLen <= 0) return true;

      // Advance pointers to segment start.
      while (rowOverrideIdx < rowOverrideRows.length && rowOverrideRows[rowOverrideIdx] < segStart) rowOverrideIdx += 1;
      while (cellOverrideIdx < overriddenRows.length && overriddenRows[cellOverrideIdx] < segStart) cellOverrideIdx += 1;

      // Row-level overrides (if their cell isn't explicitly overridden).
      const rowOverrideStartIdx = rowOverrideIdx;
      while (rowOverrideIdx < rowOverrideRows.length && rowOverrideRows[rowOverrideIdx] < segEnd) {
        const row = rowOverrideRows[rowOverrideIdx];
        const rowStyleId = rowOverrideStyleByRow.get(row) ?? 0;
        if (rowStyleId !== 0 && !overriddenRowSet.has(row)) {
          const cacheKey = `${colStyleId}|${rowStyleId}|${runStyleId}`;
          const cached = basePredicateCache.get(cacheKey);
          if (cached === false) return false;
          if (cached !== true) {
            const matches = Boolean(predicate(baseStyle(colStyleId, rowStyleId, runStyleId)));
            basePredicateCache.set(cacheKey, matches);
            if (!matches) return false;
          }
        }
        rowOverrideIdx += 1;
      }

      const rowOverrideCount = rowOverrideIdx - rowOverrideStartIdx;
      const defaultRowCount = segLen - rowOverrideCount;
      if (defaultRowCount <= 0) return true;

      // Count cell overrides that live on default rows (rowStyleId=0) within this segment.
      let defaultCellOverrideCount = 0;
      while (cellOverrideIdx < overriddenRows.length && overriddenRows[cellOverrideIdx] < segEnd) {
        const row = overriddenRows[cellOverrideIdx];
        if (!rowOverrideStyleByRow.has(row)) defaultCellOverrideCount += 1;
        cellOverrideIdx += 1;
      }

      // No remaining cells use the base style for rowStyleId=0 in this segment.
      if (defaultRowCount <= defaultCellOverrideCount) return true;

      const cacheKey = `${colStyleId}|0|${runStyleId}`;
      const cached = basePredicateCache.get(cacheKey);
      if (cached === false) return false;
      if (cached === true) return true;

      const matches = Boolean(predicate(baseStyle(colStyleId, 0, runStyleId)));
      basePredicateCache.set(cacheKey, matches);
      return matches;
    };

    let cursor = startRow;
    for (const run of runs) {
      if (!run) continue;
      if (run.endRowExclusive <= startRow) continue;
      if (run.startRow >= endRowExclusive) break;

      const runStart = Math.max(run.startRow, startRow);
      const runEnd = Math.min(run.endRowExclusive, endRowExclusive);

      // Gap before this run.
      if (cursor < runStart) {
        if (!evalSegment(cursor, runStart, 0)) return false;
      }

      // Overlapping run segment.
      if (runStart < runEnd) {
        if (!evalSegment(runStart, runEnd, run.styleId ?? 0)) return false;
      }

      cursor = Math.max(cursor, runEnd);
      if (cursor >= endRowExclusive) break;
    }

    // Trailing gap after the last run.
    if (cursor < endRowExclusive) {
      if (!evalSegment(cursor, endRowExclusive, 0)) return false;
    }
  }

  return true;
}

export function toggleBold(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean" ? options.next : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.bold));
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { bold: next } }, { label: "Bold" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function toggleItalic(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.italic));
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { italic: next } }, { label: "Italic" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function toggleUnderline(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.underline));
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { underline: next } }, { label: "Underline" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function toggleStrikethrough(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => {
          // Excel/OOXML styles can encode strike at either `font.strike` or top-level `strike`.
          // Prefer the font-level value when present (it overrides the legacy top-level field).
          const fontStrike = s?.font?.strike;
          if (typeof fontStrike === "boolean") return fontStrike;
          const strike = s?.strike;
          if (typeof strike === "boolean") return strike;
          return false;
        });
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { strike: next } }, { label: "Strikethrough" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function toggleSubscript(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => {
          const raw = s?.font?.vertAlign;
          return typeof raw === "string" && raw.toLowerCase() === "subscript";
        });
  let applied = true;
  // Excel semantics: subscript and superscript are mutually exclusive. Setting one clears the other.
  const vertAlign = next ? "subscript" : null;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { vertAlign } }, { label: "Subscript" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function toggleSuperscript(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => {
          const raw = s?.font?.vertAlign;
          return typeof raw === "string" && raw.toLowerCase() === "superscript";
        });
  let applied = true;
  // Excel semantics: subscript and superscript are mutually exclusive. Setting one clears the other.
  const vertAlign = next ? "superscript" : null;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { vertAlign } }, { label: "Superscript" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function setFontSize(doc, sheetId, range, sizePt) {
  if (!ensureSafeFormattingRange(range)) return false;
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { size: sizePt } }, { label: "Font size" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function setFontColor(doc, sheetId, range, argb) {
  if (!ensureSafeFormattingRange(range)) return false;
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { font: { color: argb } }, { label: "Font color" });
    if (ok === false) applied = false;
  }
  return applied;
}

export function setFillColor(doc, sheetId, range, argb) {
  if (!ensureSafeFormattingRange(range)) return false;
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(
      sheetId,
      r,
      { fill: { pattern: "solid", fgColor: argb } },
      { label: "Fill color" },
    );
    if (ok === false) applied = false;
  }
  return applied;
}

const DEFAULT_BORDER_ARGB = "FF000000";

export function applyAllBorders(doc, sheetId, range, { style = "thin", color } = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const resolvedColor = color ?? `#${DEFAULT_BORDER_ARGB}`;
  const edge = { style, color: resolvedColor };
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(
      sheetId,
      r,
      { border: { left: edge, right: edge, top: edge, bottom: edge } },
      { label: "Borders" },
    );
    if (ok === false) applied = false;
  }
  return applied;
}

export function applyOutsideBorders(doc, sheetId, range, { style = "thin", color } = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const resolvedColor = color ?? `#${DEFAULT_BORDER_ARGB}`;
  const edge = { style, color: resolvedColor };
  const shouldBatch = doc?.batchDepth === 0;
  if (shouldBatch) doc.beginBatch({ label: "Borders" });
  let applied = true;
  try {
    for (const r of normalizeRanges(range)) {
      const rect = normalizeCellRange(r);
      const startRow = rect.start.row;
      const endRow = rect.end.row;
      const startCol = rect.start.col;
      const endCol = rect.end.col;

      const okTop = doc.setRangeFormat(
        sheetId,
        { start: { row: startRow, col: startCol }, end: { row: startRow, col: endCol } },
        { border: { top: edge } },
        { label: "Borders" },
      );
      if (okTop === false) applied = false;

      const okBottom = doc.setRangeFormat(
        sheetId,
        { start: { row: endRow, col: startCol }, end: { row: endRow, col: endCol } },
        { border: { bottom: edge } },
        { label: "Borders" },
      );
      if (okBottom === false) applied = false;

      const okLeft = doc.setRangeFormat(
        sheetId,
        { start: { row: startRow, col: startCol }, end: { row: endRow, col: startCol } },
        { border: { left: edge } },
        { label: "Borders" },
      );
      if (okLeft === false) applied = false;

      const okRight = doc.setRangeFormat(
        sheetId,
        { start: { row: startRow, col: endCol }, end: { row: endRow, col: endCol } },
        { border: { right: edge } },
        { label: "Borders" },
      );
      if (okRight === false) applied = false;
    }
    return applied;
  } catch (err) {
    if (shouldBatch) {
      try {
        doc.cancelBatch();
      } catch {
        // ignore
      }
    }
    throw err;
  } finally {
    if (shouldBatch) doc.endBatch();
  }
}

export function setHorizontalAlign(doc, sheetId, range, align) {
  if (!ensureSafeFormattingRange(range)) return false;
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(
      sheetId,
      r,
      { alignment: { horizontal: align } },
      { label: "Horizontal align" },
    );
    if (ok === false) applied = false;
  }
  return applied;
}

export function toggleWrap(doc, sheetId, range, options = {}) {
  if (!ensureSafeFormattingRange(range)) return false;
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => getStyleWrapText(s));
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { alignment: { wrapText: next } }, { label: "Wrap" });
    if (ok === false) applied = false;
  }
  return applied;
}

export const NUMBER_FORMATS = {
  currency: "$#,##0.00",
  percent: "0%",
  date: "m/d/yyyy",
};

export function applyNumberFormatPreset(doc, sheetId, range, preset) {
  if (!ensureSafeFormattingRange(range)) return false;
  const code = NUMBER_FORMATS[preset];
  if (!code) throw new Error(`Unknown number format preset: ${preset}`);
  let applied = true;
  for (const r of normalizeRanges(range)) {
    const ok = doc.setRangeFormat(sheetId, r, { numberFormat: code }, { label: "Number format" });
    if (ok === false) applied = false;
  }
  return applied;
}
