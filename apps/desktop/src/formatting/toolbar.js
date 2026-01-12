import { normalizeRange, parseRangeA1 } from "../document/coords.js";
import { applyStylePatch } from "./styleTable.js";

// Excel grid limits (used by the selection model and layered formatting fast paths).
const EXCEL_MAX_ROW = 1_048_576 - 1;
const EXCEL_MAX_COL = 16_384 - 1;

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

function allCellsMatchRange(doc, sheetId, range, predicate) {
  const rowCount = range.end.row - range.start.row + 1;
  const colCount = range.end.col - range.start.col + 1;
  const cellCount = rowCount * colCount;

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

  // Ensure the sheet exists (DocumentController is lazily materialized).
  if (typeof model.getCell === "function") {
    model.getCell(sheetId, 0, 0);
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

  const rowCounts = axisStyleIdCounts(rowStyleIds, range.start.row, range.end.row);
  const colCounts = axisStyleIdCounts(colStyleIds, range.start.col, range.end.col);

  /** @type {Map<string, number>} */
  const overriddenCellCountByRegion = new Map();

  const sheetColCache = new Map();
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

  const baseStyle = (colStyleId, rowStyleId) => {
    const key = `${colStyleId}|${rowStyleId}`;
    const cached = baseStyleCache.get(key);
    if (cached) return cached;
    const merged = applyStylePatch(sheetColStyle(colStyleId), styleTable.get(rowStyleId));
    baseStyleCache.set(key, merged);
    return merged;
  };

  // 1) Check explicit cell-level overrides inside the selection and track which base regions
  //    still have at least one non-overridden cell.
  if (sheet.cells && sheet.cells.size > 0) {
    for (const [key, cell] of sheet.cells.entries()) {
      if (!cell || cell.styleId === 0) continue;
      const coord = parseRowColKey(key);
      if (!coord) continue;
      const { row, col } = coord;
      if (row < range.start.row || row > range.end.row) continue;
      if (col < range.start.col || col > range.end.col) continue;

      const rowStyleId = rowStyleIds.get(row) ?? 0;
      const colStyleId = colStyleIds.get(col) ?? 0;

      const regionKey = `${colStyleId}|${rowStyleId}`;
      overriddenCellCountByRegion.set(regionKey, (overriddenCellCountByRegion.get(regionKey) ?? 0) + 1);

      const cellKey = `${colStyleId}|${rowStyleId}|${cell.styleId}`;
      const cachedMatch = cellPredicateCache.get(cellKey);
      if (cachedMatch === false) return false;
      if (cachedMatch === true) continue;

      const merged = applyStylePatch(baseStyle(colStyleId, rowStyleId), styleTable.get(cell.styleId));
      const matches = Boolean(predicate(merged));
      cellPredicateCache.set(cellKey, matches);
      if (!matches) return false;
    }
  }

  // 2) Check base styles (sheet/col/row) for any region that contains at least one cell not
  //    overridden by a cell-level style.
  for (const [rowStyleId, rowsWithStyle] of rowCounts.entries()) {
    for (const [colStyleId, colsWithStyle] of colCounts.entries()) {
      const regionCellCount = rowsWithStyle * colsWithStyle;
      if (regionCellCount <= 0) continue;

      const regionKey = `${colStyleId}|${rowStyleId}`;
      const overriddenCount = overriddenCellCountByRegion.get(regionKey) ?? 0;
      if (overriddenCount >= regionCellCount) continue;

      const cacheKey = regionKey;
      const cached = basePredicateCache.get(cacheKey);
      if (cached === false) return false;
      if (cached === true) continue;

      const matches = Boolean(predicate(baseStyle(colStyleId, rowStyleId)));
      basePredicateCache.set(cacheKey, matches);
      if (!matches) return false;
    }
  }

  return true;
}

export function toggleBold(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean" ? options.next : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.bold));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { bold: next } }, { label: "Bold" });
  }
}

export function toggleItalic(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.italic));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { italic: next } }, { label: "Italic" });
  }
}

export function toggleUnderline(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.font?.underline));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { underline: next } }, { label: "Underline" });
  }
}

export function setFontSize(doc, sheetId, range, sizePt) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { size: sizePt } }, { label: "Font size" });
  }
}

export function setFontColor(doc, sheetId, range, argb) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { font: { color: argb } }, { label: "Font color" });
  }
}

export function setFillColor(doc, sheetId, range, argb) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(
      sheetId,
      r,
      { fill: { pattern: "solid", fgColor: argb } },
      { label: "Fill color" },
    );
  }
}

const DEFAULT_BORDER_ARGB = "FF000000";

export function applyAllBorders(doc, sheetId, range, { style = "thin", color } = {}) {
  const resolvedColor = color ?? `#${DEFAULT_BORDER_ARGB}`;
  const edge = { style, color: resolvedColor };
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(
      sheetId,
      r,
      { border: { left: edge, right: edge, top: edge, bottom: edge } },
      { label: "Borders" },
    );
  }
}

export function setHorizontalAlign(doc, sheetId, range, align) {
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(
      sheetId,
      r,
      { alignment: { horizontal: align } },
      { label: "Horizontal align" },
    );
  }
}

export function toggleWrap(doc, sheetId, range, options = {}) {
  const next =
    typeof options.next === "boolean"
      ? options.next
      : !allCellsMatch(doc, sheetId, range, (s) => Boolean(s.alignment?.wrapText));
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { alignment: { wrapText: next } }, { label: "Wrap" });
  }
}

export const NUMBER_FORMATS = {
  currency: "$#,##0.00",
  percent: "0%",
  date: "m/d/yyyy",
};

export function applyNumberFormatPreset(doc, sheetId, range, preset) {
  const code = NUMBER_FORMATS[preset];
  if (!code) throw new Error(`Unknown number format preset: ${preset}`);
  for (const r of normalizeRanges(range)) {
    doc.setRangeFormat(sheetId, r, { numberFormat: code }, { label: "Number format" });
  }
}
