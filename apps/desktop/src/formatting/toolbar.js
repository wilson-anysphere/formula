import { normalizeRange, parseRangeA1 } from "../document/coords.js";

const SMALL_RANGE_CELL_THRESHOLD = 5000;

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function deepMerge(base, patch) {
  if (!isPlainObject(base) || !isPlainObject(patch)) return patch;
  const out = { ...base };
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (isPlainObject(value) && isPlainObject(out[key])) {
      out[key] = deepMerge(out[key], value);
    } else {
      out[key] = value;
    }
  }
  return out;
}

function mergeStyleLayers(base, patch) {
  const normalizedBase = isPlainObject(base) ? base : {};
  if (patch == null) return normalizedBase;
  if (!isPlainObject(patch)) return patch;
  return deepMerge(normalizedBase, patch);
}

function styleFromEntry(doc, entry) {
  if (entry == null) return {};
  if (typeof entry === "number") return doc.styleTable?.get(entry) ?? {};
  if (typeof entry === "object") return entry;
  return {};
}

function mapGet(mapLike, key) {
  if (!mapLike) return undefined;
  if (mapLike instanceof Map) return mapLike.get(key);
  if (typeof mapLike === "object") return mapLike[String(key)];
  return undefined;
}

function mapEntries(mapLike) {
  if (!mapLike) return [];
  if (mapLike instanceof Map) return Array.from(mapLike.entries());
  if (typeof mapLike === "object") return Object.entries(mapLike);
  return [];
}

function effectiveCellStyle(doc, sheetId, row, col) {
  if (typeof doc.getCellFormat === "function") {
    return doc.getCellFormat(sheetId, { row, col });
  }
  // Back-compat fallback for older DocumentController surfaces.
  const cell = doc.getCell(sheetId, { row, col });
  return doc.styleTable.get(cell.styleId);
}

function normalizeRanges(rangeOrRanges) {
  if (Array.isArray(rangeOrRanges)) return rangeOrRanges;
  return [rangeOrRanges];
}

function normalizeCellRange(range) {
  const parsed = typeof range === "string" ? parseRangeA1(range) : range;
  return normalizeRange(parsed);
}

function allCellsMatchSingleRange(doc, sheetId, range, predicate) {
  const r = normalizeCellRange(range);
  const rows = r.end.row - r.start.row + 1;
  const cols = r.end.col - r.start.col + 1;
  const area = rows * cols;

  // Exact evaluation for small selections.
  if (area <= SMALL_RANGE_CELL_THRESHOLD) {
    for (let row = r.start.row; row <= r.end.row; row++) {
      for (let col = r.start.col; col <= r.end.col; col++) {
        const style = effectiveCellStyle(doc, sheetId, row, col);
        if (!predicate(style)) return false;
      }
    }
    return true;
  }

  // Conservative fast path for huge selections:
  // Avoid enumerating every coordinate; instead inspect format layers + sparse overrides.
  //
  // This must never return a false positive. It is acceptable to return false in
  // edge cases where every cell is individually overridden.
  const sheetModel = doc.model?.sheets?.get(sheetId);
  if (!sheetModel) return false;

  const sheetStyle = styleFromEntry(doc, sheetModel.defaultStyleId ?? sheetModel.sheetStyleId ?? 0);
  const rowStyleIds = sheetModel.rowStyleIds ?? sheetModel.rowStyles ?? sheetModel.rowFormats;
  const colStyleIds = sheetModel.colStyleIds ?? sheetModel.colStyles ?? sheetModel.colFormats;

  /** @type {Map<any, any>} */
  const baseStylesByColKey = new Map();

  // 1) Check column defaults across the range (<= 16,384 columns).
  for (let col = r.start.col; col <= r.end.col; col++) {
    const colEntry = mapGet(colStyleIds, col) ?? 0;
    const cacheKey = typeof colEntry === "number" ? `id:${colEntry}` : JSON.stringify(colEntry);
    if (baseStylesByColKey.has(cacheKey)) continue;
    const colStyle = styleFromEntry(doc, colEntry);
    const base = mergeStyleLayers(sheetStyle, colStyle);
    if (!predicate(base)) return false;
    baseStylesByColKey.set(cacheKey, base);
  }

  // 2) Check sparse row defaults within the range.
  const rowEntries = mapEntries(rowStyleIds);
  const baseStyles = Array.from(baseStylesByColKey.values());
  const MAX_ROW_COL_COMBOS = 5000;
  if (rowEntries.length * baseStyles.length > MAX_ROW_COL_COMBOS) {
    // Too many combinations to check cheaply; fall back to "unknown" (false).
    return false;
  }
  for (const [rawRow, rowEntry] of rowEntries) {
    const row = Number(rawRow);
    if (!Number.isInteger(row) || row < r.start.row || row > r.end.row) continue;
    const rowStyle = styleFromEntry(doc, rowEntry);
    for (const base of baseStyles) {
      const merged = mergeStyleLayers(base, rowStyle);
      if (!predicate(merged)) return false;
    }
  }

  // 3) Check sparse cell overrides in the sheet cell map (only styleId != 0).
  for (const [key, cell] of sheetModel.cells?.entries?.() ?? []) {
    if (!cell || cell.styleId === 0) continue;
    const [rowStr, colStr] = String(key).split(",");
    const row = Number(rowStr);
    const col = Number(colStr);
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;
    if (row < r.start.row || row > r.end.row || col < r.start.col || col > r.end.col) continue;
    const style = effectiveCellStyle(doc, sheetId, row, col);
    if (!predicate(style)) return false;
  }

  return true;
}

function allCellsMatch(doc, sheetId, rangeOrRanges, predicate) {
  for (const range of normalizeRanges(rangeOrRanges)) {
    if (!allCellsMatchSingleRange(doc, sheetId, range, predicate)) return false;
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
