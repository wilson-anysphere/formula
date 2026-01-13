import { isCellEmpty, rangeToA1 } from "./a1.js";
import { throwIfAborted } from "./abort.js";
import { inferColumnType, isLikelyHeaderRow } from "./schema.js";

/**
 * @typedef {import("./schema.js").InferredType} InferredType
 */

/**
 * @param {any} rect
 * @returns {{ r0: number, c0: number, r1: number, c1: number } | null}
 */
function normalizeRect(rect) {
  if (!rect || typeof rect !== "object") return null;
  const r0 = rect.r0;
  const c0 = rect.c0;
  const r1 = rect.r1;
  const c1 = rect.c1;
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return null;
  return {
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
}

/**
 * @param {any} sheet
 * @returns {any[][] | null}
 */
function getSheetMatrix(sheet) {
  if (Array.isArray(sheet?.cells)) return sheet.cells;
  if (Array.isArray(sheet?.values)) return sheet.values;
  return null;
}

/**
 * @param {any} sheet
 * @returns {Map<string, any> | null}
 */
function getSheetCellMap(sheet) {
  const cells = sheet?.cells;
  // Accept either a real Map or any Map-like object that implements `.get()`.
  // Some hosts wrap the underlying Map (proxy, custom class) while maintaining the
  // same access pattern.
  if (cells && typeof cells.get === "function") return cells;
  return null;
}

/**
 * Sheets can optionally provide an origin offset when their `cells`/`values` matrices
 * represent a window into a larger absolute sheet.
 *
 * This mirrors the origin semantics used elsewhere in `ai-context` (e.g. RAG chunking).
 *
 * @param {any} sheet
 */
function normalizeSheetOrigin(sheet) {
  if (!sheet || typeof sheet !== "object" || !sheet.origin || typeof sheet.origin !== "object") {
    return { row: 0, col: 0 };
  }
  const row = Number.isInteger(sheet.origin.row) && sheet.origin.row >= 0 ? sheet.origin.row : 0;
  const col = Number.isInteger(sheet.origin.col) && sheet.origin.col >= 0 ? sheet.origin.col : 0;
  return { row, col };
}

/**
 * @param {any} sheet
 * @param {number} row
 * @param {number} col
 */
function getCellRaw(sheet, row, col) {
  const origin = normalizeSheetOrigin(sheet);
  const matrix = getSheetMatrix(sheet);
  if (matrix) {
    const localRow = row - origin.row;
    const localCol = col - origin.col;
    if (localRow < 0 || localCol < 0) return null;
    return matrix[localRow]?.[localCol];
  }
  const map = getSheetCellMap(sheet);
  if (map) {
    const localRow = row - origin.row;
    const localCol = col - origin.col;
    // Prefer origin-adjusted coordinates for sparse maps that store local indices, but
    // also fall back to absolute indices for callers that store absolute keys.
    return (
      (localRow >= 0 && localCol >= 0
        ? map.get(`${localRow},${localCol}`) ?? map.get(`${localRow}:${localCol}`)
        : null) ??
      map.get(`${row},${col}`) ??
      map.get(`${row}:${col}`) ??
      null
    );
  }
  if (typeof sheet?.getCell === "function") return sheet.getCell(row, col);
  return null;
}

/**
 * Normalize supported cell shapes to a scalar or formula string suitable for type inference.
 *
 * This intentionally mirrors `packages/ai-rag`'s loose cell support, but is kept local so
 * `ai-context` can operate without depending on the full RAG workbook pipeline.
 *
 * @param {any} raw
 * @returns {unknown}
 */
function cellToScalar(raw) {
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    // Treat `{}` as an empty cell; it's a common sparse representation (notably from
    // `packages/ai-rag` normalization and some SpreadsheetApi adapters).
    if (raw.constructor === Object && Object.keys(raw).length === 0) return null;

    // ai-rag style: { v, f }
    if (Object.prototype.hasOwnProperty.call(raw, "v") || Object.prototype.hasOwnProperty.call(raw, "f")) {
      const formula = raw.f;
      if (typeof formula === "string") {
        const trimmed = formula.trim();
        if (trimmed) return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
      }
      return raw.v ?? null;
    }
    // Alternate shape: { value, formula }
    if (Object.prototype.hasOwnProperty.call(raw, "value") || Object.prototype.hasOwnProperty.call(raw, "formula")) {
      const formula = raw.formula;
      if (typeof formula === "string") {
        const trimmed = formula.trim();
        if (trimmed) return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
      }
      return raw.value ?? null;
    }
    if (raw instanceof Date) return raw;
  }

  if (typeof raw === "string") {
    const trimmed = raw.trim();
    return trimmed ? trimmed : "";
  }
  return raw;
}

/**
 * @param {any} sheet
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @param {number} row
 * @param {number} colCount
 * @param {AbortSignal | undefined} signal
 * @returns {unknown[]}
 */
function readRowScalars(sheet, rect, row, colCount, signal) {
  throwIfAborted(signal);
  const out = [];
  const count = Math.max(0, colCount);
  for (let offset = 0; offset < count; offset++) {
    throwIfAborted(signal);
    out.push(cellToScalar(getCellRaw(sheet, row, rect.c0 + offset)));
  }
  return out;
}

/**
 * @param {unknown} value
 */
function isHeaderCandidateValue(value) {
  if (isCellEmpty(value)) return false;
  if (typeof value !== "string") return false;
  const trimmed = value.trim();
  if (!trimmed) return false;
  if (trimmed.startsWith("=")) return false;
  if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) return false;
  return true;
}

/**
 * @param {string} sheetName
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 */
function rectToRangeA1(sheetName, rect) {
  return rangeToA1({
    sheetName,
    startRow: rect.r0,
    startCol: rect.c0,
    endRow: rect.r1,
    endCol: rect.c1,
  });
}

/**
 * @param {any} sheet
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @param {{ maxAnalyzeRows: number, maxAnalyzeCols: number, signal?: AbortSignal }} options
 * @returns {{ hasHeader: boolean, headers: string[], inferredColumnTypes: InferredType[], rowCount: number, columnCount: number }}
 */
function analyzeTableRect(sheet, rect, options) {
  const signal = options.signal;
  throwIfAborted(signal);
  const fullColumnCount = Math.max(0, rect.c1 - rect.c0 + 1);
  const maxAnalyzeCols = Math.max(0, options.maxAnalyzeCols);
  const analyzedColumnCount = maxAnalyzeCols > 0 ? Math.min(fullColumnCount, maxAnalyzeCols) : fullColumnCount;
  const totalRows = Math.max(0, rect.r1 - rect.r0 + 1);

  const headerRowValues = analyzedColumnCount > 0 ? readRowScalars(sheet, rect, rect.r0, analyzedColumnCount, signal) : [];
  const nextRowValues =
    totalRows > 1 && analyzedColumnCount > 0
      ? readRowScalars(sheet, rect, rect.r0 + 1, analyzedColumnCount, signal)
      : undefined;

  const hasHeader = isLikelyHeaderRow(headerRowValues, nextRowValues);
  const headers = [];
  for (let c = 0; c < analyzedColumnCount; c++) {
    throwIfAborted(signal);
    const raw = headerRowValues[c];
    const fallback = `Column${c + 1}`;
    headers.push(hasHeader && isHeaderCandidateValue(raw) ? String(raw).trim() : fallback);
  }

  const dataStartRow = rect.r0 + (hasHeader ? 1 : 0);
  const rowCount = Math.max(totalRows - (hasHeader ? 1 : 0), 0);

  // Sample data rows (bounded).
  const maxAnalyzeRows = Math.max(0, options.maxAnalyzeRows);
  const sampleEndRow = rowCount === 0 ? dataStartRow - 1 : Math.min(rect.r1, dataStartRow + maxAnalyzeRows - 1);

  /** @type {unknown[][]} */
  const valuesByCol = Array.from({ length: analyzedColumnCount }, () => []);
  for (let r = dataStartRow; r <= sampleEndRow; r++) {
    throwIfAborted(signal);
    for (let c = 0; c < analyzedColumnCount; c++) {
      throwIfAborted(signal);
      valuesByCol[c].push(cellToScalar(getCellRaw(sheet, r, rect.c0 + c)));
    }
  }

  /** @type {InferredType[]} */
  const inferredColumnTypes = [];
  for (let c = 0; c < analyzedColumnCount; c++) {
    throwIfAborted(signal);
    inferredColumnTypes.push(inferColumnType(valuesByCol[c], { signal }));
  }

  return { hasHeader, headers, inferredColumnTypes, rowCount, columnCount: fullColumnCount };
}

/**
 * Extract a compact workbook-level schema summary suitable for LLM context.
 *
 * This intentionally focuses on structured workbook metadata:
 * - Sheets list
 * - Tables (rect + inferred headers/types)
 * - Named ranges
 *
 * Unlike the RAG pipeline, this does not enumerate all non-empty cells.
 *
 * @param {{
 *   id: string,
 *   sheets: Array<{ name: string, cells?: any, values?: any, getCell?: (row: number, col: number) => any }>,
 *   tables?: Array<{ name: string, sheetName: string, rect: any }>,
 *   namedRanges?: Array<{ name: string, sheetName: string, rect: any }>
 * }} workbook
 * @param {{ maxAnalyzeRows?: number, maxAnalyzeCols?: number, signal?: AbortSignal }} [options]
 * @returns {{
 *   id: string,
 *   sheets: Array<{ name: string }>,
 *   tables: Array<{
 *     name: string,
 *     sheetName: string,
 *     rect: { r0: number, c0: number, r1: number, c1: number },
 *     rangeA1: string,
 *     headers: string[],
 *     inferredColumnTypes: InferredType[],
 *     rowCount: number,
 *     columnCount: number
 *   }>,
 *   namedRanges: Array<{ name: string, sheetName: string, rect: { r0: number, c0: number, r1: number, c1: number }, rangeA1: string }>
 * }}
 */
export function extractWorkbookSchema(workbook, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const workbookId = typeof workbook?.id === "string" ? workbook.id : "";
  const maxAnalyzeRowsRaw = options.maxAnalyzeRows;
  const maxAnalyzeRows =
    typeof maxAnalyzeRowsRaw === "number" && Number.isFinite(maxAnalyzeRowsRaw) && maxAnalyzeRowsRaw > 0
      ? Math.floor(maxAnalyzeRowsRaw)
      : 50;
  const maxAnalyzeColsRaw = options.maxAnalyzeCols;
  const maxAnalyzeCols =
    typeof maxAnalyzeColsRaw === "number" && Number.isFinite(maxAnalyzeColsRaw) && maxAnalyzeColsRaw > 0
      ? Math.floor(maxAnalyzeColsRaw)
      : 50;

  const workbookSheets = Array.isArray(workbook?.sheets) ? workbook.sheets : [];
  const workbookTables = Array.isArray(workbook?.tables) ? workbook.tables : [];
  const workbookNamedRanges = Array.isArray(workbook?.namedRanges) ? workbook.namedRanges : [];

  /** @type {Array<{ name: string, sheet: any }>} */
  const sheetEntries = [];
  for (const sheet of workbookSheets) {
    throwIfAborted(signal);
    const name = typeof sheet?.name === "string" ? sheet.name : "";
    if (!name) continue;
    sheetEntries.push({ name, sheet });
  }

  throwIfAborted(signal);
  sheetEntries.sort((a, b) => a.name.localeCompare(b.name));
  const sheetByName = new Map(sheetEntries.map((s) => [s.name, s.sheet]));

  /** @type {ReturnType<typeof extractWorkbookSchema>["tables"]} */
  const tables = [];
  for (const table of workbookTables) {
    throwIfAborted(signal);
    const name = typeof table?.name === "string" ? table.name.trim() : "";
    const sheetName = typeof table?.sheetName === "string" ? table.sheetName : "";
    const rect = normalizeRect(table?.rect);
    if (!name || !sheetName || !rect) continue;
    const sheet = sheetByName.get(sheetName) ?? null;
    const analysis = sheet
      ? analyzeTableRect(sheet, rect, { maxAnalyzeRows, maxAnalyzeCols, signal })
      : {
          headers: [],
          inferredColumnTypes: /** @type {InferredType[]} */ ([]),
          // Without sheet cell data we cannot run header heuristics; treat every row as data.
          rowCount: Math.max(0, rect.r1 - rect.r0 + 1),
          columnCount: Math.max(0, rect.c1 - rect.c0 + 1),
        };
    tables.push({
      name,
      sheetName,
      rect,
      rangeA1: rectToRangeA1(sheetName, rect),
      headers: analysis.headers,
      inferredColumnTypes: analysis.inferredColumnTypes,
      rowCount: analysis.rowCount,
      columnCount: analysis.columnCount,
    });
  }

  throwIfAborted(signal);
  tables.sort(
    (a, b) =>
      a.sheetName.localeCompare(b.sheetName) ||
      a.rect.r0 - b.rect.r0 ||
      a.rect.c0 - b.rect.c0 ||
      a.name.localeCompare(b.name),
  );

  /** @type {ReturnType<typeof extractWorkbookSchema>["namedRanges"]} */
  const namedRanges = [];
  for (const nr of workbookNamedRanges) {
    throwIfAborted(signal);
    const name = typeof nr?.name === "string" ? nr.name.trim() : "";
    const sheetName = typeof nr?.sheetName === "string" ? nr.sheetName : "";
    const rect = normalizeRect(nr?.rect);
    if (!name || !sheetName || !rect) continue;
    namedRanges.push({ name, sheetName, rect, rangeA1: rectToRangeA1(sheetName, rect) });
  }

  throwIfAborted(signal);
  namedRanges.sort(
    (a, b) =>
      a.sheetName.localeCompare(b.sheetName) ||
      a.rect.r0 - b.rect.r0 ||
      a.rect.c0 - b.rect.c0 ||
      a.name.localeCompare(b.name),
  );

  return {
    id: workbookId,
    sheets: sheetEntries.map((s) => ({ name: s.name })),
    tables,
    namedRanges,
  };
}
