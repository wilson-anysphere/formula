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

  /** @type {any} */
  let r0 = rect.r0;
  /** @type {any} */
  let c0 = rect.c0;
  /** @type {any} */
  let r1 = rect.r1;
  /** @type {any} */
  let c1 = rect.c1;

  // Some callers use a Range-like shape (startRow/startCol/endRow/endCol).
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) {
    r0 = rect.startRow;
    c0 = rect.startCol;
    r1 = rect.endRow;
    c1 = rect.endCol;
  }

  // Some callers use nested { start: {row,col}, end: {row,col} }.
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) {
    r0 = rect.start?.row;
    c0 = rect.start?.col;
    r1 = rect.end?.row;
    c1 = rect.end?.col;
  }

  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return null;
  return {
    r0: Math.min(r0, r1),
    c0: Math.min(c0, c1),
    r1: Math.max(r0, r1),
    c1: Math.max(c0, c1),
  };
}

/**
 * Normalize a workbook collection field (sheets/tables/namedRanges) into an array of
 * `{ key, value }` entries.
 *
 * Some hosts represent metadata as Maps (e.g. keyed by name). We keep the public
 * API typed as arrays, but accept Map/Set/object shapes at runtime for robustness.
 *
 * @param {any} value
 * @returns {Array<{ key: string, value: any }>}
 */
function normalizeCollectionEntries(value) {
  if (Array.isArray(value)) return value.map((v) => ({ key: "", value: v }));
  if (value instanceof Map) {
    return Array.from(value.entries()).map(([k, v]) => ({
      // Avoid calling `String(...)` on arbitrary objects: Map keys can be user-controlled in
      // third-party hosts, and custom `toString()` implementations can throw or leak sensitive
      // strings. Keep primitive keys, drop everything else.
      key:
        typeof k === "string"
          ? k
          : typeof k === "number" || typeof k === "boolean" || typeof k === "bigint"
            ? String(k)
            : "",
      value: v,
    }));
  }
  if (value instanceof Set) return Array.from(value.values()).map((v) => ({ key: "", value: v }));
  if (value && typeof value === "object") {
    // Plain object map: `{ key: value }`
    return Object.entries(value).map(([k, v]) => ({ key: k, value: v }));
  }
  return [];
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

const COORD_KEY_RE = /^\d+(?:,|:)\d+$/;

/**
 * @param {unknown} value
 * @returns {value is Record<string, any>}
 */
function looksLikeSparseCoordKeyedObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return false;
  if (value instanceof Date) return false;
  if (typeof value.get === "function") return false;
  // Some hosts may use `Object.create(null)` for maps; `for..in` still works.
  let seen = 0;
  for (const key in /** @type {any} */ (value)) {
    if (!Object.prototype.hasOwnProperty.call(value, key)) continue;
    seen += 1;
    if (COORD_KEY_RE.test(key)) return true;
    // Keep this check cheap: bail after a small number of keys.
    if (seen >= 20) break;
  }
  return false;
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
  if (looksLikeSparseCoordKeyedObject(cells)) {
    const obj = cells;
    return {
      get(key) {
        return obj[key];
      },
    };
  }
  return null;
}

/**
 * Formula engine typed-value encoding (see `tools/excel-oracle/value-encoding.md`).
 *
 * We support a minimal subset here to improve schema inference:
 * - {"t":"blank"} => empty
 * - {"t":"n","v":number} => number
 * - {"t":"s","v":string} => string
 * - {"t":"b","v":boolean} => boolean
 * - {"t":"e","v":"#DIV/0!"} => string (error text)
 * - {"t":"arr", ...} => "[array]" sentinel string
 *
 * @param {unknown} value
 */
function isTypedValue(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value) && typeof value.t === "string";
}

function isPlainObject(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function parseImageValue(value) {
  if (!isPlainObject(value)) return null;
  const obj = /** @type {any} */ (value);

  let payload = null;
  // DocumentController-style "in-cell image" envelope: `{ type: "image", value: {...} }`.
  if (typeof obj.type === "string") {
    if (obj.type.toLowerCase() !== "image") return null;
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else {
    // Direct payload / legacy shapes.
    payload = obj;
  }

  if (!payload) return null;

  const imageIdRaw = payload.imageId ?? payload.image_id ?? payload.id;
  if (typeof imageIdRaw !== "string") return null;
  const imageId = imageIdRaw.trim();
  if (imageId === "") return null;

  const altTextRaw = payload.altText ?? payload.alt_text ?? payload.alt;
  const altText = typeof altTextRaw === "string" ? altTextRaw.trim() : "";
  return { imageId, altText: altText === "" ? null : altText };
}

function valueToScalar(value) {
  let out = value;
  if (isTypedValue(out)) out = typedValueToScalar(out);

  // DocumentController rich text values: `{ text, runs }`.
  if (isPlainObject(out) && typeof /** @type {any} */ (out).text === "string") {
    return /** @type {any} */ (out).text;
  }

  const image = parseImageValue(out);
  if (image) return image.altText ?? "[Image]";

  return out;
}

/**
 * @param {unknown} value
 * @returns {unknown}
 */
function typedValueToScalar(value) {
  if (!isTypedValue(value)) return value;
  const v = /** @type {any} */ (value);
  switch (v.t) {
    case "blank":
      return null;
    case "n":
      return typeof v.v === "number" ? v.v : v.v == null ? null : Number(v.v);
    case "s":
      return v.v == null ? "" : String(v.v);
    case "b":
      return Boolean(v.v);
    case "e":
      return v.v == null ? "" : String(v.v);
    case "arr":
      // Avoid JSON-stringifying potentially large spilled arrays; we only need
      // something non-empty and stable for type inference.
      return "[array]";
    default:
      // Defensive: unknown typed values should never leak `[object Object]` into schema inference.
      // Prefer the embedded `v` payload when present, and otherwise fall back to a stable JSON string.
      if (Object.prototype.hasOwnProperty.call(v, "v")) return v.v ?? null;
      try {
        return JSON.stringify(value);
      } catch {
        return String(value);
      }
  }
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
    // Normalize DocumentController-style rich values (rich text, in-cell images) into scalars
    // so header/type inference can treat them like normal strings.
    const direct = valueToScalar(raw);
    if (direct !== raw) return direct;

    // Formula-engine typed values can appear directly in cell matrices; unwrap them
    // so header/type inference behaves as expected.
    if (isTypedValue(raw)) return valueToScalar(raw);

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
      return valueToScalar(raw.v ?? null);
    }
    // Alternate shape: { value, formula }
    if (Object.prototype.hasOwnProperty.call(raw, "value") || Object.prototype.hasOwnProperty.call(raw, "formula")) {
      const formula = raw.formula;
      if (typeof formula === "string") {
        const trimmed = formula.trim();
        if (trimmed) return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
      }
      return valueToScalar(raw.value ?? null);
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

  const workbookSheets = normalizeCollectionEntries(workbook?.sheets);
  const workbookTables = normalizeCollectionEntries(workbook?.tables);
  const workbookNamedRanges = normalizeCollectionEntries(workbook?.namedRanges);

  /** @type {Array<{ name: string, sheet: any }>} */
  const sheetEntries = [];
  for (const entry of workbookSheets) {
    throwIfAborted(signal);
    const rawSheet = entry.value;
    const name =
      typeof rawSheet?.name === "string"
        ? rawSheet.name
        : typeof entry.key === "string" && entry.key
          ? entry.key
          : typeof rawSheet === "string"
            ? rawSheet
            : "";
    if (!name) continue;

    const looksLikeSheetObject =
      rawSheet &&
      typeof rawSheet === "object" &&
      !Array.isArray(rawSheet) &&
      ("cells" in rawSheet || "values" in rawSheet || typeof rawSheet.getCell === "function" || "origin" in rawSheet);

    let sheet = rawSheet;
    if (looksLikeSheetObject) {
      if (typeof rawSheet?.name !== "string") sheet = { ...rawSheet, name };
    } else if (Array.isArray(rawSheet)) {
      // Allow sheet maps like `{ Sheet1: [[...]] }` by treating the value as a matrix.
      sheet = { name, values: rawSheet };
    } else if (rawSheet && typeof rawSheet.get === "function") {
      // Allow sheet maps like `{ Sheet1: new Map(...) }` by treating the value as a sparse cell map.
      sheet = { name, cells: rawSheet };
    } else if (looksLikeSparseCoordKeyedObject(rawSheet)) {
      // Allow sheet maps like `{ Sheet1: { "0,0": {...} } }` by treating the value as a sparse cell object map.
      sheet = { name, cells: rawSheet };
    } else {
      sheet = { name };
    }

    sheetEntries.push({ name, sheet });
  }

  throwIfAborted(signal);
  sheetEntries.sort((a, b) => a.name.localeCompare(b.name));
  const sheetByName = new Map(sheetEntries.map((s) => [s.name, s.sheet]));

  /** @type {ReturnType<typeof extractWorkbookSchema>["tables"]} */
  const tables = [];
  for (const entry of workbookTables) {
    throwIfAborted(signal);
    const table = entry.value;
    const name =
      typeof table?.name === "string"
        ? table.name.trim()
        : typeof entry.key === "string"
          ? entry.key.trim()
          : "";
    const sheetName = typeof table?.sheetName === "string" ? table.sheetName : "";
    const rect = normalizeRect(table?.rect ?? table);
    if (!name || !sheetName || !rect) continue;
    const sheet = sheetByName.get(sheetName) ?? null;
    const hasCellData = Boolean(sheet && (getSheetMatrix(sheet) || getSheetCellMap(sheet) || typeof sheet.getCell === "function"));
    const analysis = hasCellData
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
      a.name.localeCompare(b.name) ||
      a.rect.r1 - b.rect.r1 ||
      a.rect.c1 - b.rect.c1,
  );

  /** @type {ReturnType<typeof extractWorkbookSchema>["namedRanges"]} */
  const namedRanges = [];
  for (const entry of workbookNamedRanges) {
    throwIfAborted(signal);
    const nr = entry.value;
    const name =
      typeof nr?.name === "string"
        ? nr.name.trim()
        : typeof entry.key === "string"
          ? entry.key.trim()
          : "";
    const sheetName = typeof nr?.sheetName === "string" ? nr.sheetName : "";
    const rect = normalizeRect(nr?.rect ?? nr);
    if (!name || !sheetName || !rect) continue;
    namedRanges.push({ name, sheetName, rect, rangeA1: rectToRangeA1(sheetName, rect) });
  }

  throwIfAborted(signal);
  namedRanges.sort(
    (a, b) =>
      a.sheetName.localeCompare(b.sheetName) ||
      a.rect.r0 - b.rect.r0 ||
      a.rect.c0 - b.rect.c0 ||
      a.name.localeCompare(b.name) ||
      a.rect.r1 - b.rect.r1 ||
      a.rect.c1 - b.rect.c1,
  );

  return {
    id: workbookId,
    sheets: sheetEntries.map((s) => ({ name: s.name })),
    tables,
    namedRanges,
  };
}
