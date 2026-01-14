import { isCellEmpty, normalizeRange, parseA1Range, rangeToA1 } from "./a1.js";
import { throwIfAborted } from "./abort.js";

// Detecting connected regions in a dense matrix requires an O(rows*cols) visited grid.
// Cap the scan to avoid catastrophic allocations when callers accidentally pass
// Excel-scale ranges (1,048,576 x 16,384).
const DEFAULT_DATA_REGION_SCAN_CELL_LIMIT = 200_000;
// Region analysis (header detection + type inference) should also stay bounded for large tables.
// We only sample a prefix of the data rows to infer types / sample values.
const DEFAULT_MAX_ANALYZE_ROWS = 500;
const DEFAULT_MAX_SAMPLE_VALUES_PER_COLUMN = 3;

function isPlainObject(value) {
  return value != null && typeof value === "object" && !Array.isArray(value);
}

function isTypedValue(value) {
  return isPlainObject(value) && typeof /** @type {any} */ (value).t === "string";
}

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
      // Avoid stringifying potentially huge spills; we only need a stable non-empty marker.
      return "[array]";
    default:
      // Defensive: prefer embedded scalar payloads and otherwise fall back to stable JSON.
      if (Object.prototype.hasOwnProperty.call(v, "v")) return v.v ?? null;
      try {
        return JSON.stringify(value);
      } catch {
        return String(value);
      }
  }
}

/**
 * @param {unknown} value
 * @returns {{ imageId: string, altText: string | null } | null}
 */
function parseImageValue(value) {
  if (!isPlainObject(value)) return null;
  const obj = /** @type {any} */ (value);

  let payload = null;
  // DocumentController / formula-model envelope: `{ type: "image", value: {...} }`.
  if (typeof obj.type === "string") {
    if (obj.type.toLowerCase() !== "image") return null;
    payload = isPlainObject(obj.value) ? obj.value : null;
  } else {
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
  let out = typedValueToScalar(value);

  // DocumentController rich text values: `{ text, runs }`.
  if (isPlainObject(out) && typeof /** @type {any} */ (out).text === "string") {
    return /** @type {any} */ (out).text;
  }

  const image = parseImageValue(out);
  if (image) return image.altText ?? "[Image]";

  // Treat `{}` as an empty cell; it's a common sparse representation.
  if (isPlainObject(out) && out.constructor === Object && Object.keys(out).length === 0) return null;

  return out;
}

function isCellEffectivelyEmpty(value) {
  return isCellEmpty(valueToScalar(value));
}

/**
 * @param {unknown} value
 * @returns {string}
 */
function scalarToSampleString(value) {
  if (isCellEmpty(value)) return "";
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return String(value);
  if (value instanceof Date) {
    // Avoid calling per-instance overrides (e.g. `date.toISOString = () => "secret"`).
    let time = NaN;
    try {
      time = Date.prototype.getTime.call(value);
    } catch {
      time = NaN;
    }
    if (Number.isNaN(time)) return "";
    try {
      return Date.prototype.toISOString.call(value);
    } catch {
      // Invalid dates throw in `toISOString()`; fall back to a stable string form.
      try {
        return Date.prototype.toString.call(value);
      } catch {
        return "";
      }
    }
  }
  if (value && typeof value === "object") {
    try {
      const json = JSON.stringify(value);
      if (json === "{}") return "";
      return typeof json === "string" ? json : String(value);
    } catch {
      return "Object";
    }
  }
  return String(value);
}

/**
 * Normalize a numeric limit option.
 *
 * @param {unknown} value
 * @param {number} fallback
 */
function normalizeNonNegativeInt(value, fallback) {
  if (typeof value === "number" && Number.isFinite(value) && value >= 0) return Math.floor(value);
  return fallback;
}

/**
 * @typedef {"empty"|"number"|"boolean"|"date"|"string"|"formula"|"mixed"} InferredType
 */

/**
 * @param {unknown} value
 * @returns {InferredType}
 */
export function inferCellType(value) {
  const scalar = valueToScalar(value);
  if (isCellEmpty(scalar)) return "empty";
  if (typeof scalar === "number" && Number.isFinite(scalar)) return "number";
  if (typeof scalar === "boolean") return "boolean";
  if (scalar instanceof Date && !Number.isNaN(scalar.getTime())) return "date";
  if (typeof scalar === "string") {
    const trimmed = scalar.trim();
    if (trimmed.startsWith("=")) return "formula";

    // Numeric-like strings are common in CSV imports. Treat them as numbers for schema purposes.
    if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) return "number";

    // ISO-like dates are also common.
    if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
      const parsed = new Date(trimmed);
      if (!Number.isNaN(parsed.getTime())) return "date";
    }

    return "string";
  }
  return "string";
}

/**
 * @param {unknown[]} values
 * @param {{ signal?: AbortSignal }} [options]
 * @returns {InferredType}
 */
export function inferColumnType(values, options = {}) {
  const signal = options.signal;
  const types = new Set();
  for (const value of values) {
    throwIfAborted(signal);
    const t = inferCellType(value);
    if (t !== "empty") types.add(t);
  }

  if (types.size === 0) return "empty";
  if (types.size === 1) return /** @type {InferredType} */ (types.values().next().value);

  // "formula" plus "number" is a common computed column.
  if (types.has("formula") && types.size === 2 && (types.has("number") || types.has("date") || types.has("string"))) {
    return "formula";
  }

  return "mixed";
}

/**
 * @param {unknown} value
 */
function isHeaderCandidateValue(value) {
  const scalar = valueToScalar(value);
  if (isCellEmpty(scalar)) return false;
  if (typeof scalar !== "string") return false;
  const trimmed = scalar.trim();
  if (!trimmed) return false;
  if (trimmed.startsWith("=")) return false;
  // Disqualify pure numbers masquerading as strings.
  if (/^[+-]?\d+(?:\.\d+)?$/.test(trimmed)) return false;
  return true;
}

/**
 * @param {unknown[]} rowValues
 * @param {unknown[] | undefined} nextRowValues
 */
export function isLikelyHeaderRow(rowValues, nextRowValues) {
  const normalizedRow = rowValues.map(valueToScalar);
  const normalizedNext = nextRowValues ? nextRowValues.map(valueToScalar) : undefined;

  const nonEmpty = normalizedRow.filter((v) => !isCellEmpty(v));
  if (nonEmpty.length === 0) return false;

  const headerish = nonEmpty.filter(isHeaderCandidateValue);
  if (headerish.length / nonEmpty.length < 0.6) return false;

  const normalized = headerish.map((v) => String(v).trim().toLowerCase());
  const unique = new Set(normalized);
  if (unique.size !== normalized.length) return false;

  if (!normalizedNext) return true;
  const nextNonEmpty = normalizedNext.filter((v) => !isCellEmpty(v));
  if (nextNonEmpty.length === 0) return true;

  // If the next row is "more numeric" than the first row, it's likely data.
  const nextNumeric = nextNonEmpty.filter((v) => inferCellType(v) === "number").length;
  const nextStrings = nextNonEmpty.filter((v) => inferCellType(v) === "string").length;

  if (nextNumeric > 0) return true;
  if (nextStrings / nextNonEmpty.length < 0.6) return true;

  return false;
}

/**
 * @param {unknown[][]} values
 * @param {{ maxCells?: number, signal?: AbortSignal }} [options]
 * @returns {{ startRow: number, startCol: number, endRow: number, endCol: number }[]}
 */
export function detectDataRegions(values, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const maxCellsRaw = options.maxCells;
  const maxCells =
    typeof maxCellsRaw === "number" && Number.isFinite(maxCellsRaw) && maxCellsRaw > 0
      ? Math.floor(maxCellsRaw)
      : DEFAULT_DATA_REGION_SCAN_CELL_LIMIT;

  const rawRowCount = values.length;
  // Keep rows bounded so the outer loop can't become unbounded (even if the input
  // matrix is sparse / mostly empty).
  const rowCount = Math.max(0, Math.min(rawRowCount, maxCells));
  if (rowCount === 0) return [];

  // Only consider columns present in the scanned prefix of rows.
  let colCountRaw = 0;
  for (let r = 0; r < rowCount; r++) {
    colCountRaw = Math.max(colCountRaw, values[r]?.length ?? 0);
  }
  if (colCountRaw === 0) return [];

  // Prefer preserving rows (up to the safe cap) and clamp the column scan so the
  // visited bitmap stays within `maxCells`.
  const colCount = Math.max(0, Math.min(colCountRaw, Math.max(1, Math.floor(maxCells / rowCount))));

  if (rowCount === 0 || colCount === 0) return [];

  /**
   * `Array.shift()` is O(n) due to element re-indexing.
   * Dense regions (large contiguous blocks) would therefore take ~O(n^2) time to flood-fill.
   *
   * Use an index-based queue + typed visited grid to keep flood-fill linear in the number of
   * visited cells.
   *
   * Flattened indexing: `idx = row * colCount + col`.
   *
   * @type {Uint8Array}
   */
  const visited = new Uint8Array(rowCount * colCount);

  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }[]} */
  const regions = [];

  for (let r = 0; r < rowCount; r++) {
    throwIfAborted(signal);
    const row = values[r];
    const rowOffset = r * colCount;
    for (let c = 0; c < colCount; c++) {
      throwIfAborted(signal);
      const startIdx = rowOffset + c;
      if (visited[startIdx]) continue;
      visited[startIdx] = 1;

      if (isCellEffectivelyEmpty(row?.[c])) continue;

      let minRow = r;
      let maxRow = r;
      let minCol = c;
      let maxCol = c;

      // Use a flat number queue to avoid allocating `[row, col]` tuples for every enqueued cell.
      /** @type {number[]} */
      const queue = [r, c];
      let head = 0;

      while (head < queue.length) {
        throwIfAborted(signal);
        const qr = queue[head++];
        const qc = queue[head++];
        if (qr < minRow) minRow = qr;
        if (qr > maxRow) maxRow = qr;
        if (qc < minCol) minCol = qc;
        if (qc > maxCol) maxCol = qc;

        const qRowOffset = qr * colCount;

        // Inline neighbor exploration to avoid per-cell allocations.
        // Up
        if (qr > 0) {
          const idx = qRowOffset - colCount + qc;
          if (!visited[idx]) {
            visited[idx] = 1;
          if (!isCellEffectivelyEmpty(values[qr - 1]?.[qc])) queue.push(qr - 1, qc);
          }
        }
        // Down
        if (qr + 1 < rowCount) {
          const idx = qRowOffset + colCount + qc;
          if (!visited[idx]) {
            visited[idx] = 1;
            if (!isCellEffectivelyEmpty(values[qr + 1]?.[qc])) queue.push(qr + 1, qc);
          }
        }
        // Left
        if (qc > 0) {
          const idx = qRowOffset + qc - 1;
          if (!visited[idx]) {
            visited[idx] = 1;
            if (!isCellEffectivelyEmpty(values[qr]?.[qc - 1])) queue.push(qr, qc - 1);
          }
        }
        // Right
        if (qc + 1 < colCount) {
          const idx = qRowOffset + qc + 1;
          if (!visited[idx]) {
            visited[idx] = 1;
            if (!isCellEffectivelyEmpty(values[qr]?.[qc + 1])) queue.push(qr, qc + 1);
          }
        }
      }

      regions.push({ startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol });
    }
  }

  throwIfAborted(signal);
  regions.sort((a, b) => (a.startRow - b.startRow) || (a.startCol - b.startCol));
  throwIfAborted(signal);
  return regions;
}

/**
 * @param {unknown[][]} sheetValues
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} normalized
 * @param {{ signal?: AbortSignal, maxAnalyzeRows?: number, maxSampleValuesPerColumn?: number }} [options]
 * @returns {{
 *   hasHeader: boolean,
 *   headers: string[],
 *   inferredColumnTypes: InferredType[],
 *   columns: { name: string, type: InferredType, sampleValues: string[] }[],
 *   rowCount: number,
 *   columnCount: number,
 * }}
 */
function analyzeRegion(sheetValues, normalized, options = {}) {
  const signal = options.signal;
  const maxAnalyzeRows = normalizeNonNegativeInt(options.maxAnalyzeRows, DEFAULT_MAX_ANALYZE_ROWS);
  const maxSampleValuesPerColumn = normalizeNonNegativeInt(
    options.maxSampleValuesPerColumn,
    DEFAULT_MAX_SAMPLE_VALUES_PER_COLUMN
  );

  const startRow = normalized.startRow;
  const endRow = normalized.endRow;
  const startCol = normalized.startCol;
  const endCol = normalized.endCol;

  // Avoid allocating a full copied 2D region matrix. Only slice the first two rows for
  // header heuristics; everything else reads directly from `sheetValues`.
  const headerRowSource = sheetValues[startRow];
  const headerRowValues = Array.isArray(headerRowSource) ? headerRowSource.slice(startCol, endCol + 1) : [];
  const nextRowValues =
    startRow + 1 <= endRow && Array.isArray(sheetValues[startRow + 1])
      ? sheetValues[startRow + 1].slice(startCol, endCol + 1)
      : undefined;
  const hasHeader = isLikelyHeaderRow(headerRowValues, nextRowValues);

  const dataStartRow = hasHeader ? 1 : 0;

  // Preserve previous ragged-row behavior: columnCount is the max per-row slice length
  // within the region.
  let columnCount = 0;
  for (let r = startRow; r <= endRow; r++) {
    throwIfAborted(signal);
    const row = sheetValues[r];
    const rowLen = Array.isArray(row) ? row.length : 0;
    if (rowLen <= startCol) continue;
    const sliceLen = Math.max(0, Math.min(rowLen, endCol + 1) - startCol);
    if (sliceLen > columnCount) columnCount = sliceLen;
  }

  const headers = [];
  for (let c = 0; c < columnCount; c++) {
    throwIfAborted(signal);
    const raw = headerRowValues[c];
    const fallback = `Column${c + 1}`;
    const scalar = valueToScalar(raw);
    const headerText = typeof scalar === "string" ? scalar : scalar == null ? "" : String(scalar);
    headers.push(hasHeader && isHeaderCandidateValue(scalar) ? headerText.trim() : fallback);
  }

  /** @type {InferredType[]} */
  const inferredColumnTypes = [];
  /** @type {{ name: string, type: InferredType, sampleValues: string[] }[]} */
  const columns = [];

  const totalRows = Math.max(endRow - startRow + 1, 0);
  const rowCount = Math.max(totalRows - (hasHeader ? 1 : 0), 0);

  // Sample only a bounded number of data rows (after the header) to infer types and collect
  // sample values.
  const analyzeRows = Math.min(rowCount, maxAnalyzeRows);

  const TYPE_NUMBER = 1 << 0;
  const TYPE_BOOLEAN = 1 << 1;
  const TYPE_DATE = 1 << 2;
  const TYPE_STRING = 1 << 3;
  const TYPE_FORMULA = 1 << 4;

  /** @type {Uint8Array} */
  const typeMaskByCol = new Uint8Array(columnCount);
  /** @type {string[][]} */
  const sampleValuesByCol = Array.from({ length: columnCount }, () => []);

  /**
   * @param {InferredType} t
   */
  function typeToMask(t) {
    switch (t) {
      case "number":
        return TYPE_NUMBER;
      case "boolean":
        return TYPE_BOOLEAN;
      case "date":
        return TYPE_DATE;
      case "string":
        return TYPE_STRING;
      case "formula":
        return TYPE_FORMULA;
      default:
        return 0;
    }
  }

  /**
   * @param {number} mask
   * @returns {InferredType}
   */
  function maskToInferredType(mask) {
    if (mask === 0) return "empty";
    if ((mask & (mask - 1)) === 0) {
      switch (mask) {
        case TYPE_NUMBER:
          return "number";
        case TYPE_BOOLEAN:
          return "boolean";
        case TYPE_DATE:
          return "date";
        case TYPE_STRING:
          return "string";
        case TYPE_FORMULA:
          return "formula";
        default:
          return "mixed";
      }
    }

    // Match inferColumnType's special-casing: formula + {number|date|string} => formula.
    if ((mask & TYPE_FORMULA) !== 0) {
      const other = mask & ~TYPE_FORMULA;
      if (other !== 0 && (other & (other - 1)) === 0 && (other & TYPE_BOOLEAN) === 0) {
        return "formula";
      }
    }
    return "mixed";
  }

  const startDataRowIndex = startRow + dataStartRow;
  for (let i = 0; i < analyzeRows; i++) {
    throwIfAborted(signal);
    const row = sheetValues[startDataRowIndex + i];
    if (!Array.isArray(row)) continue;

    // Scan the selected row across all columns.
    for (let c = 0; c < columnCount; c++) {
      // Avoid per-cell abort checks; per-row is enough to remain responsive while keeping
      // overhead low for wide tables.
      const v = row[startCol + c];
      const scalar = valueToScalar(v);
      if (scalar === undefined || isCellEmpty(scalar)) continue;

      typeMaskByCol[c] |= typeToMask(inferCellType(scalar));

      const samples = sampleValuesByCol[c];
      if (samples.length < maxSampleValuesPerColumn) {
        const s = scalarToSampleString(scalar);
        if (!samples.includes(s)) samples.push(s);
      }
    }
  }

  for (let c = 0; c < columnCount; c++) {
    throwIfAborted(signal);
    const type = maskToInferredType(typeMaskByCol[c] ?? 0);
    inferredColumnTypes.push(type);
    columns.push({
      name: headers[c] ?? `Column${c + 1}`,
      type,
      sampleValues: sampleValuesByCol[c] ?? [],
    });
  }

  return {
    hasHeader,
    headers,
    inferredColumnTypes,
    columns,
    rowCount,
    columnCount,
  };
}

/**
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} outer
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} inner
 */
function rangeContains(outer, inner) {
  return (
    outer.startRow <= inner.startRow &&
    outer.startCol <= inner.startCol &&
    outer.endRow >= inner.endRow &&
    outer.endCol >= inner.endCol
  );
}

/**
 * @typedef {{ name: string, range: string, columns: { name: string, type: InferredType, sampleValues: string[] }[], rowCount: number }} TableSchema
 * @typedef {{ name: string, range: string }} NamedRangeSchema
 * @typedef {{ range: string, hasHeader: boolean, headers: string[], inferredColumnTypes: InferredType[], rowCount: number, columnCount: number }} DataRegionSchema
 * @typedef {{ name: string, tables: TableSchema[], namedRanges: NamedRangeSchema[], dataRegions: DataRegionSchema[] }} SheetSchema
 */

/**
 * Extract a schema-first representation of a sheet suitable for LLM context.
 *
 * The input model is intentionally minimal: a 2D array of values plus optional metadata
 * (named ranges, structured tables). This makes the package usable before the full
 * spreadsheet engine is wired in.
 *
 * @param {{
 *   name: string,
 *   values: unknown[][],
 *   /**
 *    * Optional coordinate origin (0-based) for the provided `values` matrix.
 *    *
 *    * When `values` is a cropped window of a larger sheet (e.g. a capped used-range
 *    * sample), `origin` lets schema extraction produce correct absolute A1 ranges.
 *    *\/
 *   origin?: { row: number, col: number },
 *   namedRanges?: NamedRangeSchema[],
 *   tables?: { name: string, range: string }[]
 * }} sheet
 * @param {{
 *   signal?: AbortSignal,
 *   /**
 *    * Maximum number of data rows (excluding the header row) to scan when inferring column types.
 *    * Defaults to 500.
 *    *\/
 *   maxAnalyzeRows?: number,
 *   /**
 *    * Maximum number of unique sample values to capture per column. Defaults to 3.
 *    *\/
 *   maxSampleValuesPerColumn?: number,
 * }} [options]
 * @returns {SheetSchema}
 */
export function extractSheetSchema(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const maxAnalyzeRows = normalizeNonNegativeInt(options.maxAnalyzeRows, DEFAULT_MAX_ANALYZE_ROWS);
  const maxSampleValuesPerColumn = normalizeNonNegativeInt(
    options.maxSampleValuesPerColumn,
    DEFAULT_MAX_SAMPLE_VALUES_PER_COLUMN
  );
  const origin = sheet && typeof sheet === "object" && sheet.origin && typeof sheet.origin === "object"
    ? {
         row: Number.isInteger(sheet.origin.row) && sheet.origin.row >= 0 ? sheet.origin.row : 0,
         col: Number.isInteger(sheet.origin.col) && sheet.origin.col >= 0 ? sheet.origin.col : 0,
      }
    : { row: 0, col: 0 };
  throwIfAborted(signal);
  const matrixRowCount = sheet.values.length;
  let matrixColCount = 0;
  for (const row of sheet.values) {
    throwIfAborted(signal);
    matrixColCount = Math.max(matrixColCount, row?.length ?? 0);
  }

  /**
   * Clamp a rect (0-based) to the bounds of `sheet.values`.
   *
   * Returns null when the rect does not intersect the provided matrix at all.
   *
   * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} rect
   */
  function clampRectToMatrix(rect) {
    if (matrixRowCount === 0 || matrixColCount === 0) return null;
    if (rect.endRow < 0 || rect.endCol < 0) return null;
    if (rect.startRow >= matrixRowCount || rect.startCol >= matrixColCount) return null;

    const startRow = Math.max(0, Math.min(rect.startRow, matrixRowCount - 1));
    const endRow = Math.max(0, Math.min(rect.endRow, matrixRowCount - 1));
    const startCol = Math.max(0, Math.min(rect.startCol, matrixColCount - 1));
    const endCol = Math.max(0, Math.min(rect.endCol, matrixColCount - 1));
    return { startRow, startCol, endRow, endCol };
  }
  const regions = detectDataRegions(sheet.values, { signal });

  /** @type {DataRegionSchema[]} */
  const dataRegions = [];
  /** @type {TableSchema[]} */
  const implicitTables = [];
  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }[]} */
  const implicitTableRects = [];

  for (let i = 0; i < regions.length; i++) {
    throwIfAborted(signal);
    const region = regions[i];
    const normalized = normalizeRange(region);
    const analyzed = analyzeRegion(sheet.values, normalized, { signal, maxAnalyzeRows, maxSampleValuesPerColumn });
    const rect = {
      startRow: normalized.startRow + origin.row,
      endRow: normalized.endRow + origin.row,
      startCol: normalized.startCol + origin.col,
      endCol: normalized.endCol + origin.col,
    };
    const range = rangeToA1({ ...rect, sheetName: sheet.name });

    dataRegions.push({
      range,
      hasHeader: analyzed.hasHeader,
      headers: analyzed.headers,
      inferredColumnTypes: analyzed.inferredColumnTypes,
      rowCount: analyzed.rowCount,
      columnCount: analyzed.columnCount,
    });

    implicitTables.push({
      name: `Region${i + 1}`,
      range,
      columns: analyzed.columns,
      rowCount: analyzed.rowCount,
    });
    // Track the rect in absolute sheet coordinates so explicit table metadata (A1 ranges)
    // can be reconciled even when `values` is a cropped window (origin offset).
    implicitTableRects.push(rect);
  }

  /** @type {{ name: string, range: string, rect: { startRow: number, startCol: number, endRow: number, endCol: number } }[]} */
  const explicitDefs = [];

  if (sheet.tables?.length) {
    const seenRanges = new Set();
    for (const table of sheet.tables) {
      throwIfAborted(signal);
      if (!table || typeof table !== "object") continue;
      if (typeof table.range !== "string" || typeof table.name !== "string") continue;

      let parsed;
      try {
        parsed = parseA1Range(table.range);
      } catch {
        continue;
      }

      if (parsed.sheetName && parsed.sheetName !== sheet.name) continue;
      const rect = normalizeRange(parsed);
      const canonicalRange = rangeToA1({ ...rect, sheetName: sheet.name });

      if (seenRanges.has(canonicalRange)) continue;
      seenRanges.add(canonicalRange);
      explicitDefs.push({ name: table.name, range: canonicalRange, rect });
    }
  }

  /** @type {{ table: TableSchema, rect: { startRow: number, startCol: number, endRow: number, endCol: number } }[]} */
  const tableEntries = [];

  if (explicitDefs.length) {
    const coveredImplicit = new Set();
    for (let i = 0; i < implicitTableRects.length; i++) {
      throwIfAborted(signal);
      const implicitRect = implicitTableRects[i];
      for (const explicit of explicitDefs) {
        throwIfAborted(signal);
        if (rangeContains(explicit.rect, implicitRect)) {
          coveredImplicit.add(i);
          break;
        }
      }
    }

    const implicitUncovered = [];
    for (let i = 0; i < implicitTables.length; i++) {
      throwIfAborted(signal);
      if (coveredImplicit.has(i)) continue;
      implicitUncovered.push({ table: implicitTables[i], rect: implicitTableRects[i] });
    }

    // Re-number implicit regions so we don't end up with confusing gaps when some
    // regions are replaced by explicit tables.
    for (let i = 0; i < implicitUncovered.length; i++) {
      throwIfAborted(signal);
      implicitUncovered[i].table.name = `Region${i + 1}`;
    }

    tableEntries.push(...implicitUncovered);

    for (const explicit of explicitDefs) {
      throwIfAborted(signal);
      const localRect =
        origin.row === 0 && origin.col === 0
          ? explicit.rect
         : {
               startRow: explicit.rect.startRow - origin.row,
               endRow: explicit.rect.endRow - origin.row,
               startCol: explicit.rect.startCol - origin.col,
               endCol: explicit.rect.endCol - origin.col,
             };
      const clamped = clampRectToMatrix(localRect);
      const analyzed = clamped
        ? analyzeRegion(sheet.values, clamped, { signal, maxAnalyzeRows, maxSampleValuesPerColumn })
        : { columns: [], rowCount: 0 };
      tableEntries.push({
        table: {
          name: explicit.name,
          range: explicit.range,
          columns: analyzed.columns,
          rowCount: analyzed.rowCount,
        },
        rect: explicit.rect,
      });
    }

    throwIfAborted(signal);
    tableEntries.sort((a, b) => (a.rect.startRow - b.rect.startRow) || (a.rect.startCol - b.rect.startCol));
  } else {
    tableEntries.push(...implicitTables.map((t, idx) => ({ table: t, rect: implicitTableRects[idx] })));
  }

  /** @type {NamedRangeSchema[]} */
  const normalizedNamedRanges = [];
  if (Array.isArray(sheet.namedRanges)) {
    const seen = new Set();
    for (const nr of sheet.namedRanges) {
      throwIfAborted(signal);
      if (!nr || typeof nr !== "object") continue;
      if (typeof nr.name !== "string" || typeof nr.range !== "string") continue;
      let parsed;
      try {
        parsed = parseA1Range(nr.range);
      } catch {
        continue;
      }
      if (parsed.sheetName && parsed.sheetName !== sheet.name) continue;
      const rect = normalizeRange(parsed);
      const canonicalRange = rangeToA1({ ...rect, sheetName: sheet.name });
      const key = `${nr.name}\u0000${canonicalRange}`;
      if (seen.has(key)) continue;
      seen.add(key);
      normalizedNamedRanges.push({ name: nr.name, range: canonicalRange });
    }
  }

  throwIfAborted(signal);
  return {
    name: sheet.name,
    tables: tableEntries.map((t) => t.table),
    namedRanges: normalizedNamedRanges,
    dataRegions,
  };
}
