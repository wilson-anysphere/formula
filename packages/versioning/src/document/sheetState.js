import { cellKey } from "../diff/semanticDiff.js";

const decoder = new TextDecoder();

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

/**
 * Deep-merge two format/style objects.
 *
 * Later layers override earlier ones, but nested objects are merged recursively.
 *
 * @param {any} base
 * @param {any} patch
 * @returns {any}
 */
function deepMerge(base, patch) {
  if (patch == null) return base;
  if (!isPlainObject(base) || !isPlainObject(patch)) return patch;
  /** @type {Record<string, any>} */
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

/**
 * Remove empty nested objects from a style tree.
 *
 * This keeps format diffs stable by avoiding `{ font: {} }`-style artifacts.
 *
 * @param {any} value
 * @returns {any}
 */
function pruneEmptyObjects(value) {
  if (!isPlainObject(value)) return value;
  /** @type {Record<string, any>} */
  const out = {};
  for (const [key, raw] of Object.entries(value)) {
    if (raw === undefined) continue;
    const pruned = pruneEmptyObjects(raw);
    if (isPlainObject(pruned) && Object.keys(pruned).length === 0) continue;
    out[key] = pruned;
  }
  return out;
}

/**
 * @param {any} format
 * @returns {any | null}
 */
function normalizeFormat(format) {
  if (format == null) return null;
  const pruned = pruneEmptyObjects(format);
  if (isPlainObject(pruned) && Object.keys(pruned).length === 0) return null;
  return pruned;
}

/**
 * Snapshot formats are stored inconsistently across schema versions:
 * - as a bare style object
 * - wrapped in `{ format }` or `{ style }` containers
 *
 * @param {any} value
 * @returns {Record<string, any> | null}
 */
function extractStyleObject(value) {
  if (value == null) return null;
  if (isPlainObject(value) && ("format" in value || "style" in value)) {
    const nested = value.format ?? value.style ?? null;
    return isPlainObject(nested) ? nested : null;
  }
  return isPlainObject(value) ? value : null;
}

/**
 * @param {any} raw
 * @param {{ indexKeys: string[] }} opts
 * @returns {Map<number, Record<string, any>>}
 */
function parseIndexedFormats(raw, opts) {
  /** @type {Map<number, Record<string, any>>} */
  const out = new Map();
  if (!raw) return out;

  if (Array.isArray(raw)) {
    for (const entry of raw) {
      let index = null;
      let formatValue = null;

      if (Array.isArray(entry)) {
        index = Number(entry[0]);
        formatValue = entry[1];
      } else if (entry && typeof entry === "object") {
        for (const key of opts.indexKeys) {
          if (key in entry) {
            index = Number(entry[key]);
            break;
          }
        }
        // Avoid treating arbitrary metadata blobs (e.g. row heights) as styles.
        formatValue = entry?.format ?? entry?.style ?? entry?.value ?? null;
      }

      if (!Number.isInteger(index) || index < 0) continue;
      const format = extractStyleObject(formatValue);
      if (!format) continue;
      out.set(index, format);
    }
    return out;
  }

  if (typeof raw === "object") {
    for (const [key, value] of Object.entries(raw)) {
      const index = Number(key);
      if (!Number.isInteger(index) || index < 0) continue;
      const format = extractStyleObject(value?.format ?? value?.style ?? value);
      if (!format) continue;
      out.set(index, format);
    }
  }

  return out;
}

/**
 * Parse sparse rectangular format runs.
 *
 * This is an optional schema extension (Task 118) used to represent formatting applied
 * to arbitrary ranges without materializing per-cell styles.
 *
 * @param {any} raw
 * @returns {Array<{ startRow: number, startCol: number, endRow: number, endCol: number, format: Record<string, any> }>}
 */
function normalizeRangeRuns(raw) {
  /** @type {Array<{ startRow: number, startCol: number, endRow: number, endCol: number, format: Record<string, any> }>} */
  const out = [];
  if (!Array.isArray(raw)) return out;
  for (const run of raw) {
    if (!run || typeof run !== "object") continue;
    let startRow = Number(run.startRow ?? run.start?.row ?? run.sr);
    let startCol = Number(run.startCol ?? run.start?.col ?? run.sc);
    let endRow = Number(run.endRow ?? run.end?.row ?? run.er);
    let endCol = Number(run.endCol ?? run.end?.col ?? run.ec);
    if (!Number.isInteger(startRow) || startRow < 0) continue;
    if (!Number.isInteger(startCol) || startCol < 0) continue;
    if (!Number.isInteger(endRow) || endRow < 0) continue;
    if (!Number.isInteger(endCol) || endCol < 0) continue;
    if (endRow < startRow) [startRow, endRow] = [endRow, startRow];
    if (endCol < startCol) [startCol, endCol] = [endCol, startCol];
    const format = extractStyleObject(run.format ?? run.style ?? run.value);
    if (!format) continue;
    out.push({ startRow, startCol, endRow, endCol, format });
  }
  return out;
}

/**
 * Convert a snapshot produced by `apps/desktop/src/document/DocumentController.encodeState()`
 * into the `SheetState` shape expected by `semanticDiff`.
 *
 * @param {Uint8Array} snapshot
 * @param {{ sheetId: string }} opts
 * @returns {{ cells: Map<string, { value?: any, formula?: string | null, format?: any }> }}
 */
export function sheetStateFromDocumentSnapshot(snapshot, opts) {
  const sheetId = opts?.sheetId;
  if (!sheetId) throw new Error("sheetId is required");

  let parsed;
  try {
    parsed = JSON.parse(decoder.decode(snapshot));
  } catch {
    throw new Error("Invalid document snapshot: not valid JSON");
  }

  const sheets = Array.isArray(parsed?.sheets) ? parsed.sheets : [];
  /** @type {Map<string, any>} */
  const cells = new Map();

  const sheet = sheets.find((s) => s?.id === sheetId);
  if (!sheet) return { cells };

  // --- Formatting defaults (layered formats) ---
  // Snapshot schema v1 stores formatting directly on each cell entry. Newer schema versions
  // can store formats as layered defaults (sheet/row/col) and keep per-cell overrides in
  // `entry.format`. For semantic diffs we want the *effective* cell format.
  const sheetDefaultFormat =
    extractStyleObject(sheet?.defaultFormat) ??
    extractStyleObject(sheet?.defaultStyle) ??
    extractStyleObject(sheet?.defaultCellFormat) ??
    extractStyleObject(sheet?.defaultCellStyle) ??
    extractStyleObject(sheet?.sheetFormat) ??
    extractStyleObject(sheet?.sheetStyle) ??
    extractStyleObject(sheet?.sheetDefaultFormat) ??
    extractStyleObject(sheet?.cellFormat) ??
    extractStyleObject(sheet?.cellStyle) ??
    extractStyleObject(sheet?.format) ??
    extractStyleObject(sheet?.style) ??
    extractStyleObject(sheet?.defaults?.format) ??
    extractStyleObject(sheet?.defaults?.style) ??
    extractStyleObject(sheet?.cellDefaults?.format) ??
    extractStyleObject(sheet?.cellDefaults?.style) ??
    extractStyleObject(sheet?.formatDefaults?.sheet) ??
    extractStyleObject(sheet?.formatDefaults?.default) ??
    null;

  const rowFormatSources = [
    sheet?.rowDefaults,
    sheet?.rowFormats,
    sheet?.rowsFormats,
    sheet?.rowStyles,
    sheet?.rowFormat,
    sheet?.rowStyle,
    sheet?.defaults?.rows,
    sheet?.defaults?.rowDefaults,
    sheet?.defaults?.rowFormats,
    sheet?.defaults?.rowStyles,
    sheet?.formatDefaults?.rows,
    sheet?.formatDefaults?.rowDefaults,
    sheet?.formatDefaults?.rowFormats,
    sheet?.rows,
  ];
  /** @type {Map<number, Record<string, any>>} */
  let rowFormats = new Map();
  for (const source of rowFormatSources) {
    rowFormats = parseIndexedFormats(source, { indexKeys: ["row", "r", "index"] });
    if (rowFormats.size > 0) break;
  }

  const colFormatSources = [
    sheet?.colDefaults,
    sheet?.colFormats,
    sheet?.colStyles,
    sheet?.colFormat,
    sheet?.colStyle,
    sheet?.columnDefaults,
    sheet?.columnFormats,
    sheet?.columnStyles,
    sheet?.defaults?.cols,
    sheet?.defaults?.columns,
    sheet?.defaults?.colDefaults,
    sheet?.defaults?.colFormats,
    sheet?.defaults?.colStyles,
    sheet?.formatDefaults?.cols,
    sheet?.formatDefaults?.columns,
    sheet?.formatDefaults?.colDefaults,
    sheet?.formatDefaults?.colFormats,
    sheet?.cols,
    sheet?.columns,
  ];
  /** @type {Map<number, Record<string, any>>} */
  let colFormats = new Map();
  for (const source of colFormatSources) {
    colFormats = parseIndexedFormats(source, { indexKeys: ["col", "c", "index"] });
    if (colFormats.size > 0) break;
  }

  const formatRuns = normalizeRangeRuns(
    sheet?.formatRuns ?? sheet?.rangeFormatRuns ?? sheet?.rangeRuns ?? sheet?.formattingRuns ?? null,
  );

  /**
   * Apply range format runs without enumerating the full sheet:
   * - Build a sparse index of stored cells by row
   * - For each run, visit only rows that contain stored cells
   *
   * Worst-case still depends on overlap (`#(runs ∩ cells)`), but avoids O(cells * runs)
   * when runs are sparse/non-overlapping.
   */

  /** @type {Map<number, Array<{ row: number, col: number, key: string, value: any, formula: any, format: any, cellFormat: any }>>} */
  const cellsByRow = new Map();

  const entries = Array.isArray(sheet?.cells) ? sheet.cells : [];
  for (const entry of entries) {
    const row = Number(entry?.row);
    const col = Number(entry?.col);
    if (!Number.isInteger(row) || row < 0) continue;
    if (!Number.isInteger(col) || col < 0) continue;

    const key = cellKey(row, col);
    const base = deepMerge(deepMerge(sheetDefaultFormat ?? {}, colFormats.get(col) ?? null), rowFormats.get(row) ?? null);
    const record = {
      row,
      col,
      key,
      value: entry?.value ?? null,
      formula: entry?.formula ?? null,
      format: base,
      cellFormat: extractStyleObject(entry?.format ?? entry?.style),
    };
    const bucket = cellsByRow.get(row) ?? [];
    bucket.push(record);
    cellsByRow.set(row, bucket);
  }

  if (formatRuns.length > 0 && cellsByRow.size > 0) {
    const sortedRows = Array.from(cellsByRow.keys()).sort((a, b) => a - b);
    const lowerBound = (arr, value) => {
      let lo = 0;
      let hi = arr.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (arr[mid] < value) lo = mid + 1;
        else hi = mid;
      }
      return lo;
    };
    const upperBound = (arr, value) => {
      let lo = 0;
      let hi = arr.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (arr[mid] <= value) lo = mid + 1;
        else hi = mid;
      }
      return lo;
    };

    for (const run of formatRuns) {
      const startIdx = lowerBound(sortedRows, run.startRow);
      const endIdx = upperBound(sortedRows, run.endRow);
      for (let i = startIdx; i < endIdx; i++) {
        const row = sortedRows[i];
        const bucket = cellsByRow.get(row);
        if (!bucket) continue;
        for (const record of bucket) {
          if (record.col < run.startCol || record.col > run.endCol) continue;
          record.format = deepMerge(record.format, run.format);
        }
      }
    }
  }

  for (const bucket of cellsByRow.values()) {
    for (const record of bucket) {
      const merged = deepMerge(record.format, record.cellFormat);
      cells.set(record.key, {
        value: record.value,
        formula: record.formula,
        format: normalizeFormat(merged),
      });
    }
  }

  return { cells };
}
