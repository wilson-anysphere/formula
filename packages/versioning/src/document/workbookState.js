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
    let startRow = Number(run.startRow ?? run.start?.row ?? run.sr ?? run.r0 ?? run.rect?.r0);
    let startCol = Number(run.startCol ?? run.start?.col ?? run.sc ?? run.c0 ?? run.rect?.c0);
    let endRow = Number(run.endRow ?? run.end?.row ?? run.er ?? run.r1 ?? run.rect?.r1);
    let endCol = Number(run.endCol ?? run.end?.col ?? run.ec ?? run.c1 ?? run.rect?.c1);
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
 * Parse per-column row interval format runs (DocumentController `formatRunsByCol` snapshot field).
 *
 * Newer document snapshots can store compressed range formatting as per-column runs:
 * `[{ col: 0, runs: [{ startRow, endRowExclusive, format }, ...] }, ...]`.
 *
 * We normalize this into the rectangle run shape used by `normalizeRangeRuns` so downstream
 * diff logic can reuse a single merge path.
 *
 * @param {any} raw
 * @returns {Array<{ startRow: number, startCol: number, endRow: number, endCol: number, format: Record<string, any> }>}
 */
function normalizeColFormatRuns(raw) {
  /** @type {Array<{ startRow: number, startCol: number, endRow: number, endCol: number, format: Record<string, any> }>} */
  const out = [];
  if (!raw) return out;

  const addRunsForCol = (colKey, rawRuns) => {
    const col = Number(colKey);
    if (!Number.isInteger(col) || col < 0) return;
    if (!Array.isArray(rawRuns)) return;
    for (const entry of rawRuns) {
      if (!entry || typeof entry !== "object") continue;
      const startRow = Number(entry.startRow ?? entry.start?.row ?? entry.sr ?? entry.r0);
      const endRowExclusiveNum = Number(entry.endRowExclusive ?? entry.endRowExcl ?? entry.erx ?? entry.r1x);
      const endRowNum = Number(entry.endRow ?? entry.end?.row ?? entry.er ?? entry.r1);
      const endRowExclusive = Number.isInteger(endRowExclusiveNum)
        ? endRowExclusiveNum
        : Number.isInteger(endRowNum)
          ? endRowNum + 1
          : NaN;
      if (!Number.isInteger(startRow) || startRow < 0) continue;
      if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;
      const format = extractStyleObject(entry.format ?? entry.style ?? entry.value);
      if (!format) continue;
      out.push({ startRow, startCol: col, endRow: endRowExclusive - 1, endCol: col, format });
    }
  };

  // Preferred encoding: array of { col, runs } entries.
  if (Array.isArray(raw)) {
    for (const entry of raw) {
      if (!entry || typeof entry !== "object") continue;
      const col = entry.col ?? entry.index ?? entry.column;
      const runs = entry.runs ?? entry.formatRuns ?? entry.segments ?? entry.items;
      addRunsForCol(col, runs);
    }
    return out;
  }

  // Also accept object keyed by column index.
  if (typeof raw === "object") {
    for (const [key, value] of Object.entries(raw)) {
      addRunsForCol(key, value?.runs ?? value?.formatRuns ?? value);
    }
  }

  return out;
}

/**
 * Excel-style worksheet visibility.
 *
 * @typedef {"visible" | "hidden" | "veryHidden"} SheetVisibility
 *
 * Small subset of per-sheet view state tracked for workbook-level diffs.
 *
 * Note: do not add large fields here (e.g. colWidths/rowHeights) — this is used
 * for version history summaries and should remain small.
 *
 * @typedef {{ frozenRows: number, frozenCols: number }} SheetViewMeta
 *
 * @typedef {{
 *   id: string,
 *   name: string | null,
 *   visibility: SheetVisibility,
 *   /**
 *    * Sheet tab color (ARGB hex, e.g. "FFFF0000") or null when cleared.
 *    *\/
 *   tabColor: string | null,
 *   view: SheetViewMeta,
 * }} SheetMeta
 * @typedef {{ id: string, cellRef: string | null, content: string | null, resolved: boolean, repliesLength: number }} CommentSummary
 *
 * @typedef {{
 *   sheets: SheetMeta[];
 *   sheetOrder: string[];
 *   metadata: Map<string, any>;
 *   namedRanges: Map<string, any>;
 *   comments: Map<string, CommentSummary>;
 *   cellsBySheet: Map<string, { cells: Map<string, any> }>;
 * }} WorkbookState
 */

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * @param {any} value
 * @returns {number}
 */
function normalizeFrozenCount(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return 0;
  return Math.max(0, Math.trunc(num));
}

/**
 * @param {any} rawVisibility
 * @returns {SheetVisibility}
 */
function normalizeSheetVisibility(rawVisibility) {
  const visibility = coerceString(rawVisibility);
  if (visibility === "visible" || visibility === "hidden" || visibility === "veryHidden") return visibility;
  return "visible";
}

/**
 * Normalize a tab color payload into an ARGB hex string (no leading "#").
 *
 * Snapshot producers vary:
 * - BranchService uses a string `"AARRGGBB"` or `null`
 * - DocumentController snapshots can store `{ rgb: "AARRGGBB" }` or `{ argb: "AARRGGBB" }`
 *
 * @param {any} raw
 * @returns {string | null}
 */
function normalizeTabColor(raw) {
  if (raw === null) return null;
  if (raw === undefined) return null;

  /** @type {string | null} */
  let rgb = null;
  if (typeof raw === "string") rgb = raw;
  else if (raw && typeof raw === "object") {
    if (typeof raw.rgb === "string") rgb = raw.rgb;
    else if (typeof raw.argb === "string") rgb = raw.argb;
  }
  if (rgb == null) return null;

  let str = rgb.trim();
  if (!str) return null;
  if (str.startsWith("#")) str = str.slice(1);

  // Allow 6-digit RGB hex by assuming opaque alpha.
  if (/^[0-9A-Fa-f]{6}$/.test(str)) str = `FF${str}`;
  if (!/^[0-9A-Fa-f]{8}$/.test(str)) return null;
  return str.toUpperCase();
}

/**
 * Extract a small view metadata object (frozen panes only) from a sheet snapshot.
 *
 * @param {any} sheet
 * @returns {SheetViewMeta}
 */
function sheetViewMetaFromSheetSnapshot(sheet) {
  const view = sheet?.view;
  if (view && typeof view === "object") {
    return {
      frozenRows: normalizeFrozenCount(view.frozenRows),
      frozenCols: normalizeFrozenCount(view.frozenCols),
    };
  }

  // Canonical DocumentController snapshot shape stores view state at top-level.
  return {
    frozenRows: normalizeFrozenCount(sheet?.frozenRows),
    frozenCols: normalizeFrozenCount(sheet?.frozenCols),
  };
}

/**
 * @param {any} value
 * @returns {boolean}
 */
function coerceBool(value) {
  return Boolean(value);
}

/**
 * @param {any} value
 * @returns {number}
 */
function repliesLength(value) {
  if (Array.isArray(value)) return value.length;
  return 0;
}

/**
 * @param {any} value
 * @param {string} fallbackId
 * @returns {CommentSummary}
 */
function commentSummaryFromValue(value, fallbackId) {
  const id = coerceString(value?.id) ?? fallbackId;
  const cellRef = coerceString(value?.cellRef);
  const content = coerceString(value?.content);
  const resolved = coerceBool(value?.resolved);
  return { id, cellRef, content, resolved, repliesLength: repliesLength(value?.replies) };
}

/**
 * @param {Uint8Array} snapshot
 * @returns {WorkbookState}
 */
export function workbookStateFromDocumentSnapshot(snapshot) {
  let parsed;
  try {
    parsed = JSON.parse(decoder.decode(snapshot));
  } catch {
    throw new Error("Invalid document snapshot: not valid JSON");
  }

  const sheetsList = Array.isArray(parsed?.sheets) ? parsed.sheets : [];
  const explicitSheetOrder = Array.isArray(parsed?.sheetOrder) ? parsed.sheetOrder : [];

  /** @type {SheetMeta[]} */
  const sheets = [];
  /** @type {string[]} */
  const sheetOrder = [];
  /** @type {Map<string, { cells: Map<string, any> }>} */
  const cellsBySheet = new Map();

  for (const sheet of sheetsList) {
    const id = coerceString(sheet?.id);
    if (!id) continue;
    const name = coerceString(sheet?.name);
    sheets.push({
      id,
      name,
      visibility: normalizeSheetVisibility(sheet?.visibility),
      tabColor: normalizeTabColor(sheet?.tabColor),
      view: sheetViewMetaFromSheetSnapshot(sheet),
    });
    sheetOrder.push(id);

    // --- Formatting defaults (layered formats) ---
    // Snapshot schema v1 stores formatting directly on each cell entry. Newer
    // schema versions can store formats as layered defaults (sheet/row/col) and
    // only keep per-cell format overrides in `entry.format`.
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
      extractStyleObject(sheet?.view?.defaultFormat) ??
      extractStyleObject(sheet?.view?.defaultStyle) ??
      extractStyleObject(sheet?.view?.sheetFormat) ??
      extractStyleObject(sheet?.view?.sheetStyle) ??
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
      sheet?.rowStyles,
      sheet?.rowFormat,
      sheet?.rowStyle,
      sheet?.view?.rowDefaults,
      sheet?.view?.rowFormats,
      sheet?.view?.rowStyles,
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
      sheet?.view?.colDefaults,
      sheet?.view?.colFormats,
      sheet?.view?.colStyles,
    ];
    /** @type {Map<number, Record<string, any>>} */
    let colFormats = new Map();
    for (const source of colFormatSources) {
      colFormats = parseIndexedFormats(source, { indexKeys: ["col", "c", "index"] });
      if (colFormats.size > 0) break;
    }

    const formatRuns = normalizeRangeRuns(
      sheet?.formatRuns ??
        sheet?.rangeFormatRuns ??
        sheet?.rangeRuns ??
        sheet?.formattingRuns ??
        sheet?.view?.formatRuns ??
        sheet?.view?.rangeFormatRuns ??
        sheet?.view?.rangeRuns ??
        sheet?.view?.formattingRuns ??
        null,
    );
    const formatRunsByCol = normalizeColFormatRuns(
      sheet?.formatRunsByCol ??
        sheet?.rangeRunsByCol ??
        sheet?.rangeFormatRunsByCol ??
        sheet?.formattingRunsByCol ??
        sheet?.view?.formatRunsByCol ??
        sheet?.view?.rangeRunsByCol ??
        sheet?.view?.rangeFormatRunsByCol ??
        sheet?.view?.formattingRunsByCol ??
        null,
    );
    const allFormatRuns = formatRunsByCol.length > 0 ? [...formatRuns, ...formatRunsByCol] : formatRuns;

    /** @type {Map<string, any>} */
    const cells = new Map();
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

    if (allFormatRuns.length > 0 && cellsByRow.size > 0) {
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

      for (const run of allFormatRuns) {
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

    cellsBySheet.set(id, { cells });
  }

  // Prefer an explicit `sheetOrder` field when present (newer DocumentController snapshots),
  // while remaining compatible with legacy snapshots that only encoded ordering via the
  // `sheets` array order.
  if (explicitSheetOrder.length > 0 && sheetOrder.length > 0) {
    const fallbackOrder = sheetOrder.slice();
    const known = new Set(fallbackOrder);
    /** @type {string[]} */
    const normalized = [];
    const seen = new Set();
    for (const raw of explicitSheetOrder) {
      if (typeof raw !== "string") continue;
      const id = raw;
      if (!known.has(id) || seen.has(id)) continue;
      seen.add(id);
      normalized.push(id);
    }
    for (const id of fallbackOrder) {
      if (seen.has(id)) continue;
      seen.add(id);
      normalized.push(id);
    }
    if (normalized.length > 0) {
      sheetOrder.length = 0;
      sheetOrder.push(...normalized);
    }
  }

  sheets.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  /** @type {Map<string, any>} */
  const metadata = new Map();
  const rawMetadata = parsed?.metadata;
  if (rawMetadata && typeof rawMetadata === "object") {
    if (Array.isArray(rawMetadata)) {
      for (const entry of rawMetadata) {
        const key = coerceString(entry?.key ?? entry?.name ?? entry?.id);
        if (!key) continue;
        metadata.set(key, structuredClone(entry?.value ?? entry));
      }
    } else {
      const keys = Object.keys(rawMetadata).sort();
      for (const key of keys) {
        metadata.set(key, structuredClone(rawMetadata[key]));
      }
    }
  }

  /** @type {Map<string, any>} */
  const namedRanges = new Map();
  const rawNamedRanges = parsed?.namedRanges;
  if (rawNamedRanges && typeof rawNamedRanges === "object") {
    if (Array.isArray(rawNamedRanges)) {
      for (const entry of rawNamedRanges) {
        const key = coerceString(entry?.name ?? entry?.id);
        if (!key) continue;
        namedRanges.set(key, structuredClone(entry));
      }
    } else {
      const keys = Object.keys(rawNamedRanges).sort();
      for (const key of keys) {
        namedRanges.set(key, structuredClone(rawNamedRanges[key]));
      }
    }
  }

  /** @type {Map<string, CommentSummary>} */
  const comments = new Map();
  const rawComments = parsed?.comments;
  if (rawComments && typeof rawComments === "object") {
    if (Array.isArray(rawComments)) {
      for (const entry of rawComments) {
        const id = coerceString(entry?.id);
        if (!id) continue;
        comments.set(id, commentSummaryFromValue(entry, id));
      }
    } else {
      const keys = Object.keys(rawComments).sort();
      for (const id of keys) {
        comments.set(id, commentSummaryFromValue(rawComments[id], id));
      }
    }
  }

  return { sheets, sheetOrder, metadata, namedRanges, comments, cellsBySheet };
}
