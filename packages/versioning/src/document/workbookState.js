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
 * @typedef {{ id: string, name: string | null }} SheetMeta
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
    sheets.push({ id, name });
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

    /** @type {Map<string, any>} */
    const cells = new Map();
    const entries = Array.isArray(sheet?.cells) ? sheet.cells : [];
    for (const entry of entries) {
      const row = Number(entry?.row);
      const col = Number(entry?.col);
      if (!Number.isInteger(row) || row < 0) continue;
      if (!Number.isInteger(col) || col < 0) continue;

      // Compute effective (layered) format without enumerating the full sheet.
      // Precedence (low -> high): sheet default, column default, row default, cell override.
      const cellFormat = extractStyleObject(entry?.format ?? entry?.style);
      const effective = normalizeFormat(
        deepMerge(
          deepMerge(
            deepMerge(sheetDefaultFormat ?? {}, colFormats.get(col) ?? null),
            rowFormats.get(row) ?? null
          ),
          cellFormat
        )
      );
      cells.set(cellKey(row, col), {
        value: entry?.value ?? null,
        formula: entry?.formula ?? null,
        format: effective,
      });
    }

    cellsBySheet.set(id, { cells });
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
