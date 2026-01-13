import * as Y from "yjs";
import { cellKey } from "../diff/semanticDiff.js";
import { parseSpreadsheetCellKey } from "./sheetState.js";

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

function isYMap(value) {
  if (value instanceof Y.Map) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  return (
    typeof maybe.get === "function" &&
    typeof maybe.set === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.keys === "function" &&
    typeof maybe.forEach === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYArray(value) {
  if (value instanceof Y.Array) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  return (
    typeof maybe.get === "function" &&
    typeof maybe.toArray === "function" &&
    typeof maybe.push === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYText(value) {
  if (value instanceof Y.Text) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  return (
    typeof maybe.toString === "function" &&
    typeof maybe.toDelta === "function" &&
    typeof maybe.applyDelta === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYAbstractType(value) {
  if (value instanceof Y.AbstractType) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return Boolean(maybe._map instanceof Map || maybe._start || maybe._item || maybe._length != null);
}

function replaceForeignRootType({ doc, name, existing, create }) {
  const t = create();
  t._map = existing?._map;
  t._start = existing?._start;
  t._length = existing?._length;

  const map = existing?._map;
  if (map instanceof Map) {
    map.forEach((item) => {
      for (let n = item; n !== null; n = n.left) {
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    n.parent = t;
  }

  doc.share.set(name, t);
  t._integrate?.(doc, null);
  return t;
}

/**
 * @param {Y.Doc} doc
 * @param {string} name
 */
function getMapRoot(doc, name) {
  const existing = doc.share.get(name);
  if (isYMap(existing)) return existing;
  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
  }
  if (isYAbstractType(existing)) return doc.getMap(name);
  return doc.getMap(name);
}

/**
 * @param {Y.Doc} doc
 * @param {string} name
 */
function getArrayRoot(doc, name) {
  const existing = doc.share.get(name);
  if (isYArray(existing)) return existing;
  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Array() });
  }
  if (isYAbstractType(existing)) return doc.getArray(name);
  return doc.getArray(name);
}

/**
 * @param {any} value
 * @param {string} key
 */
function readYMapOrObject(value, key) {
  if (isYMap(value)) return value.get(key);
  if (value && typeof value === "object") return value[key];
  return undefined;
}

/**
 * @param {any} value
 */
function coerceString(value) {
  if (isYText(value)) return value.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * @param {any} value
 * @returns {number}
 */
function normalizeFrozenCount(value) {
  const num = Number(yjsValueToJson(value));
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
 * - DocumentController snapshots can store `{ rgb: "AARRGGBB" }`
 *
 * @param {any} raw
 * @returns {string | null}
 */
function normalizeTabColor(raw) {
  if (raw === null) return null;
  if (raw === undefined) return null;

  const json = yjsValueToJson(raw);
  /** @type {string | null} */
  let rgb = null;
  if (typeof json === "string") rgb = json;
  else if (json && typeof json === "object" && typeof json.rgb === "string") rgb = json.rgb;
  if (rgb == null) return null;

  const cleaned = rgb.trim().replace(/^#/, "");
  if (!cleaned) return null;
  return cleaned.toUpperCase();
}

/**
 * Extract a small view metadata object (frozen panes only) from a sheet entry.
 *
 * @param {any} entry
 * @returns {SheetViewMeta}
 */
function sheetViewMetaFromSheetEntry(entry) {
  const rawView = readYMapOrObject(entry, "view");
  if (rawView !== undefined) {
    // Avoid converting the entire view object to JSON (it can contain large maps
    // like `colWidths`/`rowHeights`). We only need frozen pane counts here.
    return {
      frozenRows: normalizeFrozenCount(readYMapOrObject(rawView, "frozenRows")),
      frozenCols: normalizeFrozenCount(readYMapOrObject(rawView, "frozenCols")),
    };
  }

  // Legacy/experimental: stored as top-level keys.
  return {
    frozenRows: normalizeFrozenCount(readYMapOrObject(entry, "frozenRows")),
    frozenCols: normalizeFrozenCount(readYMapOrObject(entry, "frozenCols")),
  };
}

/**
 * Convert a Yjs value (potentially nested) into a plain JS value with stable
 * object key ordering.
 *
 * This is primarily used for diffing named ranges, where values are expected to
 * be JSON-ish.
 *
 * @param {any} value
 * @returns {any}
 */
function yjsValueToJson(value) {
  if (isYText(value)) return value.toString();
  if (isYArray(value)) return value.toArray().map((v) => yjsValueToJson(v));
  if (isYMap(value)) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Array.from(value.keys()).sort();
    for (const key of keys) {
      out[key] = yjsValueToJson(value.get(key));
    }
    return out;
  }

  if (Array.isArray(value)) return value.map((v) => yjsValueToJson(v));

  // Only canonicalize plain objects; preserve prototypes for non-plain types.
  if (value && typeof value === "object") {
    const proto = Object.getPrototypeOf(value);
    if (proto !== Object.prototype && proto !== null) {
      return structuredClone(value);
    }
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Object.keys(value).sort();
    for (const key of keys) {
      out[key] = yjsValueToJson(value[key]);
    }
    return out;
  }

  return value;
}

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
 * Sheet/cell format values may be stored inconsistently:
 * - as a bare style object
 * - wrapped in `{ format }` or `{ style }`
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
 * Normalize sparse row/col formats into a `Map<index, styleObject>`.
 *
 * Supported encodings:
 * - Y.Map / object: `{ "12": { ...format... } }`
 * - arrays: `[{ row: 12, format: {...} }, ...]` / `[{ col: 3, format: {...} }, ...]`
 * - tuple arrays: `[[12, {...format...}], ...]`
 *
 * @param {any} raw
 * @param {"row" | "col"} axis
 * @returns {Map<number, Record<string, any>>}
 */
function parseIndexedFormats(raw, axis) {
  /** @type {Map<number, Record<string, any>>} */
  const out = new Map();
  const json = yjsValueToJson(raw);
  if (json == null) return out;

  if (Array.isArray(json)) {
    for (const entry of json) {
      let index;
      let formatValue;
      if (Array.isArray(entry)) {
        index = entry[0];
        formatValue = entry[1];
      } else if (entry && typeof entry === "object") {
        index = entry[axis] ?? entry.index;
        formatValue = entry.format ?? entry.style ?? entry.value;
      } else {
        continue;
      }

      const idx = Number(index);
      if (!Number.isInteger(idx) || idx < 0) continue;
      const style = extractStyleObject(yjsValueToJson(formatValue));
      if (!style) continue;
      out.set(idx, style);
    }
    return out;
  }

  if (typeof json === "object") {
    for (const [key, value] of Object.entries(json)) {
      const idx = Number(key);
      if (!Number.isInteger(idx) || idx < 0) continue;
      const style = extractStyleObject(yjsValueToJson(value?.format ?? value?.style ?? value));
      if (!style) continue;
      out.set(idx, style);
    }
  }

  return out;
}

/**
 * Normalize sparse per-column row interval format runs into a `Map<col, runs[]>`.
 *
 * The collab binder stores range-run formatting on sheet metadata as:
 *   formatRunsByCol: { "0": [{ startRow, endRowExclusive, format }, ...], ... }
 *
 * This mirrors `DocumentController`'s internal `sheet.formatRunsByCol` data, but runs
 * store style objects (not style ids) because style ids are per-client.
 *
 * @param {any} raw
 * @returns {Map<number, Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>>}
 */
function parseFormatRunsByCol(raw) {
  /** @type {Map<number, Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>>} */
  const out = new Map();
  const json = yjsValueToJson(raw);
  if (!json) return out;

  /**
   * @param {any} colKey
   * @param {any} rawRuns
   */
  const addRunsForCol = (colKey, rawRuns) => {
    const col = Number(colKey);
    if (!Number.isInteger(col) || col < 0) return;

    const list = Array.isArray(rawRuns)
      ? rawRuns
      : isPlainObject(rawRuns)
        ? rawRuns.runs ?? rawRuns.formatRuns ?? rawRuns.segments ?? rawRuns.items ?? []
        : [];
    if (!Array.isArray(list) || list.length === 0) return;

    /** @type {Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>} */
    const runs = [];
    for (const entry of list) {
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

      const format = extractStyleObject(yjsValueToJson(entry.format ?? entry.style ?? entry.value));
      if (!format) continue;
      runs.push({ startRow, endRowExclusive, format });
    }

    runs.sort((a, b) => a.startRow - b.startRow);
    if (runs.length > 0) out.set(col, runs);
  };

  // Preferred encoding: object keyed by column index.
  if (typeof json === "object" && !Array.isArray(json)) {
    for (const [key, value] of Object.entries(json)) {
      addRunsForCol(key, value);
    }
    return out;
  }

  // Also accept array encodings: [{ col, runs }, ...] or [[col, runs], ...].
  if (Array.isArray(json)) {
    for (const entry of json) {
      if (Array.isArray(entry)) {
        addRunsForCol(entry[0], entry[1]);
        continue;
      }
      if (entry && typeof entry === "object") {
        const col = entry.col ?? entry.index ?? entry.column;
        const runs = entry.runs ?? entry.formatRuns ?? entry.segments ?? entry.items;
        addRunsForCol(col, runs);
      }
    }
  }

  return out;
}

/**
 * @param {any} cellData
 */
function extractCell(cellData) {
  if (isYMap(cellData)) {
    const enc = cellData.get("enc");
    return {
      ...(enc !== null && enc !== undefined
        ? { enc: yjsValueToJson(enc), value: null, formula: null }
        : { value: cellData.get("value") ?? null, formula: cellData.get("formula") ?? null }),
      format: cellData.get("format") ?? cellData.get("style") ?? null,
    };
  }
  if (cellData && typeof cellData === "object") {
    const enc = cellData.enc;
    return {
      ...(enc !== null && enc !== undefined
        ? { enc: yjsValueToJson(enc), value: null, formula: null }
        : { value: cellData.value ?? null, formula: cellData.formula ?? null }),
      format: cellData.format ?? cellData.style ?? null,
    };
  }
  return { value: cellData ?? null, formula: null, format: null };
}

/**
 * Compute effective per-cell formats for a sheet, applying layered formats:
 * sheet default -> column defaults -> row defaults -> range runs -> per-cell overrides.
 *
 * Mutates the `cells` map in-place to set `cell.format` to its effective value.
 *
 * @param {Y.Doc} doc
 * @param {string} sheetId
 * @param {any} sheetEntry
 * @param {Map<string, any>} cells
 */
function applyLayeredFormattingToCells(doc, sheetId, sheetEntry, cells) {
  if (!sheetId) return;
  if (!sheetEntry) return;
  if (!cells || cells.size === 0) return;

  const view = sheetEntry ? readYMapOrObject(sheetEntry, "view") : null;
  const viewJson = view != null ? yjsValueToJson(view) : null;

  const rawDefaultFormat = sheetEntry != null ? readYMapOrObject(sheetEntry, "defaultFormat") : undefined;
  const rawRowFormats = sheetEntry != null ? readYMapOrObject(sheetEntry, "rowFormats") : undefined;
  const rawColFormats = sheetEntry != null ? readYMapOrObject(sheetEntry, "colFormats") : undefined;
  const rawFormatRunsByCol = sheetEntry != null ? readYMapOrObject(sheetEntry, "formatRunsByCol") : undefined;

  const sheetDefaultFormat =
    extractStyleObject(
      yjsValueToJson(rawDefaultFormat !== undefined ? rawDefaultFormat : viewJson?.defaultFormat),
    ) ?? null;
  const rowFormats = parseIndexedFormats(rawRowFormats !== undefined ? rawRowFormats : viewJson?.rowFormats, "row");
  const colFormats = parseIndexedFormats(rawColFormats !== undefined ? rawColFormats : viewJson?.colFormats, "col");
  const formatRunsByCol = parseFormatRunsByCol(
    rawFormatRunsByCol !== undefined ? rawFormatRunsByCol : viewJson?.formatRunsByCol,
  );

  if (!sheetDefaultFormat && rowFormats.size === 0 && colFormats.size === 0 && formatRunsByCol.size === 0) {
    return;
  }

  /**
   * Find the run containing `row` (half-open interval `[startRow, endRowExclusive)`).
   *
   * DocumentController guarantees these runs are sorted + non-overlapping, which lets us
   * do a binary search.
   *
   * @param {Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>} runs
   * @param {number} row
   */
  const findRunForRow = (runs, row) => {
    let lo = 0;
    let hi = runs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const run = runs[mid];
      if (row < run.startRow) {
        hi = mid - 1;
      } else if (row >= run.endRowExclusive) {
        lo = mid + 1;
      } else {
        return run;
      }
    }
    return null;
  };

  for (const [key, cell] of cells.entries()) {
    const m = String(key).match(/^r(\d+)c(\d+)$/);
    if (!m) continue;
    const row = Number(m[1]);
    const col = Number(m[2]);
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;

    const cellFormat = extractStyleObject(yjsValueToJson(cell?.format ?? cell?.style));
    let merged = deepMerge(
      deepMerge(sheetDefaultFormat ?? {}, colFormats.get(col) ?? null),
      rowFormats.get(row) ?? null,
    );
    const runs = formatRunsByCol.get(col);
    if (runs && runs.length > 0) {
      const run = findRunForRow(runs, row);
      if (run) merged = deepMerge(merged, run.format);
    }
    merged = deepMerge(merged, cellFormat);
    cell.format = normalizeFormat(merged);
  }
}

/**
 * @param {any} value
 * @param {string} mapKey
 * @returns {CommentSummary}
 */
function commentSummaryFromValue(value, mapKey) {
  const id = mapKey;
  const cellRef = coerceString(readYMapOrObject(value, "cellRef"));
  const content = coerceString(readYMapOrObject(value, "content"));
  const resolved = Boolean(readYMapOrObject(value, "resolved"));

  const replies = readYMapOrObject(value, "replies");
  let repliesLength = 0;
  if (isYArray(replies)) repliesLength = replies.length;
  else if (Array.isArray(replies)) repliesLength = replies.length;

  return { id, cellRef, content, resolved, repliesLength };
}

/**
 * Iterate legacy list items stored on a Map root (i.e. CRDT list items with
 * `parentSub === null`).
 *
 * This happens if a document originally used the legacy Array schema, but the
 * root was later instantiated as a Map (e.g. by calling `doc.getMap("comments")`
 * first while the root was still an `AbstractType` placeholder). In that case
 * the comments still exist in the CRDT but are invisible via `map.keys()`.
 *
 * @param {any} mapType
 * @returns {Y.Map<any>[]} legacy comment maps
 */
function legacyListCommentsFromMapRoot(mapType) {
  /** @type {Y.Map<any>[]} */
  const out = [];
  let item = mapType?._start ?? null;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      const content = item.content?.getContent?.() ?? [];
      for (const value of content) {
        if (isYMap(value)) out.push(value);
      }
    }
    item = item.right;
  }
  return out;
}

/**
 * Iterate map entries stored on an Array root (i.e. `parentSub !== null` items
 * reachable via `_map`).
 *
 * This can happen in mixed-schema situations where some clients treat a root as
 * an Array while others treat it as a Map. Map entries are not visible via
 * `array.toArray()` but still exist in the CRDT.
 *
 * @param {any} arrayType
 * @returns {Array<[string, any]>}
 */
function mapEntriesFromArrayRoot(arrayType) {
  /** @type {Array<[string, any]>} */
  const out = [];
  const map = arrayType?._map;
  if (!(map instanceof Map)) return out;
  for (const [key, item] of map.entries()) {
    if (!item || item.deleted) continue;
    const content = item.content?.getContent?.() ?? [];
    if (content.length === 0) continue;
    out.push([key, content[content.length - 1]]);
  }
  out.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
  return out;
}

/**
 * Extract a deterministic workbook state from a Yjs doc snapshot.
 *
 * @param {Y.Doc} doc
 * @returns {WorkbookState}
 */
export function workbookStateFromYjsDoc(doc) {
  const sheetsArray = getArrayRoot(doc, "sheets");
  /** @type {SheetMeta[]} */
  const sheets = [];
  /** @type {string[]} */
  const sheetOrder = [];
  /** @type {Map<string, any>} */
  const sheetEntriesById = new Map();
  for (const entry of sheetsArray.toArray()) {
    const id = coerceString(readYMapOrObject(entry, "id"));
    if (!id) continue;
    const name = coerceString(readYMapOrObject(entry, "name"));
    sheets.push({
      id,
      name,
      visibility: normalizeSheetVisibility(readYMapOrObject(entry, "visibility")),
      tabColor: normalizeTabColor(readYMapOrObject(entry, "tabColor")),
      view: sheetViewMetaFromSheetEntry(entry),
    });
    sheetOrder.push(id);
    // Deterministic choice: pick the last matching entry by index (mirrors binder behavior).
    sheetEntriesById.set(id, entry);
  }
  sheets.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  const sheetIds = new Set(sheets.map((s) => s.id));
  const cellsMap = getMapRoot(doc, "cells");

  /** @type {Map<string, Map<string, any>>} */
  const rawCellsBySheet = new Map();
  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed?.sheetId) return;
    sheetIds.add(parsed.sheetId);

    let sheetCells = rawCellsBySheet.get(parsed.sheetId);
    if (!sheetCells) {
      sheetCells = new Map();
      rawCellsBySheet.set(parsed.sheetId, sheetCells);
    }

    const key = cellKey(parsed.row, parsed.col);
    const cell = extractCell(cellData);
    const existing = sheetCells.get(key);

    // If any representation of this coordinate is encrypted (e.g. stored under a
    // legacy key encoding), treat the cell as encrypted and do not allow plaintext
    // duplicates to overwrite the ciphertext.
    const isCanonical = rawKey === `${parsed.sheetId}:${parsed.row}:${parsed.col}`;

    const enc = cell?.enc;
    const isEncrypted = enc !== null && enc !== undefined;
    const existingEnc = existing?.enc;
    const existingIsEncrypted = existingEnc !== null && existingEnc !== undefined;

    if (isEncrypted) {
      if (!existing || !existingIsEncrypted || isCanonical) {
        // Preserve any existing format metadata if the preferred encrypted record
        // lacks it (e.g. canonical key created during encryption while a legacy
        // key still carries the existing format).
        if (existing?.format != null && cell.format == null) {
          sheetCells.set(key, { ...cell, format: existing.format });
        } else {
          sheetCells.set(key, cell);
        }
      } else if (existing.format == null && cell.format != null) {
        sheetCells.set(key, { ...existing, format: cell.format });
      }
      return;
    }

    if (existingIsEncrypted) {
      // Preserve ciphertext, but allow plaintext duplicates to contribute format
      // metadata when the encrypted record lacks it.
      if (existing.format == null && cell.format != null) {
        sheetCells.set(key, { ...existing, format: cell.format });
      }
      return;
    }

    sheetCells.set(key, cell);
  });

  /** @type {Map<string, { cells: Map<string, any> }>} */
  const cellsBySheet = new Map();
  for (const sheetId of Array.from(sheetIds).sort()) {
    const cells = rawCellsBySheet.get(sheetId) ?? new Map();
    applyLayeredFormattingToCells(doc, sheetId, sheetEntriesById.get(sheetId) ?? null, cells);
    cellsBySheet.set(sheetId, { cells });
  }

  /** @type {Map<string, any>} */
  const metadata = new Map();
  if (doc.share.has("metadata")) {
    try {
      const metadataMap = getMapRoot(doc, "metadata");
      for (const key of Array.from(metadataMap.keys()).sort()) {
        metadata.set(key, yjsValueToJson(metadataMap.get(key)));
      }
    } catch {
      // Ignore: unsupported root type.
    }
  }

  /** @type {Map<string, any>} */
  const namedRanges = new Map();
  if (doc.share.has("namedRanges")) {
    try {
      const namedRangesMap = getMapRoot(doc, "namedRanges");
      for (const key of Array.from(namedRangesMap.keys()).sort()) {
        namedRanges.set(key, yjsValueToJson(namedRangesMap.get(key)));
      }
    } catch {
      // Ignore: unsupported root type.
    }
  }

  /** @type {Map<string, CommentSummary>} */
  const comments = new Map();
  if (doc.share.has("comments")) {
    // Yjs root types are schema-defined: you must know whether a key is a Map or
    // Array. When applying updates into a fresh Doc, root types can temporarily
    // appear as a generic `AbstractType` until a constructor is chosen.
    //
    // Importantly, calling `doc.getMap("comments")` on an Array-backed root can
    // define it as a Map and make the array content inaccessible. To support
    // both historical schemas (Map or Array) we peek at the underlying state
    // before choosing a constructor.
    const existing = doc.share.get("comments");

    // Canonical schema: Y.Map keyed by comment id.
    if (isYMap(existing)) {
      const byId = new Map();
      for (const key of Array.from(existing.keys()).sort()) {
        byId.set(key, commentSummaryFromValue(existing.get(key), key));
      }
      // Recovery: legacy list items stored on a Map root (see helper above).
      for (const item of legacyListCommentsFromMapRoot(existing)) {
        const id = coerceString(readYMapOrObject(item, "id"));
        if (!id) continue;
        if (byId.has(id)) continue;
        byId.set(id, commentSummaryFromValue(item, id));
      }
      for (const id of Array.from(byId.keys()).sort()) {
        comments.set(id, byId.get(id));
      }
    } else if (isYArray(existing)) {
      /** @type {Map<string, CommentSummary>} */
      const byId = new Map();
      // Recovery: map entries stored on an Array root (mixed schema).
      for (const [key, value] of mapEntriesFromArrayRoot(existing)) {
        byId.set(key, commentSummaryFromValue(value, key));
      }
      for (const item of existing.toArray()) {
        const id = coerceString(readYMapOrObject(item, "id"));
        if (!id) continue;
        if (byId.has(id)) continue;
        byId.set(id, commentSummaryFromValue(item, id));
      }
      for (const id of Array.from(byId.keys()).sort()) {
        comments.set(id, byId.get(id));
      }
    } else {
      const placeholder = existing;
      const hasStart = placeholder?._start != null; // sequence item => likely array
      const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
      const kind = hasStart && mapSize === 0 ? "array" : "map";

      if (kind === "map") {
        const commentsMap = getMapRoot(doc, "comments");
        const byId = new Map();
        for (const key of Array.from(commentsMap.keys()).sort()) {
          byId.set(key, commentSummaryFromValue(commentsMap.get(key), key));
        }
        for (const item of legacyListCommentsFromMapRoot(commentsMap)) {
          const id = coerceString(readYMapOrObject(item, "id"));
          if (!id) continue;
          if (byId.has(id)) continue;
          byId.set(id, commentSummaryFromValue(item, id));
        }
        for (const id of Array.from(byId.keys()).sort()) {
          comments.set(id, byId.get(id));
        }
      } else {
        const commentsArray = getArrayRoot(doc, "comments");
        /** @type {[string, CommentSummary][]} */
        const entries = [];
        for (const item of commentsArray.toArray()) {
          const id = coerceString(readYMapOrObject(item, "id"));
          if (!id) continue;
          entries.push([id, commentSummaryFromValue(item, id)]);
        }
        entries.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
        for (const [id, summary] of entries) {
          comments.set(id, summary);
        }
      }
    }
  }

  return { sheets, sheetOrder, metadata, namedRanges, comments, cellsBySheet };
}

/**
 * @param {Uint8Array} snapshot
 * @returns {WorkbookState}
 */
export function workbookStateFromYjsSnapshot(snapshot) {
  const doc = new Y.Doc();
  Y.applyUpdate(doc, snapshot);
  return workbookStateFromYjsDoc(doc);
}
