import * as Y from "yjs";
import {
  applyLayeredFormatsToCells,
  mergeCellDataIntoSheetCells,
  parseSpreadsheetCellKey,
  sheetFormatLayersFromSheetEntry,
  sheetHasLayeredFormats,
} from "./sheetState.js";

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
 * - DocumentController snapshots can store `{ rgb: "AARRGGBB" }` or `{ argb: "AARRGGBB" }`
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
   else if (json && typeof json === "object") {
     if (typeof json.rgb === "string") rgb = json.rgb;
     else if (typeof json.argb === "string") rgb = json.argb;
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
  /**
   * Build a map from sheet id -> sheet entry, picking the last matching entry by array index.
   *
   * Note: this intentionally uses the same strict id semantics as `sheetStateFromYjsDoc`:
   * only string `id` values participate. (Non-string ids are still coerced for `sheets[]`
   * metadata, but will not receive layered formatting defaults.)
   *
   * @type {Map<string, any>}
   */
  const sheetEntriesById = new Map();
  for (const entry of sheetsArray.toArray()) {
    const rawId = readYMapOrObject(entry, "id");
    const id = coerceString(rawId);
    if (!id) continue;
    const strictId = yjsValueToJson(rawId);
    if (typeof strictId === "string" && strictId) sheetEntriesById.set(strictId, entry);
    const name = coerceString(readYMapOrObject(entry, "name"));
    sheets.push({
      id,
      name,
      visibility: normalizeSheetVisibility(readYMapOrObject(entry, "visibility")),
      tabColor: normalizeTabColor(readYMapOrObject(entry, "tabColor")),
      view: sheetViewMetaFromSheetEntry(entry),
    });
    sheetOrder.push(id);
  }
  sheets.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  // Complexity note:
  // `sheetStateFromYjsDoc` scans the full Yjs `cells` map. Calling it once per sheet makes
  // workbook extraction O(#sheets * #cells) for large workbooks. Instead we scan `cells`
  // exactly once, group entries by sheet, then apply layered formats per sheet using only
  // the already-grouped cell maps (total ~O(#cells + #sheets)).
  const sheetIds = new Set(sheets.map((s) => s.id));
  const cellsMap = getMapRoot(doc, "cells");

  /** @type {Map<string, Map<string, any>>} */
  const groupedCells = new Map();
  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed?.sheetId) return;
    sheetIds.add(parsed.sheetId);
    let cells = groupedCells.get(parsed.sheetId);
    if (!cells) {
      cells = new Map();
      groupedCells.set(parsed.sheetId, cells);
    }
    mergeCellDataIntoSheetCells(cells, parsed, rawKey, cellData);
  });

  /** @type {Map<string, { cells: Map<string, any> }>} */
  const cellsBySheet = new Map();
  for (const sheetId of Array.from(sheetIds).sort()) {
    const cells = groupedCells.get(sheetId) ?? new Map();
    if (cells.size > 0) {
      const sheetEntry = sheetEntriesById.get(sheetId) ?? null;
      if (sheetEntry) {
        const layers = sheetFormatLayersFromSheetEntry(sheetEntry);
        if (sheetHasLayeredFormats(layers)) {
          applyLayeredFormatsToCells(cells, layers);
        }
      }
    }
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
