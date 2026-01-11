import * as Y from "yjs";
import { parseSpreadsheetCellKey, sheetStateFromYjsDoc } from "./sheetState.js";

/**
 * @typedef {{ id: string, name: string | null }} SheetMeta
 * @typedef {{ id: string, cellRef: string | null, content: string | null, resolved: boolean, repliesLength: number }} CommentSummary
 *
 * @typedef {{
 *   sheets: SheetMeta[];
 *   sheetOrder: string[];
 *   namedRanges: Map<string, any>;
 *   comments: Map<string, CommentSummary>;
 *   cellsBySheet: Map<string, { cells: Map<string, any> }>;
 * }} WorkbookState
 */

/**
 * @param {any} value
 * @param {string} key
 */
function readYMapOrObject(value, key) {
  if (value instanceof Y.Map) return value.get(key);
  if (value && typeof value === "object") return value[key];
  return undefined;
}

/**
 * @param {any} value
 */
function coerceString(value) {
  if (value instanceof Y.Text) return value.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
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
  if (value instanceof Y.Text) return value.toString();
  if (value instanceof Y.Array) return value.toArray().map((v) => yjsValueToJson(v));
  if (value instanceof Y.Map) {
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
  if (replies instanceof Y.Array) repliesLength = replies.length;
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
        if (value instanceof Y.Map) out.push(value);
      }
    }
    item = item.right;
  }
  return out;
}

/**
 * Extract a deterministic workbook state from a Yjs doc snapshot.
 *
 * @param {Y.Doc} doc
 * @returns {WorkbookState}
 */
export function workbookStateFromYjsDoc(doc) {
  const sheetsArray = doc.getArray("sheets");
  /** @type {SheetMeta[]} */
  const sheets = [];
  /** @type {string[]} */
  const sheetOrder = [];
  for (const entry of sheetsArray.toArray()) {
    const id = coerceString(readYMapOrObject(entry, "id"));
    if (!id) continue;
    const name = coerceString(readYMapOrObject(entry, "name"));
    sheets.push({ id, name });
    sheetOrder.push(id);
  }
  sheets.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));

  const sheetIds = new Set(sheets.map((s) => s.id));
  const cellsMap = doc.getMap("cells");
  cellsMap.forEach((_, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed?.sheetId) return;
    sheetIds.add(parsed.sheetId);
  });

  /** @type {Map<string, { cells: Map<string, any> }>} */
  const cellsBySheet = new Map();
  for (const sheetId of Array.from(sheetIds).sort()) {
    cellsBySheet.set(sheetId, sheetStateFromYjsDoc(doc, { sheetId }));
  }

  /** @type {Map<string, any>} */
  const namedRanges = new Map();
  if (doc.share.has("namedRanges")) {
    try {
      const namedRangesMap = doc.getMap("namedRanges");
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
    if (existing instanceof Y.Map) {
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
    } else if (existing instanceof Y.Array) {
      /** @type {[string, CommentSummary][]} */
      const entries = [];
      for (const item of existing.toArray()) {
        const id = coerceString(readYMapOrObject(item, "id"));
        if (!id) continue;
        entries.push([id, commentSummaryFromValue(item, id)]);
      }
      entries.sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
      for (const [id, summary] of entries) {
        comments.set(id, summary);
      }
    } else {
      const placeholder = existing;
      const hasStart = placeholder?._start != null; // sequence item => likely array
      const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
      const kind = hasStart && mapSize === 0 ? "array" : "map";

      if (kind === "map") {
        const commentsMap = doc.getMap("comments");
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
        const commentsArray = doc.getArray("comments");
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

  return { sheets, sheetOrder, namedRanges, comments, cellsBySheet };
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
