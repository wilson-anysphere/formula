import * as Y from "yjs";
import { parseSpreadsheetCellKey, sheetStateFromYjsDoc } from "./sheetState.js";

/**
 * @typedef {{ id: string, name: string | null }} SheetMeta
 * @typedef {{ id: string, cellRef: string | null, content: string | null, resolved: boolean, repliesLength: number }} CommentSummary
 *
 * @typedef {{
 *   sheets: SheetMeta[];
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
 * Extract a deterministic workbook state from a Yjs doc snapshot.
 *
 * @param {Y.Doc} doc
 * @returns {WorkbookState}
 */
export function workbookStateFromYjsDoc(doc) {
  const sheetsArray = doc.getArray("sheets");
  /** @type {SheetMeta[]} */
  const sheets = [];
  for (const entry of sheetsArray.toArray()) {
    const id = coerceString(readYMapOrObject(entry, "id"));
    if (!id) continue;
    const name = coerceString(readYMapOrObject(entry, "name"));
    sheets.push({ id, name });
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
  const namedRangesMap = doc.getMap("namedRanges");
  for (const key of Array.from(namedRangesMap.keys()).sort()) {
    namedRanges.set(key, yjsValueToJson(namedRangesMap.get(key)));
  }

  /** @type {Map<string, CommentSummary>} */
  const comments = new Map();
  const commentsMap = doc.getMap("comments");
  for (const key of Array.from(commentsMap.keys()).sort()) {
    comments.set(key, commentSummaryFromValue(commentsMap.get(key), key));
  }

  return { sheets, namedRanges, comments, cellsBySheet };
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
