import * as Y from "yjs";

import { normalizeCell } from "../cell.js";
import { normalizeDocumentState } from "../state.js";
import { a1ToRowCol, rowColToA1 } from "./a1.js";

/**
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../types.js").Cell} Cell
 * @typedef {import("../types.js").CellMap} CellMap
 * @typedef {import("../types.js").SheetMeta} SheetMeta
 */

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * @param {any} value
 * @param {string} key
 */
function readYMapOrObject(value, key) {
  if (value instanceof Y.Map) return value.get(key);
  if (isRecord(value)) return value[key];
  return undefined;
}

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  if (value instanceof Y.Text) return value.toString();
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * Normalize a formula to the canonical representation used throughout the app:
 * - trim leading whitespace
 * - ensure it starts with "="
 *
 * @param {unknown} value
 * @returns {string | null}
 */
function normalizeFormula(value) {
  if (value == null) return null;
  const trimmed = String(value).trimStart();
  if (trimmed === "") return null;
  return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
}

/**
 * Convert a Yjs value (potentially nested) into a plain JS value.
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
    for (const key of keys) out[key] = yjsValueToJson(value.get(key));
    return out;
  }

  if (Array.isArray(value)) return value.map((v) => yjsValueToJson(v));

  if (isRecord(value)) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Object.keys(value).sort();
    for (const key of keys) out[key] = yjsValueToJson(value[key]);
    return out;
  }

  return value;
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
 * Delete any legacy list items (sequence entries with `parentSub === null`) from
 * an instantiated map root.
 *
 * @param {any} transaction
 * @param {any} mapType
 */
function deleteLegacyListItemsFromMapRoot(transaction, mapType) {
  let item = mapType?._start ?? null;
  while (item) {
    if (!item.deleted && item.parentSub === null) {
      item.delete(transaction);
    }
    item = item.right;
  }
}

/**
 * Delete any map entries (keyed items) from an instantiated array root.
 *
 * This can happen if a map schema was instantiated as an Array: map entries are
 * stored in `type._map` and are invisible to `array.toArray()`.
 *
 * @param {any} transaction
 * @param {any} arrayType
 */
function deleteMapEntriesFromArrayRoot(transaction, arrayType) {
  const map = arrayType?._map;
  if (!(map instanceof Map)) return;
  for (const item of map.values()) {
    if (!item?.deleted) item.delete(transaction);
  }
}

/**
 * Parse a spreadsheet cell key. Supports:
 * - `${sheetId}:${row}:${col}`
 * - `${sheetId}:${row},${col}`
 * - `r{row}c{col}` (unit-test convenience; assumed to be in Sheet1)
 *
 * @param {string} key
 * @returns {{ sheetId: string, row: number, col: number } | null}
 */
function parseSpreadsheetCellKey(key) {
  const colon = key.split(":");
  if (colon.length === 3) {
    const sheetId = colon[0];
    const row = Number(colon[1]);
    const col = Number(colon[2]);
    if (!sheetId) return null;
    if (!Number.isInteger(row) || row < 0) return null;
    if (!Number.isInteger(col) || col < 0) return null;
    return { sheetId, row, col };
  }

  if (colon.length === 2) {
    const sheetId = colon[0];
    if (!sheetId) return null;
    const m = colon[1].match(/^(\d+),(\d+)$/);
    if (!m) return null;
    const row = Number(m[1]);
    const col = Number(m[2]);
    if (!Number.isInteger(row) || row < 0) return null;
    if (!Number.isInteger(col) || col < 0) return null;
    return { sheetId, row, col };
  }

  const m = key.match(/^r(\d+)c(\d+)$/);
  if (m) {
    const row = Number(m[1]);
    const col = Number(m[2]);
    if (!Number.isInteger(row) || row < 0) return null;
    if (!Number.isInteger(col) || col < 0) return null;
    return { sheetId: "Sheet1", row, col };
  }

  return null;
}

/**
 * @param {any} cellData
 * @returns {Cell | null}
 */
function cellFromYjsValue(cellData) {
  /** @type {Cell} */
  const cell = {};
  if (cellData instanceof Y.Map) {
    const formula = normalizeFormula(cellData.get("formula"));
    const value = cellData.get("value");
    const format = cellData.get("format") ?? cellData.get("style");
    if (formula) cell.formula = formula;
    else if (value !== null && value !== undefined) cell.value = value;
    if (format !== null && format !== undefined) cell.format = yjsValueToJson(format);
    return normalizeCell(cell);
  }

  if (isRecord(cellData)) {
    const formula = normalizeFormula(cellData.formula);
    const value = cellData.value;
    const format = cellData.format ?? cellData.style;
    if (formula) cell.formula = formula;
    else if (value !== null && value !== undefined) cell.value = value;
    if (format !== null && format !== undefined) cell.format = yjsValueToJson(format);
    return normalizeCell(cell);
  }

  if (cellData !== null && cellData !== undefined) {
    cell.value = cellData;
    return normalizeCell(cell);
  }

  return null;
}

/**
 * Extract a deterministic BranchService {@link DocumentState} from a Yjs doc.
 *
 * @param {Y.Doc} doc
 * @returns {DocumentState}
 */
export function branchStateFromYjsDoc(doc) {
  const sheetsArray = doc.getArray("sheets");
  /** @type {Record<string, SheetMeta>} */
  const metaById = {};
  /** @type {string[]} */
  const order = [];

  for (const entry of sheetsArray.toArray()) {
    const id = coerceString(readYMapOrObject(entry, "id"));
    if (!id) continue;
    const name = coerceString(readYMapOrObject(entry, "name"));
    metaById[id] = { id, name };
    order.push(id);
  }

  const cellsMap = doc.getMap("cells");

  /** @type {Record<string, CellMap>} */
  const cells = {};

  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed?.sheetId) return;
    const sheetId = parsed.sheetId;
    if (!cells[sheetId]) cells[sheetId] = {};

    if (!metaById[sheetId]) {
      // Sheets can exist implicitly via cells even when the "sheets" root isn't populated.
      metaById[sheetId] = { id: sheetId, name: sheetId };
    }

    const addr = rowColToA1(parsed.row, parsed.col);
    const cell = cellFromYjsValue(cellData);
    if (cell) cells[sheetId][addr] = cell;
  });

  // Ensure every sheet in metadata has a cell map (even empty).
  for (const sheetId of Object.keys(metaById)) {
    if (!cells[sheetId]) cells[sheetId] = {};
  }

  // Append any implicit sheets not present in the order.
  const seen = new Set(order);
  for (const sheetId of Object.keys(metaById).sort()) {
    if (seen.has(sheetId)) continue;
    order.push(sheetId);
    seen.add(sheetId);
  }

  /** @type {Record<string, any>} */
  const namedRanges = {};
  if (doc.share.has("namedRanges")) {
    try {
      const namedRangesMap = doc.getMap("namedRanges");
      for (const key of Array.from(namedRangesMap.keys()).sort()) {
        namedRanges[key] = yjsValueToJson(namedRangesMap.get(key));
      }
    } catch {
      // Ignore: unsupported root type.
    }
  }

  /** @type {Record<string, any>} */
  const comments = {};
  if (doc.share.has("comments")) {
    // Yjs root types are schema-defined: you must know whether a key is a Map or
    // Array. When applying updates into a fresh Doc, root types can temporarily
    // appear as a generic `AbstractType` until a constructor is chosen.
    //
    // Importantly, calling `doc.getArray("comments")` on a Map-backed root (or
    // vice versa) can throw and/or make legacy content inaccessible. To support
    // both historical schemas (Map or Array) we inspect the root value first.
    const existing = doc.share.get("comments");

    if (existing instanceof Y.Map) {
      /** @type {Map<string, any>} */
      const byId = new Map();
      for (const id of Array.from(existing.keys()).sort()) {
        byId.set(id, yjsValueToJson(existing.get(id)));
      }
      // Recovery: legacy list items stored on a Map root.
      for (const item of legacyListCommentsFromMapRoot(existing)) {
        const id = coerceString(readYMapOrObject(item, "id"));
        if (!id) continue;
        if (byId.has(id)) continue;
        byId.set(id, yjsValueToJson(item));
      }
      for (const id of Array.from(byId.keys()).sort()) {
        comments[id] = byId.get(id);
      }
    } else if (existing instanceof Y.Array) {
      /** @type {Array<{ id: string, value: any }>} */
      const entries = [];
      for (const item of existing.toArray()) {
        const id = coerceString(readYMapOrObject(item, "id"));
        if (!id) continue;
        entries.push({ id, value: yjsValueToJson(item) });
      }
      entries.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));
      for (const entry of entries) comments[entry.id] = entry.value;
    } else {
      const placeholder = existing;
      const hasStart = placeholder?._start != null;
      const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
      const kind = hasStart && mapSize === 0 ? "array" : "map";

      if (kind === "map") {
        const commentsMap = doc.getMap("comments");
        /** @type {Map<string, any>} */
        const byId = new Map();
        for (const id of Array.from(commentsMap.keys()).sort()) {
          byId.set(id, yjsValueToJson(commentsMap.get(id)));
        }
        for (const item of legacyListCommentsFromMapRoot(commentsMap)) {
          const id = coerceString(readYMapOrObject(item, "id"));
          if (!id) continue;
          if (byId.has(id)) continue;
          byId.set(id, yjsValueToJson(item));
        }
        for (const id of Array.from(byId.keys()).sort()) {
          comments[id] = byId.get(id);
        }
      } else {
        const commentsArray = doc.getArray("comments");
        /** @type {Array<{ id: string, value: any }>} */
        const entries = [];
        for (const item of commentsArray.toArray()) {
          const id = coerceString(readYMapOrObject(item, "id"));
          if (!id) continue;
          entries.push({ id, value: yjsValueToJson(item) });
        }
        entries.sort((a, b) => (a.id < b.id ? -1 : a.id > b.id ? 1 : 0));
        for (const entry of entries) comments[entry.id] = entry.value;
      }
    }
  }

  return normalizeDocumentState({
    schemaVersion: 1,
    sheets: { order, metaById },
    cells,
    namedRanges,
    comments,
  });
}

/**
 * Replace Yjs doc contents from a BranchService {@link DocumentState}.
 *
 * This is used for branch checkout / applying merged state back into a shared
 * collaborative document.
 *
 * @param {Y.Doc} doc
 * @param {DocumentState} state
 * @param {{ origin?: any }} [opts]
 */
export function applyBranchStateToYjsDoc(doc, state, opts = {}) {
  const normalized = normalizeDocumentState(state);

  /**
   * @param {any} value
   */
  function yMapFromJsonObject(value) {
    const map = new Y.Map();
    if (!isRecord(value)) return map;
    const keys = Object.keys(value).sort();
    for (const key of keys) {
      const v = value[key];
      if (key === "replies") {
        const replies = new Y.Array();
        if (Array.isArray(v)) {
          for (const reply of v) {
            if (isRecord(reply)) replies.push([yMapFromJsonObject(reply)]);
          }
        }
        map.set("replies", replies);
        continue;
      }
      // Mentions are stored as plain arrays in the collab comment schema.
      if (key === "mentions") {
        map.set("mentions", Array.isArray(v) ? structuredClone(v) : []);
        continue;
      }
      map.set(key, structuredClone(v));
    }
    return map;
  }

  doc.transact(
    (transaction) => {
      // --- Sheets ---
      const sheetsArray = doc.getArray("sheets");
      if (sheetsArray.length > 0) sheetsArray.delete(0, sheetsArray.length);
      for (const sheetId of normalized.sheets.order) {
        const meta = normalized.sheets.metaById[sheetId];
        const entry = new Y.Map();
        entry.set("id", sheetId);
        entry.set("name", meta?.name ?? null);
        sheetsArray.push([entry]);
      }

      // --- Cells ---
      const cellsMap = doc.getMap("cells");
      for (const key of Array.from(cellsMap.keys())) {
        cellsMap.delete(key);
      }

      for (const [sheetId, sheet] of Object.entries(normalized.cells)) {
        for (const [addr, cell] of Object.entries(sheet ?? {})) {
          const normalizedCell = normalizeCell(cell);
          if (!normalizedCell) continue;
          const { row, col } = a1ToRowCol(addr);
          const key = `${sheetId}:${row}:${col}`;

          const yCell = new Y.Map();
          const formula = normalizeFormula(normalizedCell.formula);
          if (formula) {
            yCell.set("formula", formula);
            yCell.set("value", null);
          } else if (normalizedCell.value !== undefined) {
            yCell.set("value", normalizedCell.value);
          }
          if (normalizedCell.format != null) yCell.set("format", structuredClone(normalizedCell.format));
          cellsMap.set(key, yCell);
        }
      }

      // --- Named ranges ---
      const namedRangesMap = doc.getMap("namedRanges");
      for (const key of Array.from(namedRangesMap.keys())) namedRangesMap.delete(key);
      for (const [key, value] of Object.entries(normalized.namedRanges ?? {})) {
        namedRangesMap.set(key, structuredClone(value));
      }

      // --- Comments ---
      const existing = doc.share.get("comments");
      const commentsKind =
        existing instanceof Y.Array
          ? "array"
          : existing instanceof Y.Map
            ? "map"
            : (() => {
                const placeholder = existing;
                const hasStart = placeholder?._start != null;
                const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
                return placeholder && hasStart && mapSize === 0 ? "array" : "map";
              })();

      if (commentsKind === "array") {
        const commentsArray = doc.getArray("comments");
        deleteMapEntriesFromArrayRoot(transaction, commentsArray);
        if (commentsArray.length > 0) commentsArray.delete(0, commentsArray.length);

        const ids = Object.keys(normalized.comments ?? {}).sort();
        for (const id of ids) {
          const value = normalized.comments[id];
          const obj = isRecord(value) ? structuredClone(value) : { value: structuredClone(value) };
          if (isRecord(obj) && !("id" in obj)) obj.id = id;
          commentsArray.push([yMapFromJsonObject(obj)]);
        }
      } else {
        const commentsMap = doc.getMap("comments");
        for (const key of Array.from(commentsMap.keys())) commentsMap.delete(key);
        deleteLegacyListItemsFromMapRoot(transaction, commentsMap);
        for (const [id, value] of Object.entries(normalized.comments ?? {})) {
          const obj = isRecord(value) ? structuredClone(value) : { value: structuredClone(value) };
          if (isRecord(obj) && !("id" in obj)) obj.id = id;
          commentsMap.set(id, yMapFromJsonObject(obj));
        }
      }
    },
    opts.origin
  );
}

export {};
