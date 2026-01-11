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
 * @param {unknown} value
 * @returns {Y.Map<any> | null}
 */
function getYMap(value) {
  if (value instanceof Y.Map) return value;

  // Duck-type to handle multiple `yjs` module instances.
  if (!value || typeof value !== "object") return null;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YMap") return null;
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  return /** @type {Y.Map<any>} */ (maybe);
}

/**
 * @param {unknown} value
 * @returns {Y.Array<any> | null}
 */
function getYArray(value) {
  if (value instanceof Y.Array) return value;

  // Duck-type to handle multiple `yjs` module instances.
  if (!value || typeof value !== "object") return null;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YArray") return null;
  if (typeof maybe.toArray !== "function") return null;
  if (typeof maybe.push !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  return /** @type {Y.Array<any>} */ (maybe);
}

/**
 * @param {unknown} value
 * @returns {value is Y.Text}
 */
function isYText(value) {
  if (value instanceof Y.Text) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YText") return false;
  if (typeof maybe.toString !== "function") return false;
  return true;
}

/**
 * @param {any} value
 * @param {string} key
 */
function readYMapOrObject(value, key) {
  const map = getYMap(value);
  if (map) return map.get(key);
  if (isRecord(value)) return value[key];
  return undefined;
}

/**
 * @param {any} value
 * @returns {string | null}
 */
function coerceString(value) {
  if (isYText(value)) return value.toString();
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
  if (isYText(value)) return value.toString();

  const array = getYArray(value);
  if (array) return array.toArray().map((v) => yjsValueToJson(v));

  const map = getYMap(value);
  if (map) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Array.from(map.keys()).sort();
    for (const key of keys) out[key] = yjsValueToJson(map.get(key));
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
        const map = getYMap(value);
        if (map) out.push(map);
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
  const cellMap = getYMap(cellData);
  if (cellMap) {
    const enc = cellMap.get("enc");
    const formula = normalizeFormula(cellMap.get("formula"));
    const value = cellMap.get("value");
    const format = cellMap.get("format") ?? cellMap.get("style");
    if (enc !== null && enc !== undefined) {
      cell.enc = yjsValueToJson(enc);
    } else if (formula) {
      cell.formula = formula;
    } else if (value !== null && value !== undefined) {
      cell.value = value;
    }
    if (format !== null && format !== undefined) cell.format = yjsValueToJson(format);
    return normalizeCell(cell);
  }

  if (isRecord(cellData)) {
    const enc = cellData.enc;
    const formula = normalizeFormula(cellData.formula);
    const value = cellData.value;
    const format = cellData.format ?? cellData.style;
    if (enc !== null && enc !== undefined) {
      cell.enc = yjsValueToJson(enc);
    } else if (formula) {
      cell.formula = formula;
    } else if (value !== null && value !== undefined) {
      cell.value = value;
    }
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
    if (!cell) return;

    const existing = cells[sheetId][addr];

    // If any representation of this cell is encrypted, treat it as encrypted and
    // do not allow plaintext duplicates (e.g. from legacy key encodings) to
    // overwrite the ciphertext in branch snapshots.
    const canonicalKey = `${sheetId}:${parsed.row}:${parsed.col}`;
    const isCanonical = rawKey === canonicalKey;

    if (cell.enc != null) {
      if (!existing || existing.enc == null || isCanonical) {
        cells[sheetId][addr] = cell;
      } else if (existing.enc != null && existing.format == null && cell.format != null) {
        cells[sheetId][addr] = { ...existing, format: cell.format };
      }
      return;
    }

    if (existing && existing.enc != null) {
      if (existing.format == null && cell.format != null) {
        cells[sheetId][addr] = { ...existing, format: cell.format };
      }
      return;
    }

    cells[sheetId][addr] = cell;
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
  const metadata = {};
  if (doc.share.has("metadata")) {
    try {
      const metadataMap = doc.getMap("metadata");
      for (const key of Array.from(metadataMap.keys()).sort()) {
        metadata[key] = yjsValueToJson(metadataMap.get(key));
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

    const existingMap = getYMap(existing);
    if (existingMap) {
      /** @type {Map<string, any>} */
      const byId = new Map();
      for (const id of Array.from(existingMap.keys()).sort()) {
        byId.set(id, yjsValueToJson(existingMap.get(id)));
      }
      // Recovery: legacy list items stored on a Map root.
      for (const item of legacyListCommentsFromMapRoot(existingMap)) {
        const id = coerceString(readYMapOrObject(item, "id"));
        if (!id) continue;
        if (byId.has(id)) continue;
        byId.set(id, yjsValueToJson(item));
      }
      for (const id of Array.from(byId.keys()).sort()) {
        comments[id] = byId.get(id);
      }
    } else {
      const existingArray = getYArray(existing);
      if (existingArray) {
        /** @type {Map<string, any>} */
        const byId = new Map();
        // Recovery: map entries stored on an Array root (mixed schema).
        for (const [id, value] of mapEntriesFromArrayRoot(existingArray)) {
          byId.set(id, yjsValueToJson(value));
        }
        for (const item of existingArray.toArray()) {
          const id = coerceString(readYMapOrObject(item, "id"));
          if (!id) continue;
          if (byId.has(id)) continue;
          byId.set(id, yjsValueToJson(item));
        }
        for (const id of Array.from(byId.keys()).sort()) {
          comments[id] = byId.get(id);
        }
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
          /** @type {Map<string, any>} */
          const byId = new Map();
          for (const [id, value] of mapEntriesFromArrayRoot(commentsArray)) {
            byId.set(id, yjsValueToJson(value));
          }
          for (const item of commentsArray.toArray()) {
            const id = coerceString(readYMapOrObject(item, "id"));
            if (!id) continue;
            if (byId.has(id)) continue;
            byId.set(id, yjsValueToJson(item));
          }
          for (const id of Array.from(byId.keys()).sort()) {
            comments[id] = byId.get(id);
          }
        }
      }
    }
  }

  return normalizeDocumentState({
    schemaVersion: 1,
    sheets: { order, metaById },
    cells,
    metadata,
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

  /**
   * Deep clone a Yjs value so it can be re-inserted into the document without
   * re-integrating an already-attached type (which can throw inside Yjs).
   *
   * Used to preserve unknown metadata when applying a snapshot.
   *
   * @param {any} value
   * @returns {any}
   */
  function cloneYjsValue(value) {
    if (isYText(value)) {
      const out = new Y.Text();
      out.applyDelta(structuredClone(value.toDelta()));
      return out;
    }
    const array = getYArray(value);
    if (array) {
      const out = new Y.Array();
      for (const item of array.toArray()) out.push([cloneYjsValue(item)]);
      return out;
    }
    const map = getYMap(value);
    if (map) {
      const out = new Y.Map();
      for (const key of Array.from(map.keys()).sort()) {
        out.set(key, cloneYjsValue(map.get(key)));
      }
      return out;
    }
    if (Array.isArray(value)) return value.map((v) => cloneYjsValue(v));
    if (isRecord(value)) return structuredClone(value);
    if (value && typeof value === "object") return structuredClone(value);
    return value;
  }

  doc.transact(
    (transaction) => {
      // --- Sheets ---
      const sheetsArray = doc.getArray("sheets");
      if (normalized.sheets.order.length > 0) {
        /** @type {Map<string, Y.Map<any>>} */
        const existingById = new Map();
        for (const entry of sheetsArray.toArray()) {
          const id = coerceString(readYMapOrObject(entry, "id"));
          if (!id) continue;
          const map = getYMap(entry);
          if (map && !existingById.has(id)) existingById.set(id, map);
        }

        /** @type {Y.Map<any>[]} */
        const desiredEntries = [];
        for (const sheetId of normalized.sheets.order) {
          const meta = normalized.sheets.metaById[sheetId];
          const existing = existingById.get(sheetId);
          const entry = new Y.Map();
          if (existing) {
            for (const key of Array.from(existing.keys()).sort()) {
              if (key === "id" || key === "name") continue;
              entry.set(key, cloneYjsValue(existing.get(key)));
            }
          }
          entry.set("id", sheetId);
          entry.set("name", meta?.name ?? null);
          desiredEntries.push(entry);
        }

        if (sheetsArray.length > 0) sheetsArray.delete(0, sheetsArray.length);
        sheetsArray.push(desiredEntries);
      } else if (sheetsArray.length === 0) {
        // Yjs workbooks are expected to have at least one sheet. If the branch
        // state is empty (legacy init), preserve app invariants by creating a
        // default sheet.
        const entry = new Y.Map();
        entry.set("id", "Sheet1");
        entry.set("name", "Sheet1");
        sheetsArray.push([entry]);
      }

      // --- Cells ---
      const cellsMap = doc.getMap("cells");
      /** @type {Map<string, Cell>} */
      const desiredCells = new Map();

      for (const [sheetId, sheet] of Object.entries(normalized.cells)) {
        for (const [addr, cell] of Object.entries(sheet ?? {})) {
          const normalizedCell = normalizeCell(cell);
          if (!normalizedCell) continue;
          const { row, col } = a1ToRowCol(addr);
          desiredCells.set(`${sheetId}:${row}:${col}`, normalizedCell);
        }
      }

      /** @type {string[]} */
      const toDelete = [];
      cellsMap.forEach((_cellData, rawKey) => {
        if (typeof rawKey !== "string") return;
        const parsed = parseSpreadsheetCellKey(rawKey);
        if (!parsed) return;
        const canonical = `${parsed.sheetId}:${parsed.row}:${parsed.col}`;
        if (!desiredCells.has(canonical) || rawKey !== canonical) {
          toDelete.push(rawKey);
        }
      });
      for (const key of toDelete) cellsMap.delete(key);

      for (const [key, normalizedCell] of desiredCells) {
        let yCell = getYMap(cellsMap.get(key));
        if (!yCell) {
          yCell = new Y.Map();
          cellsMap.set(key, yCell);
        }

        if (normalizedCell.enc !== undefined && normalizedCell.enc !== null) {
          // Preserve ciphertext exactly; branch snapshots treat it as opaque.
          yCell.set("enc", structuredClone(normalizedCell.enc));
          yCell.delete("value");
          yCell.delete("formula");
        } else {
          yCell.delete("enc");

          const formula = normalizeFormula(normalizedCell.formula);
          if (formula) {
            yCell.set("formula", formula);
            // CollabSession clears values for formulas; follow the same convention.
            yCell.set("value", null);
          } else if (normalizedCell.value !== undefined) {
            yCell.set("value", normalizedCell.value);
            yCell.delete("formula");
          } else {
            yCell.delete("value");
            yCell.delete("formula");
          }
        }

        if (normalizedCell.format != null) {
          yCell.set("format", structuredClone(normalizedCell.format));
          yCell.delete("style");
        } else {
          yCell.delete("format");
          yCell.delete("style");
        }
      }

      // --- Named ranges ---
      const namedRangesMap = doc.getMap("namedRanges");
      for (const key of Array.from(namedRangesMap.keys())) namedRangesMap.delete(key);
      for (const [key, value] of Object.entries(normalized.namedRanges ?? {})) {
        namedRangesMap.set(key, structuredClone(value));
      }

      // --- Metadata ---
      const metadataMap = doc.getMap("metadata");
      for (const key of Array.from(metadataMap.keys())) metadataMap.delete(key);
      for (const [key, value] of Object.entries(normalized.metadata ?? {})) {
        metadataMap.set(key, structuredClone(value));
      }

      // --- Comments ---
      const existing = doc.share.get("comments");
      const commentsKind =
        getYArray(existing)
          ? "array"
          : getYMap(existing)
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
