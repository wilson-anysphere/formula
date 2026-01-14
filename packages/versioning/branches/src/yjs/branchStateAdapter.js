import * as Y from "yjs";

import { normalizeCell } from "../cell.js";
import { normalizeDocumentState } from "../state.js";
import { a1ToRowCol, rowColToA1 } from "./a1.js";
import { getArrayRoot, getMapRoot, getYArray, getYMap, getYText, yjsValueToJson } from "../../../../collab/yjs-utils/src/index.ts";

/**
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../types.js").Cell} Cell
 * @typedef {import("../types.js").CellMap} CellMap
 * @typedef {import("../types.js").SheetMeta} SheetMeta
 */

/**
 * @typedef {{ Map: new () => any, Array: new () => any, Text: new () => any }} YjsTypeConstructors
 */

/**
 * @param {any} value
 * @returns {value is Record<string, any>}
 */
function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

/**
 * Return constructors for Y.Map/Y.Array/Y.Text that match the module instance used
 * to create `doc`.
 *
 * In pnpm workspaces it is possible to load both the ESM + CJS builds of Yjs in
 * the same process. Yjs types cannot be moved across module instances; the
 * safest approach is to construct nested types using constructors from the
 * target doc's module instance.
 *
 * @param {any} doc
 * @returns {YjsTypeConstructors}
 */
function getDocConstructors(doc) {
  const DocCtor = /** @type {any} */ (doc)?.constructor;
  if (typeof DocCtor !== "function") {
    return { Map: Y.Map, Array: Y.Array, Text: Y.Text };
  }

  try {
    const probe = new DocCtor();
    return {
      Map: probe.getMap("__ctor_probe_map").constructor,
      Array: probe.getArray("__ctor_probe_array").constructor,
      Text: probe.getText("__ctor_probe_text").constructor,
    };
  } catch {
    return { Map: Y.Map, Array: Y.Array, Text: Y.Text };
  }
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
  const text = getYText(value);
  if (text) return text.toString();
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
  const trimmed = String(value).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
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
  const sheetsArray = getArrayRoot(doc, "sheets");
  /** @type {Record<string, SheetMeta>} */
  const metaById = {};
  /** @type {string[]} */
  const order = [];

  for (const entry of sheetsArray.toArray()) {
    const id = coerceString(readYMapOrObject(entry, "id"));
    if (!id) continue;
    const name = coerceString(readYMapOrObject(entry, "name"));
    /** @type {SheetMeta} */
    const meta = { id, name };

    // Optional sheet metadata mirrored from the collab workbook schema.
    const rawVisibility = readYMapOrObject(entry, "visibility");
    const visibility = coerceString(rawVisibility);
    if (visibility === "visible" || visibility === "hidden" || visibility === "veryHidden") {
      meta.visibility = visibility;
    }

    const rawTabColor = readYMapOrObject(entry, "tabColor");
    if (rawTabColor === null) {
      meta.tabColor = null;
    } else {
      const tabColor = coerceString(rawTabColor);
      if (tabColor != null) meta.tabColor = tabColor;
    }

    // Per-sheet view state (e.g. frozen panes) can be stored either under a
    // dedicated `view` object or (legacy/experimental) as top-level fields.
    //
    // Layered formatting defaults (sheet/row/col) and range-run formatting are
    // stored on sheet metadata as top-level keys (`defaultFormat`, `rowFormats`,
    // `colFormats`, `formatRunsByCol`), but some BranchService-style snapshots may
    // still embed them inside `view`. Prefer the top-level keys when present.
    const rawView = readYMapOrObject(entry, "view");
    const rawDefaultFormat = readYMapOrObject(entry, "defaultFormat");
    const rawRowFormats = readYMapOrObject(entry, "rowFormats");
    const rawColFormats = readYMapOrObject(entry, "colFormats");
    const rawFormatRunsByCol = readYMapOrObject(entry, "formatRunsByCol");

    /** @type {any} */
    let view = rawView !== undefined ? yjsValueToJson(rawView) : null;

    if (view == null && rawView === undefined) {
      const frozenRows = readYMapOrObject(entry, "frozenRows");
      const frozenCols = readYMapOrObject(entry, "frozenCols");
      const colWidths = readYMapOrObject(entry, "colWidths");
      const rowHeights = readYMapOrObject(entry, "rowHeights");
      if (frozenRows !== undefined || frozenCols !== undefined || colWidths !== undefined || rowHeights !== undefined) {
        view = {
          frozenRows: yjsValueToJson(frozenRows) ?? 0,
          frozenCols: yjsValueToJson(frozenCols) ?? 0,
          ...(colWidths !== undefined ? { colWidths: yjsValueToJson(colWidths) } : {}),
          ...(rowHeights !== undefined ? { rowHeights: yjsValueToJson(rowHeights) } : {}),
        };
      }
    }

    const hasTopLevelFormats =
      rawDefaultFormat !== undefined ||
      rawRowFormats !== undefined ||
      rawColFormats !== undefined ||
      rawFormatRunsByCol !== undefined;
    if (hasTopLevelFormats) {
      if (!isRecord(view)) view = {};
      if (rawDefaultFormat !== undefined) view.defaultFormat = yjsValueToJson(rawDefaultFormat);
      if (rawRowFormats !== undefined) view.rowFormats = yjsValueToJson(rawRowFormats);
      if (rawColFormats !== undefined) view.colFormats = yjsValueToJson(rawColFormats);
      if (rawFormatRunsByCol !== undefined) view.formatRunsByCol = yjsValueToJson(rawFormatRunsByCol);
    }

    if (isRecord(view)) {
      meta.view = view;
    }

    metaById[id] = meta;
    order.push(id);
  }

  const cellsMap = getMapRoot(doc, "cells");

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
      const namedRangesMap = getMapRoot(doc, "namedRanges");
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
      const metadataMap = getMapRoot(doc, "metadata");
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
          const commentsMap = getMapRoot(doc, "comments");
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
          const commentsArray = getArrayRoot(doc, "comments");
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
  const docConstructors = getDocConstructors(doc);

  /**
   * @param {any} value
   */
  function yMapFromJsonObject(value) {
    const map = new docConstructors.Map();
    if (!isRecord(value)) return map;
    const keys = Object.keys(value).sort();
    for (const key of keys) {
      const v = value[key];
      if (key === "replies") {
        const replies = new docConstructors.Array();
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
    const text = getYText(value);
    if (text) {
      const out = new docConstructors.Text();
      out.applyDelta(structuredClone(text.toDelta()));
      return out;
    }
    const array = getYArray(value);
    if (array) {
      const out = new docConstructors.Array();
      for (const item of array.toArray()) out.push([cloneYjsValue(item)]);
      return out;
    }
    const map = getYMap(value);
    if (map) {
      const out = new docConstructors.Map();
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
      const sheetsArray = getArrayRoot(doc, "sheets");
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
          const entry = new docConstructors.Map();
          if (existing) {
            for (const key of Array.from(existing.keys()).sort()) {
              if (
                key === "id" ||
                key === "name" ||
                key === "view" ||
                key === "frozenRows" ||
                key === "frozenCols" ||
                key === "colWidths" ||
                key === "rowHeights"
              ) {
                continue;
              }
              entry.set(key, cloneYjsValue(existing.get(key)));
            }
          }
          entry.set("id", sheetId);
          entry.set("name", meta?.name ?? null);
          if (meta?.view !== undefined) entry.set("view", structuredClone(meta.view));
          // Normalize layered formatting + range-run metadata onto the top-level sheet entry.
          //
          // Collab schema prefers storing these fields as top-level keys so the binder and
          // versioning diffs can observe them without needing to parse `view`.
          //
          // BranchService `DocumentState` stores them inside `meta.view`, so extract them here.
          const view = meta?.view ?? null;
          if (view && typeof view === "object") {
            if (Object.prototype.hasOwnProperty.call(view, "defaultFormat")) {
              entry.set("defaultFormat", structuredClone(view.defaultFormat));
            } else {
              entry.delete("defaultFormat");
            }

            if (Object.prototype.hasOwnProperty.call(view, "rowFormats")) {
              const raw = view.rowFormats;
              if (raw && typeof raw === "object") {
                const map = new docConstructors.Map();
                for (const key of Object.keys(raw).sort()) {
                  map.set(key, structuredClone(raw[key]));
                }
                entry.set("rowFormats", map);
              } else {
                entry.delete("rowFormats");
              }
            } else {
              entry.delete("rowFormats");
            }

            if (Object.prototype.hasOwnProperty.call(view, "colFormats")) {
              const raw = view.colFormats;
              if (raw && typeof raw === "object") {
                const map = new docConstructors.Map();
                for (const key of Object.keys(raw).sort()) {
                  map.set(key, structuredClone(raw[key]));
                }
                entry.set("colFormats", map);
              } else {
                entry.delete("colFormats");
              }
            } else {
              entry.delete("colFormats");
            }

            if (Object.prototype.hasOwnProperty.call(view, "formatRunsByCol")) {
              const raw = view.formatRunsByCol;
              const map = new docConstructors.Map();
              if (Array.isArray(raw)) {
                for (const item of raw) {
                  const col = Number(item?.col);
                  if (!Number.isInteger(col) || col < 0) continue;
                  const runs = Array.isArray(item?.runs) ? item.runs : [];
                  map.set(String(col), structuredClone(runs));
                }
              } else if (raw && typeof raw === "object") {
                // Be forgiving: accept legacy object encodings keyed by column.
                for (const key of Object.keys(raw).sort()) {
                  map.set(key, structuredClone(raw[key]));
                }
              }
              entry.set("formatRunsByCol", map);
            } else {
              entry.delete("formatRunsByCol");
            }
          } else {
            entry.delete("defaultFormat");
            entry.delete("rowFormats");
            entry.delete("colFormats");
            entry.delete("formatRunsByCol");
          }
          if (meta && "visibility" in meta) {
            if (meta.visibility === "visible" || meta.visibility === "hidden" || meta.visibility === "veryHidden") {
              entry.set("visibility", meta.visibility);
            } else if (meta.visibility === null) {
              entry.delete("visibility");
            }
          }
          if (meta && "tabColor" in meta) {
            if (meta.tabColor === null) {
              entry.delete("tabColor");
            } else if (typeof meta.tabColor === "string") {
              entry.set("tabColor", meta.tabColor);
            }
          }
          desiredEntries.push(entry);
        }

        if (sheetsArray.length > 0) sheetsArray.delete(0, sheetsArray.length);
        sheetsArray.push(desiredEntries);
      } else if (sheetsArray.length === 0) {
        // Yjs workbooks are expected to have at least one sheet. If the branch
        // state is empty (legacy init), preserve app invariants by creating a
        // default sheet.
        const entry = new docConstructors.Map();
        entry.set("id", "Sheet1");
        entry.set("name", "Sheet1");
        entry.set("visibility", "visible");
        sheetsArray.push([entry]);
      }

      // --- Cells ---
      const cellsMap = getMapRoot(doc, "cells");
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
      const legacyKeysToDelete = [];
      /** @type {Set<string>} */
      const canonicalKeysToClear = new Set();

      // Drop legacy key encodings (`${sheetId}:${row},${col}` or `r{row}c{col}`),
      // but preserve *canonical* cell maps for removed cells by clearing them via
      // `value: null` / `formula: null` markers.
      //
      // Rationale for clearing (instead of deleting) removed canonical cells:
      // - Root `cells.delete(key)` operations do not create new Yjs Items, so
      //   later overwrites cannot reliably establish causal ordering against a
      //   delete (important for conflict monitors).
      // - Deep observers can miss root deletes depending on listener shape.
      //
      // We still delete legacy encodings to keep the document canonical and to
      // avoid permanently re-propagating duplicate keys after a checkout.
      cellsMap.forEach((_cellData, rawKey) => {
        if (typeof rawKey !== "string") return;
        const parsed = parseSpreadsheetCellKey(rawKey);
        if (!parsed) return;
        const canonical = `${parsed.sheetId}:${parsed.row}:${parsed.col}`;

        if (rawKey !== canonical) {
          legacyKeysToDelete.push(rawKey);
          if (!desiredCells.has(canonical)) canonicalKeysToClear.add(canonical);
          return;
        }

        if (!desiredCells.has(canonical)) canonicalKeysToClear.add(canonical);
      });

      for (const key of legacyKeysToDelete) {
        cellsMap.delete(key);
      }

      for (const canonicalKey of canonicalKeysToClear) {
        let yCell = getYMap(cellsMap.get(canonicalKey));
        if (!yCell) {
          yCell = new docConstructors.Map();
          cellsMap.set(canonicalKey, yCell);
        }

        yCell.delete("enc");
        yCell.set("value", null);
        yCell.set("formula", null);
        yCell.delete("format");
        yCell.delete("style");
      }

      for (const [key, normalizedCell] of desiredCells) {
        let yCell = getYMap(cellsMap.get(key));
        if (!yCell) {
          yCell = new docConstructors.Map();
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
            // When applying branch snapshots (global checkout/merge semantics),
            // represent formula clears as an explicit `null` marker instead of a
            // map delete. Map deletes don't create Yjs Items, which prevents
            // later overwrites from deterministically referencing the clear.
            yCell.set("formula", null);
          } else {
            // Format-only / empty cells: represent emptiness with explicit null
            // markers (instead of deleting keys) so other clients can causally
            // reference clears via Item.origin.
            yCell.set("value", null);
            yCell.set("formula", null);
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
      const namedRangesMap = getMapRoot(doc, "namedRanges");
      for (const key of Array.from(namedRangesMap.keys())) namedRangesMap.delete(key);
      for (const [key, value] of Object.entries(normalized.namedRanges ?? {})) {
        namedRangesMap.set(key, structuredClone(value));
      }

      // --- Metadata ---
      const metadataMap = getMapRoot(doc, "metadata");
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
        const commentsArray = getArrayRoot(doc, "comments");
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
        const commentsMap = getMapRoot(doc, "comments");
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
