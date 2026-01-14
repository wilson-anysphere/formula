import * as Y from "yjs";

import { normalizeCell } from "../cell.js";
import { normalizeDocumentState } from "../state.js";
import { a1ToRowCol, rowColToA1 } from "./a1.js";
import {
  cloneYjsValue,
  getArrayRoot,
  getDocTypeConstructors,
  getMapRoot,
  getYArray,
  getYMap,
  getYText,
  yjsValueToJson,
} from "../../../../collab/yjs-utils/src/index.ts";

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

function isPlainObject(value) {
  if (!value || typeof value !== "object") return false;
  if (Array.isArray(value)) return false;
  const proto = Object.getPrototypeOf(value);
  return proto === Object.prototype || proto === null;
}

// Drawing ids can be authored via remote/shared state (sheet view state). Keep validation strict
// so BranchService snapshot extraction doesn't accidentally materialize or deep-clone pathological
// ids (e.g. multi-megabyte Y.Text values) when producing version history / branch commits.
const MAX_DRAWING_ID_STRING_CHARS = 4096;

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
  if (text) return yjsValueToJson(text);
  if (typeof value === "string") return value;
  if (value == null) return null;
  return String(value);
}

/**
 * Normalize a `drawing.id` payload into a safe string/number identifier.
 *
 * @param {unknown} value
 * @returns {string | number | null}
 */
function normalizeDrawingIdValue(value) {
  const text = getYText(value);
  if (text) {
    // Avoid `text.toString()` for oversized ids: it would allocate a large JS string.
    if (typeof text.length === "number" && text.length > MAX_DRAWING_ID_STRING_CHARS) return null;
    value = yjsValueToJson(text);
  }

  if (typeof value === "string") {
    if (value.length > MAX_DRAWING_ID_STRING_CHARS) return null;
    const trimmed = value.trim();
    if (!trimmed) return null;
    return trimmed;
  }

  if (typeof value === "number") {
    if (!Number.isSafeInteger(value)) return null;
    return value;
  }

  return null;
}

/**
 * Convert a `drawings` list into JSON without materializing oversized `drawing.id` strings.
 *
 * @param {unknown} raw
 * @returns {any[] | null}
 */
function drawingsValueToJsonSafe(raw) {
  if (raw === null) return null;
  if (raw === undefined) return null;

  const yArr = getYArray(raw);
  const isArr = Array.isArray(raw);
  if (!yArr && !isArr) return null;

  /** @type {any[]} */
  const out = [];
  const len = yArr ? yArr.length : raw.length;

  for (let idx = 0; idx < len; idx += 1) {
    const entry = yArr ? yArr.get(idx) : raw[idx];

    const map = getYMap(entry);
    if (map) {
      const normalizedId = normalizeDrawingIdValue(map.get("id"));
      if (normalizedId == null) continue;

      /** @type {any} */
      const obj = { id: normalizedId };
      const keys = Array.from(map.keys()).sort();
      for (const key of keys) {
        if (key === "id") continue;
        obj[String(key)] = yjsValueToJson(map.get(key));
      }
      out.push(obj);
      continue;
    }

    if (isPlainObject(entry)) {
      const normalizedId = normalizeDrawingIdValue(entry.id);
      if (normalizedId == null) continue;

      /** @type {any} */
      const obj = { id: normalizedId };
      const keys = Object.keys(entry).sort();
      for (const key of keys) {
        if (key === "id") continue;
        obj[key] = yjsValueToJson(entry[key]);
      }
      out.push(obj);
    }
  }

  return out;
}

/**
 * Convert a sheet `view` object into JSON, treating `view.drawings` specially so we don't
 * materialize oversized `drawing.id` strings.
 *
 * @param {unknown} rawView
 * @returns {any}
 */
function sheetViewValueToJsonSafe(rawView) {
  if (rawView == null) return yjsValueToJson(rawView);

  const map = getYMap(rawView);
  if (map) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Array.from(map.keys()).sort();
    for (const key of keys) {
      if (key === "drawings") {
        const rawDrawings = map.get(key);
        if (rawDrawings === null) out.drawings = null;
        else out.drawings = drawingsValueToJsonSafe(rawDrawings) ?? [];
        continue;
      }
      out[String(key)] = yjsValueToJson(map.get(key));
    }
    return out;
  }

  if (isPlainObject(rawView)) {
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Object.keys(rawView).sort();
    for (const key of keys) {
      if (key === "drawings") {
        const rawDrawings = rawView.drawings;
        if (rawDrawings === null) out.drawings = null;
        else out.drawings = drawingsValueToJsonSafe(rawDrawings) ?? [];
        continue;
      }
      out[key] = yjsValueToJson(rawView[key]);
    }
    return out;
  }

  // Unknown/invalid view type. Avoid materializing it (e.g. huge Y.Text); treat as absent.
  return null;
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
  const json = yjsValueToJson(value);
  if (json == null) return null;
  const trimmed = String(json).trim();
  const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
  const stripped = strippedLeading.trim();
  if (stripped === "") return null;
  return `=${stripped}`;
}

/**
 * Parse a run of ASCII digits into a number.
 *
 * Leading zeros are allowed (this is used for legacy cell key formats).
 *
 * @param {string} value
 * @param {number} start
 * @param {number} end
 * @returns {number | null}
 */
function parseUnsignedInt(value, start, end) {
  if (end <= start) return null;
  let out = 0;
  for (let i = start; i < end; i++) {
    const code = value.charCodeAt(i);
    if (code < 48 || code > 57) return null;
    out = out * 10 + (code - 48);
  }
  return out;
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
 * @param {{ sheetId: string, row: number, col: number, isCanonical: boolean }} [out]
 * @returns {{ sheetId: string, row: number, col: number, isCanonical: boolean } | null}
 */
function parseSpreadsheetCellKey(key, out) {
  if (typeof key !== "string" || key.length === 0) return null;

  const firstColon = key.indexOf(":");
  if (firstColon !== -1) {
    const sheetId = key.slice(0, firstColon);
    if (!sheetId) return null;

    const secondColon = key.indexOf(":", firstColon + 1);
    if (secondColon !== -1) {
      // Reject 3+ colon encodings (unsupported).
      if (key.indexOf(":", secondColon + 1) !== -1) return null;

      const rowStart = firstColon + 1;
      const rowEnd = secondColon;
      const colStart = secondColon + 1;
      const colEnd = key.length;

      // Fast path: digit-only row/col segments are the common case.
      const rowDigits = parseUnsignedInt(key, rowStart, rowEnd);
      const colDigits = parseUnsignedInt(key, colStart, colEnd);
      if (rowDigits != null && colDigits != null) {
        if (!Number.isInteger(rowDigits) || rowDigits < 0) return null;
        if (!Number.isInteger(colDigits) || colDigits < 0) return null;
        const isCanonical =
          (rowEnd - rowStart === 1 || key.charCodeAt(rowStart) !== 48) &&
          (colEnd - colStart === 1 || key.charCodeAt(colStart) !== 48);
        if (out) {
          out.sheetId = sheetId;
          out.row = rowDigits;
          out.col = colDigits;
          out.isCanonical = isCanonical;
          return out;
        }
        return { sheetId, row: rowDigits, col: colDigits, isCanonical };
      }

      // Fallback: preserve legacy acceptance semantics (e.g. `1e0`, whitespace).
      const rowStr = key.slice(rowStart, rowEnd);
      const colStr = key.slice(colStart, colEnd);
      const row = Number(rowStr);
      const col = Number(colStr);
      if (!Number.isInteger(row) || row < 0) return null;
      if (!Number.isInteger(col) || col < 0) return null;
      if (out) {
        out.sheetId = sheetId;
        out.row = row;
        out.col = col;
        out.isCanonical = false;
        return out;
      }
      return { sheetId, row, col, isCanonical: false };
    }

    // Legacy `${sheetId}:${row},${col}` encoding.
    const comma = key.indexOf(",", firstColon + 1);
    if (comma === -1) return null;
    const row = parseUnsignedInt(key, firstColon + 1, comma);
    if (row == null) return null;
    const col = parseUnsignedInt(key, comma + 1, key.length);
    if (col == null) return null;
    if (out) {
      out.sheetId = sheetId;
      out.row = row;
      out.col = col;
      out.isCanonical = false;
      return out;
    }
    return { sheetId, row, col, isCanonical: false };
  }

  // Unit-test convenience `r{row}c{col}` encoding (assumed to be in Sheet1).
  if (key.charCodeAt(0) === 114) {
    const cIdx = key.indexOf("c", 1);
    if (cIdx !== -1) {
      const row = parseUnsignedInt(key, 1, cIdx);
      const col = parseUnsignedInt(key, cIdx + 1, key.length);
      if (row != null && col != null) {
        if (out) {
          out.sheetId = "Sheet1";
          out.row = row;
          out.col = col;
          out.isCanonical = false;
          return out;
        }
        return { sheetId: "Sheet1", row, col, isCanonical: false };
      }
    }
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

  for (let i = 0; i < sheetsArray.length; i++) {
    const entry = sheetsArray.get(i);
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
    let view = rawView !== undefined ? sheetViewValueToJsonSafe(rawView) : null;

    if (view == null && rawView === undefined) {
      const frozenRows = readYMapOrObject(entry, "frozenRows");
      const frozenCols = readYMapOrObject(entry, "frozenCols");
      const backgroundImageId = readYMapOrObject(entry, "backgroundImageId") ?? readYMapOrObject(entry, "background_image_id");
      const colWidths = readYMapOrObject(entry, "colWidths");
      const rowHeights = readYMapOrObject(entry, "rowHeights");
      const mergedRanges =
        readYMapOrObject(entry, "mergedRanges") ??
        readYMapOrObject(entry, "mergedCells") ??
        readYMapOrObject(entry, "merged_cells");
      const drawings = readYMapOrObject(entry, "drawings");
      if (
        frozenRows !== undefined ||
        frozenCols !== undefined ||
        backgroundImageId !== undefined ||
        colWidths !== undefined ||
        rowHeights !== undefined ||
        mergedRanges !== undefined ||
        drawings !== undefined
      ) {
        view = {
          frozenRows: yjsValueToJson(frozenRows) ?? 0,
          frozenCols: yjsValueToJson(frozenCols) ?? 0,
          // Always include the key (even when null) so BranchService can distinguish
          // explicit clears from omissions during semantic merges.
           backgroundImageId: yjsValueToJson(backgroundImageId) ?? null,
           ...(colWidths !== undefined ? { colWidths: yjsValueToJson(colWidths) } : {}),
           ...(rowHeights !== undefined ? { rowHeights: yjsValueToJson(rowHeights) } : {}),
           ...(mergedRanges !== undefined ? { mergedRanges: yjsValueToJson(mergedRanges) } : {}),
           ...(drawings !== undefined
             ? { drawings: drawings === null ? null : drawingsValueToJsonSafe(drawings) ?? [] }
             : {}),
          };
        }
      }

    // Ensure BranchService snapshots can represent explicit clears for `backgroundImageId`.
    //
    // Collaboration schema typically deletes the key on clear, but BranchService treats missing
    // keys as "no change" to support older clients. Include `backgroundImageId: null` so clears
    // survive commits + semantic merges.
    if (view == null) {
      view = { frozenRows: 0, frozenCols: 0, backgroundImageId: null, mergedRanges: [], drawings: [] };
    } else if (isRecord(view)) {
      const hasKey =
        Object.prototype.hasOwnProperty.call(view, "backgroundImageId") ||
        Object.prototype.hasOwnProperty.call(view, "background_image_id");
      if (!hasKey) {
        const topLevelBackgroundImageId =
          readYMapOrObject(entry, "backgroundImageId") ?? readYMapOrObject(entry, "background_image_id");
        if (topLevelBackgroundImageId !== undefined) {
          view.backgroundImageId = yjsValueToJson(topLevelBackgroundImageId) ?? null;
        } else {
          view.backgroundImageId = null;
        }
      }

      const hasMergedRanges =
        Object.prototype.hasOwnProperty.call(view, "mergedRanges") ||
        Object.prototype.hasOwnProperty.call(view, "mergedCells") ||
        Object.prototype.hasOwnProperty.call(view, "merged_cells");
      if (!hasMergedRanges) {
        const topLevelMergedRanges =
          readYMapOrObject(entry, "mergedRanges") ??
          readYMapOrObject(entry, "mergedCells") ??
          readYMapOrObject(entry, "merged_cells");
        if (topLevelMergedRanges !== undefined) {
          view.mergedRanges = yjsValueToJson(topLevelMergedRanges) ?? [];
        } else {
          view.mergedRanges = [];
        }
      } else if (!Object.prototype.hasOwnProperty.call(view, "mergedRanges")) {
        view.mergedRanges = view.mergedCells ?? view.merged_cells ?? [];
      }

      if (!Object.prototype.hasOwnProperty.call(view, "drawings")) {
        const topLevelDrawings = readYMapOrObject(entry, "drawings");
        if (topLevelDrawings !== undefined) {
          view.drawings = drawingsValueToJsonSafe(topLevelDrawings) ?? [];
        } else {
          view.drawings = [];
        }
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

  /** @type {{ sheetId: string, row: number, col: number, isCanonical: boolean }} */
  const parsedScratch = { sheetId: "", row: 0, col: 0, isCanonical: false };
  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey, parsedScratch);
    if (!parsed) return;
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
    const isCanonical = parsed.isCanonical === true;

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
        for (let i = 0; i < existingArray.length; i++) {
          const item = existingArray.get(i);
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
          for (let i = 0; i < commentsArray.length; i++) {
            const item = commentsArray.get(i);
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
  const docConstructors = getDocTypeConstructors(doc);

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

  doc.transact(
    (transaction) => {
      // --- Sheets ---
      const sheetsArray = getArrayRoot(doc, "sheets");
      if (normalized.sheets.order.length > 0) {
        /** @type {Map<string, Y.Map<any>>} */
        const existingById = new Map();
        for (let i = 0; i < sheetsArray.length; i++) {
          const entry = sheetsArray.get(i);
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
               key === "backgroundImageId" ||
               key === "background_image_id" ||
               key === "colWidths" ||
                key === "rowHeights" ||
                key === "mergedRanges" ||
                key === "mergedCells" ||
                key === "merged_cells" ||
                key === "drawings"
              ) {
                continue;
              }
              entry.set(key, cloneYjsValue(existing.get(key), docConstructors));
            }
          }
          entry.set("id", sheetId);
          entry.set("name", meta?.name ?? null);
          if (meta?.view !== undefined) {
            const nextView = structuredClone(meta.view);
            if (isRecord(nextView)) {
              // Keep the Yjs document canonical and lightweight: omit empty list fields that the
              // collaboration schema commonly represents by deleting the key.
              if (Array.isArray(nextView.mergedRanges) && nextView.mergedRanges.length === 0) delete nextView.mergedRanges;
              if (Array.isArray(nextView.drawings) && nextView.drawings.length === 0) delete nextView.drawings;
            }
            entry.set("view", nextView);
          }
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
      /** @type {{ sheetId: string, row: number, col: number, isCanonical: boolean }} */
      const parsedScratch = { sheetId: "", row: 0, col: 0, isCanonical: false };
      cellsMap.forEach((_cellData, rawKey) => {
        if (typeof rawKey !== "string") return;
        const parsed = parseSpreadsheetCellKey(rawKey, parsedScratch);
        if (!parsed) return;
        const canonical = parsed.isCanonical === true ? rawKey : `${parsed.sheetId}:${parsed.row}:${parsed.col}`;
        if (parsed.isCanonical !== true) legacyKeysToDelete.push(rawKey);
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
