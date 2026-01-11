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
    const placeholder = doc.share.get("comments");
    const hasStart = placeholder?._start != null;
    const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
    const kind = hasStart && mapSize === 0 ? "array" : "map";

    if (kind === "map") {
      const commentsMap = doc.getMap("comments");
      for (const id of Array.from(commentsMap.keys()).sort()) {
        comments[id] = yjsValueToJson(commentsMap.get(id));
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

  doc.transact(
    () => {
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
    const placeholder = doc.share.get("comments");
    const hasStart = placeholder?._start != null;
    const mapSize = placeholder?._map instanceof Map ? placeholder._map.size : 0;
    const kind = placeholder && hasStart && mapSize === 0 ? "array" : "map";

      if (kind === "array") {
        const commentsArray = doc.getArray("comments");
        if (commentsArray.length > 0) commentsArray.delete(0, commentsArray.length);
        const ids = Object.keys(normalized.comments ?? {}).sort();
      for (const id of ids) {
        const value = normalized.comments[id];
        const obj = isRecord(value) ? structuredClone(value) : { value: structuredClone(value) };
        if (isRecord(obj) && !("id" in obj)) obj.id = id;
        commentsArray.push([obj]);
      }
    } else {
      const commentsMap = doc.getMap("comments");
      for (const key of Array.from(commentsMap.keys())) commentsMap.delete(key);
        for (const [id, value] of Object.entries(normalized.comments ?? {})) {
          commentsMap.set(id, structuredClone(value));
        }
      }
    },
    opts.origin
  );
}

export {};
