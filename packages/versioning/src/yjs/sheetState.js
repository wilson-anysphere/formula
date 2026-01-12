import * as Y from "yjs";
import { cellKey } from "../diff/semanticDiff.js";
import { parseCellKey } from "../../../collab/session/src/cell-key.js";

function isYMap(value) {
  if (value instanceof Y.Map) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YMap") return false;
  return (
    typeof maybe.get === "function" &&
    typeof maybe.set === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.forEach === "function"
  );
}

function isYArray(value) {
  if (value instanceof Y.Array) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YArray") return false;
  return typeof maybe.get === "function" && typeof maybe.toArray === "function";
}

function isYText(value) {
  if (value instanceof Y.Text) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YText") return false;
  return typeof maybe.toString === "function";
}

/**
 * Convert a Yjs value (potentially nested) into a plain JS value.
 *
 * This is primarily used for encrypted cell payloads where we need deterministic,
 * deep-equality-friendly data but must not attempt to decrypt.
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
    const keys = [];
    value.forEach((_v, k) => keys.push(String(k)));
    keys.sort();
    for (const k of keys) out[k] = yjsValueToJson(value.get(k));
    return out;
  }

  if (Array.isArray(value)) return value.map((v) => yjsValueToJson(v));

  if (value && typeof value === "object") {
    const proto = Object.getPrototypeOf(value);
    // Only canonicalize plain objects; preserve prototypes for non-plain types.
    if (proto !== Object.prototype && proto !== null) {
      return structuredClone(value);
    }
    /** @type {Record<string, any>} */
    const out = {};
    const keys = Object.keys(value).sort();
    for (const key of keys) out[key] = yjsValueToJson(value[key]);
    return out;
  }

  return value;
}

/**
 * @param {Y.Doc} doc
 * @param {string} name
 */
function getMapRoot(doc, name) {
  const existing = doc.share.get(name);
  if (isYMap(existing)) return existing;
  // Placeholder / missing roots are safe to instantiate via Yjs' constructors.
  return doc.getMap(name);
}

/**
 * Parse a spreadsheet cell key. Supports:
 * - `${sheetId}:${row}:${col}` (docs/06-collaboration.md)
 * - `${sheetId}:${row},${col}` (legacy internal encoding)
 * - `r{row}c{col}` (unit-test convenience, resolved against `defaultSheetId`)
 *
 * @param {string} key
 * @param {{ defaultSheetId?: string }} [opts]
 * @returns {{ sheetId: string, row: number, col: number } | null}
 */
export function parseSpreadsheetCellKey(key, opts = {}) {
  const parsed = parseCellKey(key, { defaultSheetId: opts.defaultSheetId ?? "Sheet1" });
  if (!parsed) return null;
  return { sheetId: parsed.sheetId, row: parsed.row, col: parsed.col };
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
 * Convert a Yjs doc into a per-sheet state suitable for semantic diff.
 *
 * @param {Y.Doc} doc
 * @param {{ sheetId?: string | null }} [opts]
 */
export function sheetStateFromYjsDoc(doc, opts = {}) {
  const targetSheetId = opts.sheetId ?? null;
  const cellsMap = getMapRoot(doc, "cells");

  /** @type {Map<string, any>} */
  const cells = new Map();
  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed) return;
    if (targetSheetId != null && parsed.sheetId !== targetSheetId) return;
    cells.set(cellKey(parsed.row, parsed.col), extractCell(cellData));
  });

  return { cells };
}

/**
 * @param {Uint8Array} snapshot
 * @param {{ sheetId?: string | null }} [opts]
 */
export function sheetStateFromYjsSnapshot(snapshot, opts = {}) {
  const doc = new Y.Doc();
  Y.applyUpdate(doc, snapshot);
  return sheetStateFromYjsDoc(doc, opts);
}
