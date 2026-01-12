import * as Y from "yjs";
import { cellKey } from "../diff/semanticDiff.js";
import { parseCellKey } from "../../../collab/session/src/cell-key.js";

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

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
function getArrayRoot(doc, name) {
  const existing = doc.share.get(name);
  if (isYArray(existing)) return existing;
  return doc.getArray(name);
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
 * @param {any} sheetEntry
 * @param {string} key
 */
function readSheetEntryField(sheetEntry, key) {
  if (isYMap(sheetEntry)) return sheetEntry.get(key);
  if (sheetEntry && typeof sheetEntry === "object") return sheetEntry[key];
  return undefined;
}

/**
 * @param {Y.Doc} doc
 * @param {string} sheetId
 * @returns {any | null}
 */
function findSheetEntryById(doc, sheetId) {
  if (!sheetId) return null;
  const sheets = getArrayRoot(doc, "sheets");
  // Deterministic choice: pick the last matching entry by index (mirrors binder behavior).
  let found = null;
  for (let i = 0; i < sheets.length; i++) {
    const entry = sheets.get(i);
    const id = yjsValueToJson(readSheetEntryField(entry, "id"));
    if (id === sheetId) found = entry;
  }
  return found;
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

  // Layered formatting defaults (Task 44): sheet/row/col formats can be stored on the
  // `sheets` metadata root. We only compute effective formats for the requested sheet.
  const sheetEntry = targetSheetId != null ? findSheetEntryById(doc, targetSheetId) : null;
  const view = sheetEntry ? readSheetEntryField(sheetEntry, "view") : null;
  const viewJson = view != null ? yjsValueToJson(view) : null;

  const rawDefaultFormat =
    sheetEntry != null ? readSheetEntryField(sheetEntry, "defaultFormat") : undefined;
  const rawRowFormats = sheetEntry != null ? readSheetEntryField(sheetEntry, "rowFormats") : undefined;
  const rawColFormats = sheetEntry != null ? readSheetEntryField(sheetEntry, "colFormats") : undefined;

  const sheetDefaultFormat =
    extractStyleObject(yjsValueToJson(rawDefaultFormat !== undefined ? rawDefaultFormat : viewJson?.defaultFormat)) ??
    null;
  const rowFormats = parseIndexedFormats(rawRowFormats !== undefined ? rawRowFormats : viewJson?.rowFormats, "row");
  const colFormats = parseIndexedFormats(rawColFormats !== undefined ? rawColFormats : viewJson?.colFormats, "col");

  /** @type {Map<string, any>} */
  const cells = new Map();
  cellsMap.forEach((cellData, rawKey) => {
    const parsed = parseSpreadsheetCellKey(rawKey);
    if (!parsed) return;
    if (targetSheetId != null && parsed.sheetId !== targetSheetId) return;

    const key = cellKey(parsed.row, parsed.col);
    const cell = extractCell(cellData);
    const existing = cells.get(key);

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
          cells.set(key, { ...cell, format: existing.format });
        } else {
          cells.set(key, cell);
        }
      } else if (existing.format == null && cell.format != null) {
        cells.set(key, { ...existing, format: cell.format });
      }
      return;
    }

    if (existingIsEncrypted) {
      // Preserve ciphertext, but allow plaintext duplicates to contribute format
      // metadata when the encrypted record lacks it.
      if (existing.format == null && cell.format != null) {
        cells.set(key, { ...existing, format: cell.format });
      }
      return;
    }

    cells.set(key, cell);
  });

  if (targetSheetId != null && (sheetDefaultFormat || rowFormats.size > 0 || colFormats.size > 0)) {
    for (const [key, cell] of cells.entries()) {
      const m = String(key).match(/^r(\d+)c(\d+)$/);
      if (!m) continue;
      const row = Number(m[1]);
      const col = Number(m[2]);
      if (!Number.isInteger(row) || !Number.isInteger(col)) continue;

      const cellFormat = extractStyleObject(yjsValueToJson(cell?.format ?? cell?.style));
      let merged = deepMerge(deepMerge(sheetDefaultFormat ?? {}, colFormats.get(col) ?? null), rowFormats.get(row) ?? null);
      merged = deepMerge(merged, cellFormat);
      cell.format = normalizeFormat(merged);
    }
  }

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
