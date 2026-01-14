import * as Y from "yjs";
import { cellKey } from "../diff/semanticDiff.js";
import { parseCellKey } from "../../../collab/session/src/cell-key.js";
import { getArrayRoot, getMapRoot, getYMap, yjsValueToJson } from "../../../collab/yjs-utils/src/index.ts";

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
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
 * Normalize sparse per-column row interval format runs into a `Map<col, runs[]>`.
 *
 * The collab binder stores range-run formatting on sheet metadata as:
 *   formatRunsByCol: { "0": [{ startRow, endRowExclusive, format }, ...], ... }
 *
 * This mirrors `DocumentController`'s internal `sheet.formatRunsByCol` data, but runs
 * store style objects (not style ids) because style ids are per-client.
 *
 * @param {any} raw
 * @returns {Map<number, Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>>}
 */
function parseFormatRunsByCol(raw) {
  /** @type {Map<number, Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>>} */
  const out = new Map();
  const json = yjsValueToJson(raw);
  if (!json) return out;

  /**
   * @param {any} colKey
   * @param {any} rawRuns
   */
  const addRunsForCol = (colKey, rawRuns) => {
    const col = Number(colKey);
    if (!Number.isInteger(col) || col < 0) return;

    const list = Array.isArray(rawRuns)
      ? rawRuns
      : isPlainObject(rawRuns)
        ? rawRuns.runs ?? rawRuns.formatRuns ?? rawRuns.segments ?? rawRuns.items ?? []
        : [];
    if (!Array.isArray(list) || list.length === 0) return;

    /** @type {Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>} */
    const runs = [];
    for (const entry of list) {
      if (!entry || typeof entry !== "object") continue;
      const startRow = Number(entry.startRow ?? entry.start?.row ?? entry.sr ?? entry.r0);
      const endRowExclusiveNum = Number(entry.endRowExclusive ?? entry.endRowExcl ?? entry.erx ?? entry.r1x);
      const endRowNum = Number(entry.endRow ?? entry.end?.row ?? entry.er ?? entry.r1);
      const endRowExclusive = Number.isInteger(endRowExclusiveNum)
        ? endRowExclusiveNum
        : Number.isInteger(endRowNum)
          ? endRowNum + 1
          : NaN;
      if (!Number.isInteger(startRow) || startRow < 0) continue;
      if (!Number.isInteger(endRowExclusive) || endRowExclusive <= startRow) continue;

      const format = extractStyleObject(yjsValueToJson(entry.format ?? entry.style ?? entry.value));
      if (!format) continue;
      runs.push({ startRow, endRowExclusive, format });
    }

    runs.sort((a, b) => a.startRow - b.startRow);
    if (runs.length > 0) out.set(col, runs);
  };

  // Preferred encoding: object keyed by column index.
  if (typeof json === "object" && !Array.isArray(json)) {
    for (const [key, value] of Object.entries(json)) {
      addRunsForCol(key, value);
    }
    return out;
  }

  // Also accept array encodings: [{ col, runs }, ...] or [[col, runs], ...].
  if (Array.isArray(json)) {
    for (const entry of json) {
      if (Array.isArray(entry)) {
        addRunsForCol(entry[0], entry[1]);
        continue;
      }
      if (entry && typeof entry === "object") {
        const col = entry.col ?? entry.index ?? entry.column;
        const runs = entry.runs ?? entry.formatRuns ?? entry.segments ?? entry.items;
        addRunsForCol(col, runs);
      }
    }
  }

  return out;
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
 * @param {any} sheetEntry
 * @param {string} key
 */
function readSheetEntryField(sheetEntry, key) {
  const map = getYMap(sheetEntry);
  if (map) return map.get(key);
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
  const map = getYMap(cellData);
  if (map) {
    const enc = map.get("enc");
    return {
      ...(enc !== null && enc !== undefined
        ? { enc: yjsValueToJson(enc), value: null, formula: null }
        : { value: map.get("value") ?? null, formula: map.get("formula") ?? null }),
      format: map.get("format") ?? map.get("style") ?? null,
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
 * Merge a single cell record into a sheet-local `cells` map using the same rules as
 * `sheetStateFromYjsDoc`.
 *
 * This is extracted so workbook snapshot extraction can group cells by sheet in a
 * single pass over the Yjs `cells` map while preserving the exact per-cell merge
 * behavior (encrypted precedence, canonical key preference, and format metadata
 * layering across duplicates).
 *
 * @param {Map<string, any>} cells
 * @param {{ sheetId: string, row: number, col: number }} parsed
 * @param {string} rawKey
 * @param {any} cellData
 */
export function mergeCellDataIntoSheetCells(cells, parsed, rawKey, cellData) {
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
}

/**
 * Extract layered sheet/row/col/range-run formatting metadata from a `sheets` array
 * entry.
 *
 * @param {any | null} sheetEntry
 */
export function sheetFormatLayersFromSheetEntry(sheetEntry) {
  const view = sheetEntry ? readSheetEntryField(sheetEntry, "view") : null;

  const rawDefaultFormat = sheetEntry != null ? readSheetEntryField(sheetEntry, "defaultFormat") : undefined;
  const rawRowFormats = sheetEntry != null ? readSheetEntryField(sheetEntry, "rowFormats") : undefined;
  const rawColFormats = sheetEntry != null ? readSheetEntryField(sheetEntry, "colFormats") : undefined;
  const rawFormatRunsByCol = sheetEntry != null ? readSheetEntryField(sheetEntry, "formatRunsByCol") : undefined;

  // Older snapshots stored formatting defaults nested under `sheet.view`. Avoid converting
  // the full `view` object to JSON (it can contain large unrelated payloads); instead, only
  // read the formatting keys we care about.
  const viewDefaultFormat = rawDefaultFormat !== undefined ? rawDefaultFormat : readSheetEntryField(view, "defaultFormat");
  const viewRowFormats = rawRowFormats !== undefined ? rawRowFormats : readSheetEntryField(view, "rowFormats");
  const viewColFormats = rawColFormats !== undefined ? rawColFormats : readSheetEntryField(view, "colFormats");
  const viewFormatRunsByCol =
    rawFormatRunsByCol !== undefined ? rawFormatRunsByCol : readSheetEntryField(view, "formatRunsByCol");

  const sheetDefaultFormat =
    extractStyleObject(yjsValueToJson(viewDefaultFormat)) ?? null;
  const rowFormats = parseIndexedFormats(viewRowFormats, "row");
  const colFormats = parseIndexedFormats(viewColFormats, "col");
  const formatRunsByCol = parseFormatRunsByCol(viewFormatRunsByCol);

  return { sheetDefaultFormat, rowFormats, colFormats, formatRunsByCol };
}

/**
 * @param {{ sheetDefaultFormat: any | null, rowFormats: Map<any, any>, colFormats: Map<any, any>, formatRunsByCol: Map<any, any> }} layers
 */
export function sheetHasLayeredFormats(layers) {
  return Boolean(
    layers.sheetDefaultFormat ||
      layers.rowFormats.size > 0 ||
      layers.colFormats.size > 0 ||
      layers.formatRunsByCol.size > 0,
  );
}

/**
 * Apply layered formatting defaults to all cells in a sheet.
 *
 * Matches `sheetStateFromYjsDoc` semantics exactly: only call this when the sheet
 * has any layered formats configured (`sheetHasLayeredFormats` is true).
 *
 * @param {Map<string, any>} cells
 * @param {{ sheetDefaultFormat: any | null, rowFormats: Map<number, Record<string, any>>, colFormats: Map<number, Record<string, any>>, formatRunsByCol: Map<number, Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>> }} layers
 */
export function applyLayeredFormatsToCells(cells, layers) {
  /**
   * Find the run containing `row` (half-open interval `[startRow, endRowExclusive)`).
   *
   * DocumentController guarantees these runs are sorted + non-overlapping, which lets us
   * do a binary search.
   *
   * @param {Array<{ startRow: number, endRowExclusive: number, format: Record<string, any> }>} runs
   * @param {number} row
   */
  const findRunForRow = (runs, row) => {
    let lo = 0;
    let hi = runs.length - 1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const run = runs[mid];
      if (row < run.startRow) {
        hi = mid - 1;
      } else if (row >= run.endRowExclusive) {
        lo = mid + 1;
      } else {
        return run;
      }
    }
    return null;
  };

  for (const [key, cell] of cells.entries()) {
    const m = String(key).match(/^r(\d+)c(\d+)$/);
    if (!m) continue;
    const row = Number(m[1]);
    const col = Number(m[2]);
    if (!Number.isInteger(row) || !Number.isInteger(col)) continue;

    const cellFormat = extractStyleObject(yjsValueToJson(cell?.format ?? cell?.style));
    let merged = deepMerge(
      deepMerge(layers.sheetDefaultFormat ?? {}, layers.colFormats.get(col) ?? null),
      layers.rowFormats.get(row) ?? null,
    );
    const runs = layers.formatRunsByCol.get(col);
    if (runs && runs.length > 0) {
      const run = findRunForRow(runs, row);
      if (run) merged = deepMerge(merged, run.format);
    }
    merged = deepMerge(merged, cellFormat);
    cell.format = normalizeFormat(merged);
  }
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
    mergeCellDataIntoSheetCells(cells, parsed, rawKey, cellData);
  });

  if (targetSheetId != null && cells.size > 0) {
    // Layered formatting defaults (Task 44): sheet/row/col formats can be stored on the
    // `sheets` metadata root. Only compute effective formats when the sheet has cells.
    const sheetEntry = findSheetEntryById(doc, targetSheetId);
    if (sheetEntry) {
      const formatLayers = sheetFormatLayersFromSheetEntry(sheetEntry);
      if (sheetHasLayeredFormats(formatLayers)) {
        applyLayeredFormatsToCells(cells, formatLayers);
      }
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
