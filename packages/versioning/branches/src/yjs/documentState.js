import * as Y from "yjs";
import { a1ToRowCol, rowColToA1 } from "./a1.js";

/**
 * @typedef {import("../types.js").Cell} Cell
 * @typedef {import("../types.js").DocumentState} DocumentState
 */

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
 * Supports:
 * - `${sheetId}:${row}:${col}`
 * - `${sheetId}:${row},${col}`
 *
 * @param {string} key
 * @returns {{ sheetId: string, row: number, col: number } | null}
 */
function parseCellKey(key) {
  const parts = key.split(":");
  if (parts.length === 3) {
    const sheetId = parts[0];
    const row = Number(parts[1]);
    const col = Number(parts[2]);
    if (!sheetId) return null;
    if (!Number.isInteger(row) || row < 0) return null;
    if (!Number.isInteger(col) || col < 0) return null;
    return { sheetId, row, col };
  }

  if (parts.length === 2) {
    const sheetId = parts[0];
    if (!sheetId) return null;
    const m = parts[1].match(/^(\d+),(\d+)$/);
    if (!m) return null;
    const row = Number(m[1]);
    const col = Number(m[2]);
    if (!Number.isInteger(row) || row < 0) return null;
    if (!Number.isInteger(col) || col < 0) return null;
    return { sheetId, row, col };
  }

  return null;
}

/**
 * @param {unknown} cellData
 * @returns {Cell | null}
 */
function extractCell(cellData) {
  const cellMap = getYMap(cellData);
  if (!cellMap) return null;

  const formulaRaw = cellMap.get("formula");
  const formula = typeof formulaRaw === "string" ? formulaRaw.trim() : "";
  const value = cellMap.get("value");

  // `format` is the canonical field (Task 18). `style` is accepted as a legacy alias.
  const format = cellMap.get("format") ?? cellMap.get("style") ?? null;

  /** @type {Cell} */
  const out = {};

  if (formula) out.formula = formula;
  else if (value !== null && value !== undefined) out.value = value;

  if (format !== null && format !== undefined) out.format = format;

  if (Object.keys(out).length === 0) return null;

  // Enforce mutual exclusion between formula/value.
  if (out.formula !== undefined) delete out.value;
  return out;
}

/**
 * Convert a Yjs spreadsheet document into a versioning DocumentState snapshot.
 *
 * @param {Y.Doc} ydoc
 * @returns {DocumentState}
 */
export function yjsDocToDocumentState(ydoc) {
  const cells = ydoc.getMap("cells");

  /** @type {DocumentState} */
  const state = { sheets: {} };

  cells.forEach((cellData, rawKey) => {
    if (typeof rawKey !== "string") return;
    const parsed = parseCellKey(rawKey);
    if (!parsed) return;

    const cell = extractCell(cellData);
    if (!cell) return;

    const addr = rowColToA1(parsed.row, parsed.col);
    const sheet = state.sheets[parsed.sheetId] ?? {};
    state.sheets[parsed.sheetId] = sheet;
    sheet[addr] = cell;
  });

  return state;
}

/**
 * @param {unknown} value
 * @returns {Cell | null}
 */
function normalizeDocumentCell(value) {
  if (!value || typeof value !== "object") return null;
  const cell = /** @type {any} */ (value);

  /** @type {Cell} */
  const out = {};

  const formulaRaw = cell.formula;
  const formula = typeof formulaRaw === "string" ? formulaRaw.trim() : "";
  if (formula) out.formula = formula;

  const v = cell.value;
  if (out.formula === undefined && v !== null && v !== undefined) out.value = v;

  const format = cell.format;
  if (format !== null && format !== undefined) out.format = format;

  if (Object.keys(out).length === 0) return null;
  return out;
}

/**
 * Apply a DocumentState snapshot into a Yjs spreadsheet document.
 *
 * This mutates the live workbook state (global checkout/merge semantics).
 *
 * @param {Y.Doc} ydoc
 * @param {DocumentState} state
 * @param {{ origin?: any }} [opts]
 */
export function applyDocumentStateToYjsDoc(ydoc, state, opts = {}) {
  const cells = ydoc.getMap("cells");

  /** @type {Map<string, Cell>} */
  const desired = new Map();

  for (const [sheetId, sheet] of Object.entries(state.sheets ?? {})) {
    if (!sheetId) continue;
    if (!sheet || typeof sheet !== "object") continue;
    for (const [a1, cellValue] of Object.entries(sheet)) {
      const normalized = normalizeDocumentCell(cellValue);
      if (!normalized) continue;
      const { row, col } = a1ToRowCol(a1);
      desired.set(`${sheetId}:${row}:${col}`, normalized);
    }
  }

  ydoc.transact(
    () => {
      // Delete cells that are no longer present in the snapshot.
      /** @type {string[]} */
      const toDelete = [];
      cells.forEach((_cellData, rawKey) => {
        if (typeof rawKey !== "string") return;
        const parsed = parseCellKey(rawKey);
        if (!parsed) return;
        if (!desired.has(rawKey)) toDelete.push(rawKey);
      });
      for (const key of toDelete) cells.delete(key);

      // Upsert desired cells.
      for (const [key, nextCell] of desired) {
        let cellMap = getYMap(cells.get(key));
        if (!cellMap) {
          cellMap = new Y.Map();
          cells.set(key, cellMap);
        }

        if (nextCell.formula !== undefined) {
          cellMap.set("formula", nextCell.formula);
          // CollabSession clears values for formulas; follow the same convention.
          cellMap.set("value", null);
        } else if (nextCell.value !== undefined) {
          cellMap.set("value", nextCell.value);
          cellMap.delete("formula");
        } else {
          cellMap.delete("value");
          cellMap.delete("formula");
        }

        if (nextCell.format !== undefined) {
          cellMap.set("format", nextCell.format);
          cellMap.delete("style");
        } else {
          cellMap.delete("format");
          cellMap.delete("style");
        }
      }
    },
    opts.origin
  );
}
