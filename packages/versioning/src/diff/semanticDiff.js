import { isDeepStrictEqual } from "node:util";
import { normalizeFormula } from "../formula/normalize.js";

/**
 * @typedef {{ row: number, col: number }} CellRef
 * @typedef {{ value?: any, formula?: string | null, format?: any, enc?: any }} Cell
 *
 * @typedef {{ cells: Map<string, Cell> }} SheetState
 *
 * @typedef {{ cell: CellRef, oldValue?: any, newValue?: any, oldFormula?: string | null, newFormula?: string | null }} CellChange
 * @typedef {{ oldLocation: CellRef, newLocation: CellRef, value: any, formula?: string | null }} MoveChange
 *
 * @typedef {{
 *   added: CellChange[],
 *   removed: CellChange[],
 *   modified: CellChange[],
 *   moved: MoveChange[],
 *   formatOnly: CellChange[],
 * }} DiffResult
 */

/**
 * @param {number} row
 * @param {number} col
 */
export function cellKey(row, col) {
  return `r${row}c${col}`;
}

/**
 * @param {string} key
 * @returns {CellRef}
 */
export function parseCellKey(key) {
  const m = key.match(/^r(\d+)c(\d+)$/);
  if (!m) throw new Error(`Invalid cell key: ${key}`);
  return { row: Number(m[1]), col: Number(m[2]) };
}

/**
 * Stable stringify for objects so we can build deterministic signatures.
 * @param {any} value
 * @returns {string}
 */
function stableStringify(value) {
  if (value === null) return "null";
  const t = typeof value;
  if (t === "string") return JSON.stringify(value);
  if (t === "number" || t === "boolean") return JSON.stringify(value);
  if (t === "undefined") return "undefined";
  if (t === "bigint") return JSON.stringify(String(value));
  if (t === "function") return "\"[Function]\"";
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  if (t === "object") {
    const keys = Object.keys(value).sort();
    return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

/**
 * @param {Cell} cell
 * @returns {string}
 */
function cellSignature(cell) {
  const enc = cell?.enc;
  const isEncrypted = enc !== null && enc !== undefined;
  const normalized = {
    enc: isEncrypted ? enc : null,
    value: isEncrypted ? null : (cell?.value ?? null),
    formula: isEncrypted ? null : (normalizeFormula(cell?.formula) ?? null),
    format: cell?.format ?? null,
  };
  return stableStringify(normalized);
}

/**
 * @param {Cell} a
 * @param {Cell} b
 */
function sameValueAndFormula(a, b) {
  const aEnc = a?.enc;
  const bEnc = b?.enc;
  const aEncrypted = aEnc !== null && aEnc !== undefined;
  const bEncrypted = bEnc !== null && bEnc !== undefined;
  if (aEncrypted || bEncrypted) {
    if (!aEncrypted || !bEncrypted) return false;
    return isDeepStrictEqual(aEnc, bEnc);
  }
  const av = a?.value ?? null;
  const bv = b?.value ?? null;
  const af = normalizeFormula(a?.formula) ?? null;
  const bf = normalizeFormula(b?.formula) ?? null;
  return isDeepStrictEqual(av, bv) && af === bf;
}

/**
 * @param {Cell} a
 * @param {Cell} b
 */
function sameFormat(a, b) {
  const af = a?.format ?? null;
  const bf = b?.format ?? null;
  return isDeepStrictEqual(af, bf);
}

/**
 * Semantic sheet diff:
 * - Detects added/removed/modified cells
 * - Detects moves by matching removed content to added content (value + normalized formula + format)
 * - Classifies format-only edits separately
 *
 * @param {SheetState} before
 * @param {SheetState} after
 * @returns {DiffResult}
 */
export function semanticDiff(before, after) {
  /** @type {DiffResult} */
  const result = { added: [], removed: [], modified: [], moved: [], formatOnly: [] };

  const beforeKeys = new Set(before.cells.keys());
  const afterKeys = new Set(after.cells.keys());

  /** @type {string[]} */
  const removedKeys = [];
  /** @type {string[]} */
  const addedKeys = [];

  for (const key of beforeKeys) {
    if (!afterKeys.has(key)) removedKeys.push(key);
  }
  for (const key of afterKeys) {
    if (!beforeKeys.has(key)) addedKeys.push(key);
  }

  // Common keys: modified / formatOnly detection.
  for (const key of beforeKeys) {
    if (!afterKeys.has(key)) continue;
    const beforeCell = before.cells.get(key);
    const afterCell = after.cells.get(key);
    if (sameValueAndFormula(beforeCell, afterCell)) {
      if (!sameFormat(beforeCell, afterCell)) {
        result.formatOnly.push({
          cell: parseCellKey(key),
          oldValue: beforeCell?.value ?? null,
          newValue: afterCell?.value ?? null,
          oldFormula: beforeCell?.formula ?? null,
          newFormula: afterCell?.formula ?? null,
        });
      }
      continue;
    }
    result.modified.push({
      cell: parseCellKey(key),
      oldValue: beforeCell?.value ?? null,
      newValue: afterCell?.value ?? null,
      oldFormula: beforeCell?.formula ?? null,
      newFormula: afterCell?.formula ?? null,
    });
  }

  // Build signature index for added keys to detect moves.
  /** @type {Map<string, string[]>} */
  const addedBySig = new Map();
  for (const key of addedKeys) {
    const cell = after.cells.get(key);
    const sig = cellSignature(cell);
    const list = addedBySig.get(sig) ?? [];
    list.push(key);
    addedBySig.set(sig, list);
  }

  /** @type {Set<string>} */
  const movedDestinations = new Set();
  for (const key of removedKeys) {
    const cell = before.cells.get(key);
    const sig = cellSignature(cell);
    const list = addedBySig.get(sig);
    if (list && list.length > 0) {
      const destKey = list.shift();
      movedDestinations.add(destKey);
      result.moved.push({
        oldLocation: parseCellKey(key),
        newLocation: parseCellKey(destKey),
        value: cell?.value ?? null,
        formula: cell?.formula ?? null,
      });
      continue;
    }
    result.removed.push({
      cell: parseCellKey(key),
      oldValue: cell?.value ?? null,
      oldFormula: cell?.formula ?? null,
    });
  }

  for (const key of addedKeys) {
    if (movedDestinations.has(key)) continue;
    const cell = after.cells.get(key);
    result.added.push({
      cell: parseCellKey(key),
      newValue: cell?.value ?? null,
      newFormula: cell?.formula ?? null,
    });
  }

  return result;
}
