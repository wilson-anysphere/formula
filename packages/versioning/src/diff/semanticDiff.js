import { normalizeFormula } from "../formula/normalize.js";
import { deepEqual } from "./deepEqual.js";

/**
 * @typedef {{ row: number, col: number }} CellRef
 * @typedef {{ value?: any, formula?: string | null, format?: any, enc?: any }} Cell
 *
 * @typedef {{ cells: Map<string, Cell> }} SheetState
 *
 * @typedef {{
 *   cell: CellRef,
 *   oldValue?: any,
 *   newValue?: any,
 *   oldFormula?: string | null,
 *   newFormula?: string | null,
 *   oldEncrypted?: boolean,
 *   newEncrypted?: boolean,
 *   oldKeyId?: string | null,
 *   newKeyId?: string | null,
 * }} CellChange
 * @typedef {{
 *   oldLocation: CellRef,
 *   newLocation: CellRef,
 *   value: any,
 *   formula?: string | null,
 *   encrypted?: boolean,
 *   keyId?: string | null,
 * }} MoveChange
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
  if (t === "number") {
    if (Number.isNaN(value)) return "NaN";
    if (Object.is(value, -0)) return "-0";
    if (!Number.isFinite(value)) return String(value);
    return JSON.stringify(value);
  }
  if (t === "boolean") return JSON.stringify(value);
  if (t === "undefined") return "undefined";
  // BigInt is not JSON-serializable; preserve type to avoid collisions with strings.
  if (t === "bigint") return `BigInt(${String(value)})`;
  if (t === "function") return "\"[Function]\"";
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(",")}]`;
  if (t === "object") {
    const keys = Object.keys(value).sort();
    return `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringify(value[k])}`).join(",")}}`;
  }
  return JSON.stringify(value);
}

/**
 * @param {Cell | undefined} cell
 * @returns {{ encrypted: boolean, keyId: string | null }}
 */
function encryptionMeta(cell) {
  const enc = cell?.enc;
  if (enc === null || enc === undefined) return { encrypted: false, keyId: null };
  const keyId = enc && typeof enc === "object" && typeof enc.keyId === "string" ? enc.keyId : null;
  return { encrypted: true, keyId };
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
    return deepEqual(aEnc, bEnc);
  }
  const av = a?.value ?? null;
  const bv = b?.value ?? null;
  const af = normalizeFormula(a?.formula) ?? null;
  const bf = normalizeFormula(b?.formula) ?? null;
  return deepEqual(av, bv) && af === bf;
}

/**
 * @param {Cell} a
 * @param {Cell} b
 */
function sameFormat(a, b) {
  const af = a?.format ?? null;
  const bf = b?.format ?? null;
  return deepEqual(af, bf);
}

/**
 * Semantic sheet diff:
 * - Detects added/removed/modified cells
 * - Detects moves by matching removed content to added content (value + normalized formula + format + enc)
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
    const beforeEncMeta = encryptionMeta(beforeCell);
    const afterEncMeta = encryptionMeta(afterCell);
    if (sameValueAndFormula(beforeCell, afterCell)) {
      if (!sameFormat(beforeCell, afterCell)) {
        result.formatOnly.push({
          cell: parseCellKey(key),
          oldValue: beforeEncMeta.encrypted ? null : (beforeCell?.value ?? null),
          newValue: afterEncMeta.encrypted ? null : (afterCell?.value ?? null),
          oldFormula: beforeEncMeta.encrypted ? null : (beforeCell?.formula ?? null),
          newFormula: afterEncMeta.encrypted ? null : (afterCell?.formula ?? null),
          ...(beforeEncMeta.encrypted || afterEncMeta.encrypted
            ? {
                oldEncrypted: beforeEncMeta.encrypted,
                newEncrypted: afterEncMeta.encrypted,
                oldKeyId: beforeEncMeta.keyId,
                newKeyId: afterEncMeta.keyId,
              }
            : {}),
        });
      }
      continue;
    }
    result.modified.push({
      cell: parseCellKey(key),
      oldValue: beforeEncMeta.encrypted ? null : (beforeCell?.value ?? null),
      newValue: afterEncMeta.encrypted ? null : (afterCell?.value ?? null),
      oldFormula: beforeEncMeta.encrypted ? null : (beforeCell?.formula ?? null),
      newFormula: afterEncMeta.encrypted ? null : (afterCell?.formula ?? null),
      ...(beforeEncMeta.encrypted || afterEncMeta.encrypted
        ? {
            oldEncrypted: beforeEncMeta.encrypted,
            newEncrypted: afterEncMeta.encrypted,
            oldKeyId: beforeEncMeta.keyId,
            newKeyId: afterEncMeta.keyId,
          }
        : {}),
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
      const encMeta = encryptionMeta(cell);
      const destKey = list.shift();
      movedDestinations.add(destKey);
      result.moved.push({
        oldLocation: parseCellKey(key),
        newLocation: parseCellKey(destKey),
        value: encMeta.encrypted ? null : (cell?.value ?? null),
        formula: encMeta.encrypted ? null : (cell?.formula ?? null),
        ...(encMeta.encrypted ? { encrypted: true, keyId: encMeta.keyId } : {}),
      });
      continue;
    }
    const encMeta = encryptionMeta(cell);
    result.removed.push({
      cell: parseCellKey(key),
      oldValue: encMeta.encrypted ? null : (cell?.value ?? null),
      oldFormula: encMeta.encrypted ? null : (cell?.formula ?? null),
      ...(encMeta.encrypted ? { oldEncrypted: true, oldKeyId: encMeta.keyId } : {}),
    });
  }

  for (const key of addedKeys) {
    if (movedDestinations.has(key)) continue;
    const cell = after.cells.get(key);
    const encMeta = encryptionMeta(cell);
    result.added.push({
      cell: parseCellKey(key),
      newValue: encMeta.encrypted ? null : (cell?.value ?? null),
      newFormula: encMeta.encrypted ? null : (cell?.formula ?? null),
      ...(encMeta.encrypted ? { newEncrypted: true, newKeyId: encMeta.keyId } : {}),
    });
  }

  return result;
}
