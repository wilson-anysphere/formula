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
 * @param {string} key
 * @returns {CellRef}
 */
export function parseCellKey(key) {
  if (typeof key !== "string" || key.length < 3) throw new Error(`Invalid cell key: ${key}`);
  if (key.charCodeAt(0) !== 114) throw new Error(`Invalid cell key: ${key}`); // 'r'
  const cIdx = key.indexOf("c", 1);
  if (cIdx === -1) throw new Error(`Invalid cell key: ${key}`);
  const row = parseUnsignedInt(key, 1, cIdx);
  const col = parseUnsignedInt(key, cIdx + 1, key.length);
  if (row == null || col == null) throw new Error(`Invalid cell key: ${key}`);
  return { row, col };
}

/**
 * Stable stringify for objects so we can build deterministic signatures.
 *
 * This is intentionally defensive (cycle-safe) so diff helpers can be used in
 * browser bundles without crashing on unexpected input graphs.
 * @param {any} value
 * @returns {string}
 */
function stableStringify(value) {
  /** @type {WeakSet<object>} */
  const stack = new WeakSet();
  return stableStringifyInner(value, stack);
}

/**
 * @param {any} value
 * @param {WeakSet<object>} stack
 * @returns {string}
 */
function stableStringifyInner(value, stack) {
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
  if (t === "symbol") return String(value);
  if (t === "function") return "\"[Function]\"";

  if (Array.isArray(value)) {
    if (stack.has(value)) return "[Circular]";
    stack.add(value);
    try {
      // Preserve array holes so `[ , ]` doesn't collapse to `[]`.
      /** @type {string[]} */
      const parts = [];
      for (let i = 0; i < value.length; i += 1) {
        if (!Object.prototype.hasOwnProperty.call(value, i)) {
          parts.push("<hole>");
        } else {
          parts.push(stableStringifyInner(value[i], stack));
        }
      }
      return `[${parts.join(",")}]`;
    } finally {
      stack.delete(value);
    }
  }

  if (t === "object") {
    if (stack.has(value)) return "[Circular]";
    stack.add(value);
    try {
      const tag = Object.prototype.toString.call(value);
      if (tag === "[object Date]" && typeof value.toISOString === "function") {
        return `Date(${value.toISOString()})`;
      }
      if (tag === "[object RegExp]") {
        return `RegExp(${value.source},${value.flags})`;
      }

      const keys = Object.keys(value).sort();
      const body = `{${keys.map((k) => `${JSON.stringify(k)}:${stableStringifyInner(value[k], stack)}`).join(",")}}`;
      // Prefix non-plain objects with their tag to avoid collisions with `{}`.
      return tag === "[object Object]" ? body : `${tag}${body}`;
    } finally {
      stack.delete(value);
    }
  }

  return JSON.stringify(value);
}

/**
 * @param {Cell | undefined} cell
 * @returns {{ encrypted: boolean, keyId: string | null }}
 */
function encryptionMeta(cell) {
  const enc = cell?.enc;
  // Backwards compatibility: some legacy snapshots explicitly stored `enc: null` for
  // plaintext cells. Treat `null` the same as "unset" so diffs don't hide values.
  //
  // Fail closed for all other non-null markers: if an `enc` field exists with an
  // unexpected shape, consider the cell encrypted so diffs never fall back to
  // potentially-sensitive plaintext.
  if (enc == null) return { encrypted: false, keyId: null };
  const keyId = enc && typeof enc === "object" && typeof enc.keyId === "string" ? enc.keyId : null;
  return { encrypted: true, keyId };
}

/**
 * @param {Cell} cell
 * @returns {string}
 */
function cellSignature(cell) {
  const enc = cell?.enc;
  const isEncrypted = enc != null;
  const normalized = {
    // Preserve `undefined` vs `null` for `enc` so move detection never conflates
    // unencrypted format-only cells with encrypted marker cells (`enc: null`).
    enc,
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
  const aEncrypted = aEnc != null;
  const bEncrypted = bEnc != null;
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

  /** @type {string[]} */
  const removedKeys = [];
  /** @type {string[]} */
  const addedKeys = [];

  // Common keys: modified / formatOnly detection + collect removed keys.
  for (const key of before.cells.keys()) {
    if (!after.cells.has(key)) {
      removedKeys.push(key);
      continue;
    }
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

  // Collect added keys.
  for (const key of after.cells.keys()) {
    if (!before.cells.has(key)) addedKeys.push(key);
  }

  // Build signature index for added keys to detect moves.
  /** @type {Map<string, { keys: string[], idx: number }>} */
  const addedBySig = new Map();
  for (const key of addedKeys) {
    const cell = after.cells.get(key);
    const sig = cellSignature(cell);
    let entry = addedBySig.get(sig);
    if (!entry) {
      entry = { keys: [], idx: 0 };
      addedBySig.set(sig, entry);
    }
    entry.keys.push(key);
  }

  /** @type {Set<string>} */
  const movedDestinations = new Set();
  for (const key of removedKeys) {
    const cell = before.cells.get(key);
    const sig = cellSignature(cell);
    const entry = addedBySig.get(sig);
    if (entry && entry.idx < entry.keys.length) {
      const encMeta = encryptionMeta(cell);
      const destKey = entry.keys[entry.idx++];
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
