import { cellsEqual, normalizeCell } from "./cell.js";
import { normalizeFormula } from "../../src/formula/normalize.js";

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
 * @param {import("./types.js").Cell | null | undefined} cell
 * @returns {string | null}
 */
function cellFingerprint(cell) {
  const normalized = normalizeCell(cell);
  if (normalized === null) return null;

  return stableStringify({
    enc: normalized.enc,
    value: normalized.value ?? null,
    formula: normalizeFormula(normalized.formula) ?? null,
    format: normalized.format ?? null
  });
}

/**
 * Detects simple single-cell moves relative to a base state.
 *
 * This is a deliberately conservative heuristic:
 * - only treats a (delete at X, add at Y) as a move if the deleted cell content
 *   exactly equals the added cell content (including format).
 * - only matches 1:1 (first unique match wins).
 *
 * @param {import("./types.js").CellMap} base
 * @param {import("./types.js").CellMap} next
 * @returns {Map<string, string>} map of `from -> to`
 */
export function detectCellMoves(base, next) {
  /** @type {Map<string, string>} */
  const moves = new Map();

  /** @type {Map<string, string[]>} */
  const additionsByFingerprint = new Map();

  for (const [addr, nextCell] of Object.entries(next)) {
    const baseCell = base[addr];
    if (!cellsEqual(baseCell, nextCell) && normalizeCell(baseCell) === null) {
      const fingerprint = cellFingerprint(nextCell);
      if (!fingerprint) continue;
      const list = additionsByFingerprint.get(fingerprint) ?? [];
      list.push(addr);
      additionsByFingerprint.set(fingerprint, list);
    }
  }

  /** @type {Set<string>} */
  const consumedAddrs = new Set();

  for (const [addr, baseCell] of Object.entries(base)) {
    const nextCell = next[addr];
    if (!cellsEqual(baseCell, nextCell) && normalizeCell(nextCell) === null) {
      const fingerprint = cellFingerprint(baseCell);
      if (!fingerprint) continue;
      const candidateAddrs = additionsByFingerprint.get(fingerprint) ?? [];
      const target = candidateAddrs.find((a) => !consumedAddrs.has(a));
      if (!target) continue;
      moves.set(addr, target);
      consumedAddrs.add(target);
    }
  }

  return moves;
}

/**
 * Applies a move map `from -> to` to a non-base sheet state.
 *
 * This is used by the merge engine to operate in a "rename-aware" coordinate
 * system, similar to Git's rename detection:
 * - if one branch moved a cell and the other branch didn't touch the
 *   destination, we treat the other branch's changes to the source cell as
 *   changes to the moved cell.
 *
 * The move is applied when:
 * - `sheet[from]` exists (non-empty)
 * - `sheet[to]` is unchanged relative to `base[to]` (so we don't overwrite)
 *
 * @param {import("./types.js").CellMap} base
 * @param {import("./types.js").CellMap} sheet
 * @param {Map<string, string>} moveMap
 * @returns {import("./types.js").CellMap}
 */
export function applyCellMovesToSheet(base, sheet, moveMap) {
  /** @type {import("./types.js").CellMap} */
  const out = { ...sheet };

  for (const [from, to] of moveMap.entries()) {
    const baseTo = base[to];
    const fromCell = sheet[from];
    const toCell = sheet[to];

    // Only relocate if destination wasn't touched (still equals base).
    if (!cellsEqual(toCell, baseTo)) continue;
    if (fromCell === undefined) continue;

    out[to] = fromCell;
    delete out[from];
  }

  return out;
}

/**
 * Applies a move map to the base sheet itself.
 *
 * This gives the merge algorithm a base representation that aligns with moved
 * coordinates, so that "move in one branch + edit in the other" can merge
 * without spurious add/add conflicts at the destination cell.
 *
 * @param {import("./types.js").CellMap} base
 * @param {Map<string, string>} moveMap
 * @returns {import("./types.js").CellMap}
 */
export function applyCellMovesToBaseSheet(base, moveMap) {
  /** @type {import("./types.js").CellMap} */
  const out = { ...base };

  for (const [from, to] of moveMap.entries()) {
    const fromCell = base[from];
    if (fromCell === undefined) continue;

    out[to] = fromCell;
    delete out[from];
  }

  return out;
}
