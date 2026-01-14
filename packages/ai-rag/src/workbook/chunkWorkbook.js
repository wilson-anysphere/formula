import { extractCells } from "./extractCells.js";
import { rectIntersectionArea, rectSize, rectToA1 } from "./rect.js";
import { getSheetCellMap, getSheetMatrix } from "./normalizeCell.js";
import { throwIfAborted } from "../utils/abort.js";

const DEFAULT_EXTRACT_MAX_ROWS = 50;
const DEFAULT_EXTRACT_MAX_COLS = 50;
// Region detection for matrix-backed sheets can allocate large visited grids.
// Cap the number of cells we consider to avoid catastrophic allocations on
// Excel-scale sheets.
const DEFAULT_DETECT_REGIONS_CELL_LIMIT = 200000;
const DEFAULT_MAX_DATA_REGIONS_PER_SHEET = 50;
const DEFAULT_MAX_FORMULA_REGIONS_PER_SHEET = 50;

/**
 * Encode a user-controlled id segment so delimiters in names cannot create
 * collisions between different (sheetName, tableName, ...) tuples.
 *
 * Keep "kind" segments (e.g. "table") unencoded for readability.
 *
 * @param {unknown} value
 */
function encodeIdPart(value) {
  // Avoid calling `String(...)` on arbitrary objects: custom `toString()` implementations
  // can throw or leak sensitive strings into persisted chunk ids.
  const raw =
    typeof value === "string"
      ? value
      : typeof value === "number" || typeof value === "boolean" || typeof value === "bigint"
        ? String(value)
        : "[invalid]";
  return encodeURIComponent(raw);
}

const NON_WHITESPACE_RE = /\S/;
// Equivalent to:
//   const trimmed = String(formula).trim();
//   if (trimmed === "") return false;
//   const strippedLeading = trimmed.startsWith("=") ? trimmed.slice(1) : trimmed;
//   return strippedLeading.trim() !== "";
const NON_EMPTY_FORMULA_RE = /^\s*=?\s*\S/;
// Like `NON_EMPTY_FORMULA_RE`, but requires an explicit leading "=" after whitespace.
const FORMULA_STRING_RE = /^\s*=\s*\S/;

function hasNonEmptyFormulaText(formula) {
  if (formula == null) return false;
  const text = typeof formula === "string" ? formula : String(formula);
  return NON_EMPTY_FORMULA_RE.test(text);
}

/**
 * @param {unknown} value
 */
function hasNonEmptyTrimmedText(value) {
  if (value == null) return false;
  const text = typeof value === "string" ? value : String(value);
  return NON_WHITESPACE_RE.test(text);
}

function isNonEmptyCell(cell) {
  if (cell == null) return false;

  // Strings support both raw values and formulas (bundler-style sheet matrices sometimes store scalars).
  if (typeof cell === "string") return hasNonEmptyFormulaText(cell);

  // Preserve rich object values as "non-empty" (matches normalizeCell's fallback).
  if (cell instanceof Date) return true;

  if (typeof cell === "object" && !Array.isArray(cell)) {
    const hasVF =
      Object.prototype.hasOwnProperty.call(cell, "v") || Object.prototype.hasOwnProperty.call(cell, "f");
    if (hasVF) {
      const f = cell.f;
      if (typeof f === "string") {
        if (hasNonEmptyFormulaText(f)) return true;
      } else if (f != null && hasNonEmptyTrimmedText(f)) {
        return true;
      }

      const v = cell.v;
      if (v == null) return false;
      if (typeof v === "string") return hasNonEmptyTrimmedText(v);
      return true;
    }

    const hasValueFormula =
      Object.prototype.hasOwnProperty.call(cell, "value") || Object.prototype.hasOwnProperty.call(cell, "formula");
    if (hasValueFormula) {
      const formula = cell.formula;
      if (typeof formula === "string" && hasNonEmptyFormulaText(formula)) return true;

      const value = cell.value;
      if (value == null || value === "") return false;
      if (typeof value === "string") return hasNonEmptyTrimmedText(value);
      return true;
    }

    // Treat `{}` as an empty cell; it's a common sparse representation.
    if (cell.constructor === Object) {
      // Avoid `Object.keys()` allocation in the common `{}` sparse representation.
      for (const _ in cell) return true;
      return false;
    }
    // Any other object is considered a value.
    return true;
  }

  // Numbers, booleans, etc.
  return true;
}

function isFormulaCell(cell) {
  if (cell == null) return false;
  if (typeof cell === "string") {
    return FORMULA_STRING_RE.test(cell);
  }

  if (typeof cell === "object" && !Array.isArray(cell)) {
    if (Object.prototype.hasOwnProperty.call(cell, "f")) {
      const f = cell.f;
      if (typeof f === "string") return hasNonEmptyFormulaText(f);
      return f != null && hasNonEmptyTrimmedText(f);
    }
    if (Object.prototype.hasOwnProperty.call(cell, "formula")) {
      const formula = cell.formula;
      if (typeof formula !== "string") return false;
      return hasNonEmptyFormulaText(formula);
    }
  }

  return false;
}

function hasNonFormulaNonEmptyCell(cells, signal) {
  // This can be large when callers raise extractMaxRows/Cols. Keep it abortable.
  let abortCountdown = 0;
  for (const row of cells) {
    throwIfAborted(signal);
    for (const cell of row) {
      if (abortCountdown === 0) {
        throwIfAborted(signal);
        abortCountdown = 256;
      }
      abortCountdown -= 1;
      if (isNonEmptyCell(cell) && !isFormulaCell(cell)) {
        throwIfAborted(signal);
        return true;
      }
    }
  }
  throwIfAborted(signal);
  return false;
}

/**
 * Packed coordinate key used to avoid allocating `${row},${col}` strings for
 * region detection.
 *
 * We prefer a packed Number key when the result stays within the 53-bit safe-integer
 * range (fast, cheap),
 * otherwise we fall back to BigInt packing. If BigInt isn't available at runtime,
 * we fall back to the original string key representation.
 *
 * @typedef {number | bigint | string} CoordKey
 */

// Pack row/col into a JS Number when the result stays within the 53-bit safe integer
// range. This is equivalent to `(row << 20) | col` as an *integer* operation, but
// avoids JS bitwise 32-bit truncation so it also works for large row indices.
const PACK_COL_BITS = 20;
const PACK_COL_FACTOR = 1 << PACK_COL_BITS; // 2^20
const MAX_SAFE_PACKED_ROW = Math.floor(Number.MAX_SAFE_INTEGER / PACK_COL_FACTOR);
const MAX_UINT32 = 2 ** 32 - 1;
const HAS_BIGINT = typeof BigInt === "function";
const BIGINT_SHIFT_32 = HAS_BIGINT ? BigInt(32) : null;
const BIGINT_MASK_32 = HAS_BIGINT ? BigInt(MAX_UINT32) : null;
const BIGINT_ONE = HAS_BIGINT ? BigInt(1) : null;
const BIGINT_ROW_STEP = HAS_BIGINT ? (BigInt(1) << /** @type {bigint} */ (BIGINT_SHIFT_32)) : null;

/**
 * @param {number} row
 * @param {number} col
 * @returns {CoordKey}
 */
function packCoordKey(row, col) {
  // Fast path: pack into a Number when it remains a safe integer. This covers
  // Excel-scale sheets (rows up to ~1M) without allocating strings or BigInts.
  if (row >= 0 && col >= 0 && col < PACK_COL_FACTOR && row <= MAX_SAFE_PACKED_ROW) {
    return row * PACK_COL_FACTOR + col;
  }

  // General path: pack into a BigInt (row in high 32 bits, col in low 32 bits).
  if (HAS_BIGINT && row >= 0 && col >= 0 && col <= MAX_UINT32) {
    const shift = /** @type {bigint} */ (BIGINT_SHIFT_32);
    return (BigInt(row) << shift) | BigInt(col);
  }

  // Last resort: preserve legacy behavior.
  return `${row},${col}`;
}

/**
 * @param {import('./workbookTypes').Workbook} workbook
 * @returns {Map<string, import('./workbookTypes').Sheet>}
 */
function sheetMap(workbook) {
  const map = new Map();
  for (const s of workbook.sheets || []) {
    if (!s || typeof s !== "object") continue;
    if (typeof s.name !== "string" || s.name.trim() === "") continue;
    map.set(s.name, s);
  }
  return map;
}

/**
 * Detect connected regions (4-neighbor) for a predicate over sheet cells.
 *
 * @param {import('./workbookTypes').Sheet} sheet
 * @param {(cell: any) => boolean} predicate
 * @param {{ signal?: AbortSignal, cellLimit?: number } | undefined} [opts]
 * @returns {{
 *   components: { rect: { r0: number, c0: number, r1: number, c1: number }, count: number }[],
 *   truncated: boolean,
 *   boundsRect: { r0: number, c0: number, r1: number, c1: number } | null
 * }}
 */
function detectRegions(sheet, predicate, opts) {
  const signal = opts?.signal;
  const cellLimit = opts?.cellLimit ?? DEFAULT_DETECT_REGIONS_CELL_LIMIT;
  throwIfAborted(signal);
  const matrix = getSheetMatrix(sheet);
  if (matrix) {
    /** @type {Set<CoordKey>} */
    const coords = new Set();
    let truncated = false;
    let minRow = Number.POSITIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxRow = Number.NEGATIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;

    // Treat matrix-backed sheets as sparse: use `for..in` to iterate only defined
    // rows/cols (avoids scanning/allocating for large sparse arrays).
    try {
      for (const rKey in matrix) {
        throwIfAborted(signal);
        const r = Number(rKey);
        if (!Number.isInteger(r) || r < 0) continue;
        const row = matrix[r];
        if (!Array.isArray(row)) continue;
        for (const cKey in row) {
          throwIfAborted(signal);
          const c = Number(cKey);
          if (!Number.isInteger(c) || c < 0) continue;
          if (!predicate(row[c])) continue;
          coords.add(packCoordKey(r, c));
          minRow = Math.min(minRow, r);
          minCol = Math.min(minCol, c);
          maxRow = Math.max(maxRow, r);
          maxCol = Math.max(maxCol, c);

          if (coords.size > cellLimit) {
            truncated = true;
            break;
          }
        }
        if (truncated) break;
      }
    } catch {
      // Fall back to no regions on unexpected enumerable shapes.
      return { components: [], truncated: false, boundsRect: null };
    }

    if (coords.size === 0)
      return { components: [], truncated: false, boundsRect: null };

    /** @type {{ rect: { r0: number, c0: number, r1: number, c1: number }, count: number }[]} */
    const components = [];

    for (const startKey of coords) {
      throwIfAborted(signal);
      // The frontier set is mutated during flood fill; a startKey can be deleted as
      // part of a previous component. Skip if it's no longer present.
      if (!coords.delete(startKey)) continue;
      const stack = [startKey];

      let r0 = Number.POSITIVE_INFINITY;
      let r1 = Number.NEGATIVE_INFINITY;
      let c0 = Number.POSITIVE_INFINITY;
      let c1 = Number.NEGATIVE_INFINITY;
      let count = 0;

      while (stack.length) {
        throwIfAborted(signal);
        const curKey = stack.pop();
        if (curKey == null) continue;
        count += 1;
        if (typeof curKey === "number") {
          const r = Math.floor(curKey / PACK_COL_FACTOR);
          const c = curKey - r * PACK_COL_FACTOR;
          r0 = Math.min(r0, r);
          r1 = Math.max(r1, r);
          c0 = Math.min(c0, c);
          c1 = Math.max(c1, c);

          if (r > 0) {
            const nk = curKey - PACK_COL_FACTOR;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (r < MAX_SAFE_PACKED_ROW) {
            const nk = curKey + PACK_COL_FACTOR;
            if (coords.delete(nk)) stack.push(nk);
          } else {
            const nk = packCoordKey(r + 1, c);
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c > 0) {
            const nk = curKey - 1;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c + 1 < PACK_COL_FACTOR) {
            const nk = curKey + 1;
            if (coords.delete(nk)) stack.push(nk);
          } else {
            const nk = packCoordKey(r, c + 1);
            if (coords.delete(nk)) stack.push(nk);
          }

          continue;
        }

        if (typeof curKey === "bigint") {
          const shift = /** @type {bigint} */ (BIGINT_SHIFT_32);
          const mask = /** @type {bigint} */ (BIGINT_MASK_32);
          const rowStep = /** @type {bigint} */ (BIGINT_ROW_STEP);
          const one = /** @type {bigint} */ (BIGINT_ONE);
          const r = Number(curKey >> shift);
          const c = Number(curKey & mask);
          r0 = Math.min(r0, r);
          r1 = Math.max(r1, r);
          c0 = Math.min(c0, c);
          c1 = Math.max(c1, c);

          if (r > 0) {
            // If the neighbor would be represented as a Number key, use that form so
            // cross-representation boundaries (BigInt <-> Number) still connect.
            const nk =
              r - 1 <= MAX_SAFE_PACKED_ROW && c < PACK_COL_FACTOR
                ? (r - 1) * PACK_COL_FACTOR + c
                : curKey - rowStep;
            if (coords.delete(nk)) stack.push(nk);
          }

          {
            const nk = curKey + rowStep;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c > 0) {
            const nk =
              r <= MAX_SAFE_PACKED_ROW && c - 1 < PACK_COL_FACTOR
                ? r * PACK_COL_FACTOR + (c - 1)
                : curKey - one;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c < MAX_UINT32) {
            const nk = curKey + one;
            if (coords.delete(nk)) stack.push(nk);
          } else {
            const nk = packCoordKey(r, c + 1);
            if (coords.delete(nk)) stack.push(nk);
          }

          continue;
        }

        const idx = curKey.indexOf(",");
        const r = Number(curKey.slice(0, idx));
        const c = Number(curKey.slice(idx + 1));
        r0 = Math.min(r0, r);
        r1 = Math.max(r1, r);
        c0 = Math.min(c0, c);
        c1 = Math.max(c1, c);

        if (r > 0) {
          const nk = packCoordKey(r - 1, c);
          if (coords.delete(nk)) stack.push(nk);
        }

        {
          const nk = packCoordKey(r + 1, c);
          if (coords.delete(nk)) stack.push(nk);
        }

        if (c > 0) {
          const nk = packCoordKey(r, c - 1);
          if (coords.delete(nk)) stack.push(nk);
        }

        {
          const nk = packCoordKey(r, c + 1);
          if (coords.delete(nk)) stack.push(nk);
        }
      }

      components.push({ rect: { r0, c0, r1, c1 }, count });
    }

    components.sort(
      (a, b) =>
        a.rect.r0 - b.rect.r0 ||
        a.rect.c0 - b.rect.c0 ||
        a.rect.r1 - b.rect.r1 ||
        a.rect.c1 - b.rect.c1
    );

    /** @type {{ r0: number, c0: number, r1: number, c1: number } | null} */
    const boundsRect =
      minRow === Number.POSITIVE_INFINITY
        ? null
        : {
            r0: minRow,
            c0: minCol,
            r1: maxRow,
            c1: maxCol,
          };

    // Drop trivial single-cell regions (often incidental labels).
    return {
      components: components.filter((c) => c.count >= 2),
      truncated,
      boundsRect,
    };
  }

  const map = getSheetCellMap(sheet);
  if (map) {
    /**
     * Parse a non-negative integer from a substring without allocating intermediate strings.
     * Treat empty/whitespace-only segments as 0 (matching `Number("")` -> 0).
     *
     * @param {string} text
     * @param {number} start
     * @param {number} end
     * @returns {number | null}
     */
    function parseNonNegativeInt(text, start, end) {
      // Trim ASCII whitespace.
      while (start < end && text.charCodeAt(start) <= 32) start += 1;
      while (end > start && text.charCodeAt(end - 1) <= 32) end -= 1;
      if (start >= end) return 0;
      let acc = 0;
      for (let i = start; i < end; i += 1) {
        const code = text.charCodeAt(i);
        if (code < 48 || code > 57) {
          // Preserve legacy behavior for unusual key encodings (e.g. "1e3", "+1"):
          // fall back to `Number()` parsing, even though it allocates a substring.
          const num = Number(text.slice(start, end));
          if (!Number.isInteger(num) || num < 0) return null;
          return num;
        }
        acc = acc * 10 + (code - 48);
      }
      return acc;
    }

    /** @type {Set<CoordKey>} */
    const coords = new Set();
    let truncated = false;
    let minRow = Number.POSITIVE_INFINITY;
    let minCol = Number.POSITIVE_INFINITY;
    let maxRow = Number.NEGATIVE_INFINITY;
    let maxCol = Number.NEGATIVE_INFINITY;

    for (const [key, raw] of map.entries()) {
      throwIfAborted(signal);
      const rawKey = String(key);
      const commaIdx = rawKey.indexOf(",");
      const colonIdx = commaIdx >= 0 ? -1 : rawKey.indexOf(":");
      const idx = commaIdx >= 0 ? commaIdx : colonIdx;
      if (idx < 0) continue;
      const delimiter = commaIdx >= 0 ? "," : ":";
      // Reject keys that contain more than one delimiter (e.g. "1,2,3").
      if (rawKey.indexOf(delimiter, idx + 1) !== -1) continue;

      const row = parseNonNegativeInt(rawKey, 0, idx);
      const col = parseNonNegativeInt(rawKey, idx + 1, rawKey.length);
      if (row == null || col == null) continue;
      if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) continue;
      if (!predicate(raw)) continue;
      coords.add(packCoordKey(row, col));

      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);

      if (coords.size > cellLimit) {
        truncated = true;
        break;
      }
    }

    if (coords.size === 0)
      return { components: [], truncated: false, boundsRect: null };

    /** @type {{ rect: { r0: number, c0: number, r1: number, c1: number }, count: number }[]} */
    const components = [];

    for (const startKey of coords) {
      throwIfAborted(signal);
      if (!coords.delete(startKey)) continue;
      const stack = [startKey];

      let r0 = Number.POSITIVE_INFINITY;
      let r1 = Number.NEGATIVE_INFINITY;
      let c0 = Number.POSITIVE_INFINITY;
      let c1 = Number.NEGATIVE_INFINITY;
      let count = 0;

      while (stack.length) {
        throwIfAborted(signal);
        const curKey = stack.pop();
        if (curKey == null) continue;
        count += 1;
        if (typeof curKey === "number") {
          const r = Math.floor(curKey / PACK_COL_FACTOR);
          const c = curKey - r * PACK_COL_FACTOR;
          r0 = Math.min(r0, r);
          r1 = Math.max(r1, r);
          c0 = Math.min(c0, c);
          c1 = Math.max(c1, c);

          if (r > 0) {
            const nk = curKey - PACK_COL_FACTOR;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (r < MAX_SAFE_PACKED_ROW) {
            const nk = curKey + PACK_COL_FACTOR;
            if (coords.delete(nk)) stack.push(nk);
          } else {
            const nk = packCoordKey(r + 1, c);
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c > 0) {
            const nk = curKey - 1;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c + 1 < PACK_COL_FACTOR) {
            const nk = curKey + 1;
            if (coords.delete(nk)) stack.push(nk);
          } else {
            const nk = packCoordKey(r, c + 1);
            if (coords.delete(nk)) stack.push(nk);
          }

          continue;
        }

        if (typeof curKey === "bigint") {
          const shift = /** @type {bigint} */ (BIGINT_SHIFT_32);
          const mask = /** @type {bigint} */ (BIGINT_MASK_32);
          const rowStep = /** @type {bigint} */ (BIGINT_ROW_STEP);
          const one = /** @type {bigint} */ (BIGINT_ONE);
          const r = Number(curKey >> shift);
          const c = Number(curKey & mask);
          r0 = Math.min(r0, r);
          r1 = Math.max(r1, r);
          c0 = Math.min(c0, c);
          c1 = Math.max(c1, c);

          if (r > 0) {
            const nk =
              r - 1 <= MAX_SAFE_PACKED_ROW && c < PACK_COL_FACTOR
                ? (r - 1) * PACK_COL_FACTOR + c
                : curKey - rowStep;
            if (coords.delete(nk)) stack.push(nk);
          }

          {
            const nk = curKey + rowStep;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c > 0) {
            const nk =
              r <= MAX_SAFE_PACKED_ROW && c - 1 < PACK_COL_FACTOR
                ? r * PACK_COL_FACTOR + (c - 1)
                : curKey - one;
            if (coords.delete(nk)) stack.push(nk);
          }

          if (c < MAX_UINT32) {
            const nk = curKey + one;
            if (coords.delete(nk)) stack.push(nk);
          } else {
            const nk = packCoordKey(r, c + 1);
            if (coords.delete(nk)) stack.push(nk);
          }

          continue;
        }

        const idx = curKey.indexOf(",");
        const r = Number(curKey.slice(0, idx));
        const c = Number(curKey.slice(idx + 1));
        r0 = Math.min(r0, r);
        r1 = Math.max(r1, r);
        c0 = Math.min(c0, c);
        c1 = Math.max(c1, c);

        if (r > 0) {
          const nk = packCoordKey(r - 1, c);
          if (coords.delete(nk)) stack.push(nk);
        }

        {
          const nk = packCoordKey(r + 1, c);
          if (coords.delete(nk)) stack.push(nk);
        }

        if (c > 0) {
          const nk = packCoordKey(r, c - 1);
          if (coords.delete(nk)) stack.push(nk);
        }

        {
          const nk = packCoordKey(r, c + 1);
          if (coords.delete(nk)) stack.push(nk);
        }
      }

      components.push({ rect: { r0, c0, r1, c1 }, count });
    }

    components.sort(
      (a, b) =>
        a.rect.r0 - b.rect.r0 ||
        a.rect.c0 - b.rect.c0 ||
        a.rect.r1 - b.rect.r1 ||
        a.rect.c1 - b.rect.c1
    );

    /** @type {{ r0: number, c0: number, r1: number, c1: number } | null} */
    const boundsRect =
      minRow === Number.POSITIVE_INFINITY
        ? null
        : {
            r0: minRow,
            c0: minCol,
            r1: maxRow,
            c1: maxCol,
          };

    return {
      components: components.filter((c) => c.count >= 2),
      truncated,
      boundsRect,
    };
  }

  return { components: [], truncated: false, boundsRect: null };
}

/**
 * @param {{ r0: number, c0: number, r1: number, c1: number }} rect
 * @param {{ r0: number, c0: number, r1: number, c1: number }[]} existing
 */
function overlapsExisting(rect, existing) {
  for (const ex of existing) {
    const inter = rectIntersectionArea(rect, ex);
    if (inter === 0) continue;
    const ratio = inter / Math.min(rectSize(rect), rectSize(ex));
    if (ratio > 0.8) return true;
  }
  return false;
}

/**
 * Sort most-important first (largest component), deterministically.
 *
 * @param {{ rect: { r0: number, c0: number, r1: number, c1: number }, count: number }} a
 * @param {{ rect: { r0: number, c0: number, r1: number, c1: number }, count: number }} b
 */
function compareComponentImportance(a, b) {
  return (
    b.count - a.count ||
    a.rect.r0 - b.rect.r0 ||
    a.rect.c0 - b.rect.c0 ||
    a.rect.r1 - b.rect.r1 ||
    a.rect.c1 - b.rect.c1
  );
}

/**
 * Stable ordering for ids/tests (reading order).
 *
 * @param {{ rect: { r0: number, c0: number, r1: number, c1: number } }} a
 * @param {{ rect: { r0: number, c0: number, r1: number, c1: number } }} b
 */
function compareComponentRect(a, b) {
  return (
    a.rect.r0 - b.rect.r0 ||
    a.rect.c0 - b.rect.c0 ||
    a.rect.r1 - b.rect.r1 ||
    a.rect.c1 - b.rect.c1
  );
}

/**
 * @template T extends { rect: any, count: number }
 * @param {T[]} components
 * @param {number} max
 * @returns {T[]}
 */
function capComponents(components, max) {
  if (!Number.isFinite(max) || max <= 0) return [];
  if (components.length <= max) return components.slice().sort(compareComponentRect);
  const mostImportant = components.slice().sort(compareComponentImportance).slice(0, max);
  return mostImportant.sort(compareComponentRect);
}

/**
 * Chunk workbook into semantic regions.
 *
 * Strategy:
 * - Use explicit tables & named ranges first (stable, user-authored).
 * - Detect remaining data regions by connected non-empty cell blocks.
 * - Detect formula-heavy regions by connected formula blocks.
 *
 * @param {import('./workbookTypes').Workbook} workbook
 * @param {{
 *   signal?: AbortSignal,
 *   extractMaxRows?: number,
 *   extractMaxCols?: number,
 *   detectRegionsCellLimit?: number,
 *   maxDataRegionsPerSheet?: number,
 *   maxFormulaRegionsPerSheet?: number,
 *   // Back-compat alias: treat as maxDataRegionsPerSheet/maxFormulaRegionsPerSheet when provided.
 *   maxRegionsPerSheet?: number
 * }} [options]
 * @returns {import('./workbookTypes').WorkbookChunk[]}
 */
function chunkWorkbook(workbook, options = {}) {
  const signal = options.signal;
  const extractMaxRows = options.extractMaxRows ?? DEFAULT_EXTRACT_MAX_ROWS;
  const extractMaxCols = options.extractMaxCols ?? DEFAULT_EXTRACT_MAX_COLS;
  const detectRegionsCellLimit = options.detectRegionsCellLimit ?? DEFAULT_DETECT_REGIONS_CELL_LIMIT;
  const maxRegionsAlias = options.maxRegionsPerSheet;
  const maxDataRegionsPerSheet =
    options.maxDataRegionsPerSheet ??
    (typeof maxRegionsAlias === "number" ? maxRegionsAlias : DEFAULT_MAX_DATA_REGIONS_PER_SHEET);
  const maxFormulaRegionsPerSheet =
    options.maxFormulaRegionsPerSheet ??
    (typeof maxRegionsAlias === "number" ? maxRegionsAlias : DEFAULT_MAX_FORMULA_REGIONS_PER_SHEET);
  throwIfAborted(signal);
  const sheets = sheetMap(workbook);
  /** @type {import('./workbookTypes').WorkbookChunk[]} */
  const chunks = [];

  /** @type {{ sheetName: string, rect: any }[]} */
  const occupied = [];

  for (const table of workbook.tables || []) {
    throwIfAborted(signal);
    const sheetName = typeof table?.sheetName === "string" ? table.sheetName : "";
    if (!sheetName) continue;
    const sheet = sheets.get(sheetName);
    if (!sheet) continue;
    const id = `${encodeIdPart(workbook.id)}::${encodeIdPart(sheetName)}::table::${encodeIdPart(table?.name ?? "")}`;
    chunks.push({
      id,
      workbookId: workbook.id,
      sheetName,
      kind: "table",
      title: table?.name ?? "",
      rect: table.rect,
      cells: extractCells(sheet, table.rect, {
        maxRows: extractMaxRows,
        maxCols: extractMaxCols,
        signal,
      }),
      meta: { tableName: table?.name ?? "" },
    });
    occupied.push({ sheetName, rect: table.rect });
  }

  for (const nr of workbook.namedRanges || []) {
    throwIfAborted(signal);
    const sheetName = typeof nr?.sheetName === "string" ? nr.sheetName : "";
    if (!sheetName) continue;
    const sheet = sheets.get(sheetName);
    if (!sheet) continue;
    const id = `${encodeIdPart(workbook.id)}::${encodeIdPart(sheetName)}::namedRange::${encodeIdPart(nr?.name ?? "")}`;
    chunks.push({
      id,
      workbookId: workbook.id,
      sheetName,
      kind: "namedRange",
      title: nr?.name ?? "",
      rect: nr.rect,
      cells: extractCells(sheet, nr.rect, {
        maxRows: extractMaxRows,
        maxCols: extractMaxCols,
        signal,
      }),
      meta: { namedRange: nr?.name ?? "" },
    });
    occupied.push({ sheetName, rect: nr.rect });
  }

  for (const sheet of workbook.sheets || []) {
    throwIfAborted(signal);
    const sheetName = typeof sheet?.name === "string" ? sheet.name : "";
    if (!sheetName) continue;
    const existingRects = occupied
      .filter((o) => o.sheetName === sheetName)
      .map((o) => o.rect);

    const dataDetection = detectRegions(sheet, isNonEmptyCell, {
      signal,
      cellLimit: detectRegionsCellLimit,
    });
    const dataComponents = dataDetection.components.filter(
      (c) => !overlapsExisting(c.rect, existingRects)
    );

    /** @type {Array<{ rect: any, count: number, isTruncationFallback?: boolean }>} */
    const dataCandidates = dataComponents.map((c) => ({ ...c }));
    if (dataDetection.truncated && dataDetection.boundsRect) {
      if (!overlapsExisting(dataDetection.boundsRect, existingRects)) {
        dataCandidates.push({
          rect: dataDetection.boundsRect,
          count: Number.POSITIVE_INFINITY,
          isTruncationFallback: true,
        });
      }
    }

    const dataRegions = capComponents(dataCandidates, maxDataRegionsPerSheet);
    for (const region of dataRegions) {
      throwIfAborted(signal);
      const rect = region.rect;
      const cells = extractCells(sheet, rect, {
        maxRows: extractMaxRows,
        maxCols: extractMaxCols,
        signal,
      });

      // If this region is entirely formulas, prefer a formulaRegion chunk instead of
      // emitting a redundant dataRegion chunk that would suppress it.
      if (!hasNonFormulaNonEmptyCell(cells, signal)) continue;

      const coordKey = `${rect.r0},${rect.c0},${rect.r1},${rect.c1}`;
      const suffix = region.isTruncationFallback ? `truncated::${coordKey}` : coordKey;
      const id = `${encodeIdPart(workbook.id)}::${encodeIdPart(sheetName)}::dataRegion::${suffix}`;
      chunks.push({
        id,
        workbookId: workbook.id,
        sheetName,
        kind: "dataRegion",
        title: region.isTruncationFallback
          ? `Data region (truncated) ${rectToA1(rect)}`
          : `Data region ${rectToA1(rect)}`,
        rect,
        cells,
        meta: region.isTruncationFallback
          ? {
              truncated: true,
              reason: "detectRegionsCellLimit",
              detectRegionsCellLimit,
            }
          : undefined,
      });
      occupied.push({ sheetName, rect });
      existingRects.push(rect);
    }

    const formulaDetection = detectRegions(sheet, isFormulaCell, {
      signal,
      cellLimit: detectRegionsCellLimit,
    });
    const formulaComponents = formulaDetection.components.filter(
      (c) => !overlapsExisting(c.rect, existingRects)
    );

    /** @type {Array<{ rect: any, count: number, isTruncationFallback?: boolean }>} */
    const formulaCandidates = formulaComponents.map((c) => ({ ...c }));
    if (formulaDetection.truncated && formulaDetection.boundsRect) {
      if (!overlapsExisting(formulaDetection.boundsRect, existingRects)) {
        formulaCandidates.push({
          rect: formulaDetection.boundsRect,
          count: Number.POSITIVE_INFINITY,
          isTruncationFallback: true,
        });
      }
    }

    const formulaRegions = capComponents(formulaCandidates, maxFormulaRegionsPerSheet);
    for (const region of formulaRegions) {
      throwIfAborted(signal);
      const rect = region.rect;
      const coordKey = `${rect.r0},${rect.c0},${rect.r1},${rect.c1}`;
      const suffix = region.isTruncationFallback ? `truncated::${coordKey}` : coordKey;
      const id = `${encodeIdPart(workbook.id)}::${encodeIdPart(sheetName)}::formulaRegion::${suffix}`;
      chunks.push({
        id,
        workbookId: workbook.id,
        sheetName,
        kind: "formulaRegion",
        title: region.isTruncationFallback
          ? `Formula region (truncated) ${rectToA1(rect)}`
          : `Formula region ${rectToA1(rect)}`,
        rect,
        cells: extractCells(sheet, rect, {
          maxRows: extractMaxRows,
          maxCols: extractMaxCols,
          signal,
        }),
        meta: region.isTruncationFallback
          ? {
              truncated: true,
              reason: "detectRegionsCellLimit",
              detectRegionsCellLimit,
            }
          : undefined,
      });
    }
  }

  return chunks;
}

export { chunkWorkbook };
