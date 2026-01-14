import { normalizeFormula } from "../../src/formula/normalize.js";

/**
 * @param {any} value
 */
function isPlainObject(value) {
  return (
    value !== null &&
    typeof value === "object" &&
    (Object.getPrototypeOf(value) === Object.prototype ||
      Object.getPrototypeOf(value) === null)
  );
}

/**
 * @param {unknown} a
 * @param {unknown} b
 * @returns {boolean}
 */
export function deepEqual(a, b) {
  if (a === b) return true;
  if (Number.isNaN(a) && Number.isNaN(b)) return true;

  if (Array.isArray(a) && Array.isArray(b)) {
    if (a.length !== b.length) return false;
    for (let i = 0; i < a.length; i += 1) {
      if (!deepEqual(a[i], b[i])) return false;
    }
    return true;
  }

  if (isPlainObject(a) && isPlainObject(b)) {
    const keysA = Object.keys(a);
    const keysB = Object.keys(b);
    if (keysA.length !== keysB.length) return false;
    for (const key of keysA) {
      if (!Object.prototype.hasOwnProperty.call(b, key)) return false;
      if (!deepEqual(a[key], b[key])) return false;
    }
    return true;
  }

  return false;
}

/**
 * @param {import("./types.js").Cell | null | undefined} cell
 * @returns {cell is import("./types.js").Cell}
 */
export function isNonEmptyCell(cell) {
  return cell !== null && cell !== undefined && typeof cell === "object";
}

/**
 * `null` is the canonical empty cell in merge/diff code.
 *
 * @param {import("./types.js").Cell | null | undefined} cell
 * @returns {import("./types.js").Cell | null}
 */
export function normalizeCell(cell) {
  if (!isNonEmptyCell(cell)) return null;

  /** @type {import("./types.js").Cell} */
  const normalized = {};

  const enc = /** @type {any} */ (cell).enc;
  // Treat any `enc` marker (including `null`) as encrypted so callers never
  // fall back to plaintext fields when an encryption marker exists.
  if (enc !== undefined) {
    normalized.enc = enc;
  }

  const formula = /** @type {any} */ (cell).formula;
  if (typeof formula === "string" && formula.trim().length > 0) {
    normalized.formula = formula;
  }

  const value = /** @type {any} */ (cell).value;
  // `null` is treated as empty (DocumentController uses null for blank cells).
  if (value !== null && value !== undefined) {
    normalized.value = value;
  }

  const format = /** @type {any} */ (cell).format;
  if (format !== null && format !== undefined) {
    normalized.format = format;
  }

  if (
    normalized.formula === undefined &&
    normalized.value === undefined &&
    normalized.format === undefined &&
    normalized.enc === undefined
  ) {
    return null;
  }

  // Enforce mutual exclusion for formula/value when possible.
  if (normalized.enc !== undefined) {
    delete normalized.value;
    delete normalized.formula;
  } else if (normalized.formula !== undefined) {
    delete normalized.value;
  }

  return normalized;
}

/**
 * Full equality: value/formula + format.
 *
 * @param {import("./types.js").Cell | null | undefined} a
 * @param {import("./types.js").Cell | null | undefined} b
 */
export function cellsEqual(a, b) {
  const ca = normalizeCell(a);
  const cb = normalizeCell(b);
  if (ca === null && cb === null) return true;
  if (ca === null || cb === null) return false;
  return deepEqual(ca, cb);
}

/**
 * Content equality: value/formula only; ignores formatting.
 *
 * @param {import("./types.js").Cell | null | undefined} a
 * @param {import("./types.js").Cell | null | undefined} b
 */
export function cellContentEqual(a, b) {
  const ca = normalizeCell(a);
  const cb = normalizeCell(b);
  const contentA =
    ca === null
      ? null
      : ca.enc !== undefined
        ? { enc: ca.enc }
        : ca.formula !== undefined
          ? { formula: ca.formula }
          : ca.value !== undefined
            ? { value: ca.value }
            : null;
  const contentB =
    cb === null
      ? null
      : cb.enc !== undefined
        ? { enc: cb.enc }
        : cb.formula !== undefined
          ? { formula: cb.formula }
          : cb.value !== undefined
            ? { value: cb.value }
            : null;
  return deepEqual(contentA, contentB);
}

/**
 * Semantic content equivalence:
 * - identical literal values
 * - formulas whose ASTs are equivalent
 *
 * @param {import("./types.js").Cell | null | undefined} a
 * @param {import("./types.js").Cell | null | undefined} b
 */
export function cellContentEquivalent(a, b) {
  const ca = normalizeCell(a);
  const cb = normalizeCell(b);

  const contentA =
    ca === null
      ? null
      : ca.enc !== undefined
        ? { kind: "enc", enc: ca.enc }
      : ca.formula !== undefined
        ? { kind: "formula", formula: ca.formula }
        : ca.value !== undefined
          ? { kind: "value", value: ca.value }
          : null;
  const contentB =
    cb === null
      ? null
      : cb.enc !== undefined
        ? { kind: "enc", enc: cb.enc }
      : cb.formula !== undefined
        ? { kind: "formula", formula: cb.formula }
        : cb.value !== undefined
          ? { kind: "value", value: cb.value }
          : null;

  if (contentA === null && contentB === null) return true;
  if (contentA === null || contentB === null) return false;
  if (contentA.kind !== contentB.kind) return false;

  if (contentA.kind === "formula") {
    return normalizeFormula(contentA.formula) === normalizeFormula(contentB.formula);
  }

  if (contentA.kind === "enc") {
    return deepEqual(contentA.enc, contentB.enc);
  }

  return deepEqual(contentA.value, contentB.value);
}

/**
 * @param {import("./types.js").Cell | null | undefined} cell
 * @returns {import("./types.js").JsonObject | null}
 */
export function cellFormat(cell) {
  const c = normalizeCell(cell);
  if (c === null) return null;
  return c.format ?? null;
}

/**
 * @param {import("./types.js").Cell | null | undefined} cell
 */
export function cellContent(cell) {
  const c = normalizeCell(cell);
  if (c === null) return null;
  if (c.enc !== undefined) {
    return {
      enc: c.enc,
      formula: undefined,
      value: undefined,
    };
  }
  return {
    value: c.value,
    formula: c.formula
  };
}

/**
 * @param {import("./types.js").Cell | null | undefined} base
 * @param {import("./types.js").Cell | null | undefined} next
 */
export function didContentChange(base, next) {
  return !cellContentEqual(base, next);
}

/**
 * @param {import("./types.js").Cell | null | undefined} base
 * @param {import("./types.js").Cell | null | undefined} next
 */
export function didFormatChange(base, next) {
  return !deepEqual(cellFormat(base), cellFormat(next));
}
