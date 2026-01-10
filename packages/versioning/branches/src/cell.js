import { areFormulasAstEquivalent } from "./formulaAst.js";

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

  if (Object.prototype.hasOwnProperty.call(cell, "formula")) {
    normalized.formula = cell.formula ?? "";
  }

  if (Object.prototype.hasOwnProperty.call(cell, "value")) {
    normalized.value = cell.value;
  }

  if (Object.prototype.hasOwnProperty.call(cell, "format") && cell.format) {
    normalized.format = cell.format;
  }

  if (
    normalized.formula === undefined &&
    normalized.value === undefined &&
    normalized.format === undefined
  ) {
    return null;
  }

  // Enforce mutual exclusion for formula/value when possible.
  if (normalized.formula !== undefined) {
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
  const contentA = ca === null ? null : { value: ca.value, formula: ca.formula };
  const contentB = cb === null ? null : { value: cb.value, formula: cb.formula };
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

  if (ca === null && cb === null) return true;
  if (ca === null || cb === null) return false;

  if (ca.formula !== undefined || cb.formula !== undefined) {
    if (ca.formula === undefined || cb.formula === undefined) return false;
    return areFormulasAstEquivalent(ca.formula, cb.formula);
  }

  return deepEqual(ca.value, cb.value);
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

