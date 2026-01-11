/**
 * @typedef {{
 *   text: string,
 *   runs?: Array<{ start: number, end: number, style?: Record<string, unknown> }>
 * }} RichText
 *
 * @typedef {string | number | boolean | RichText | null} CellValue
 *
 * Canonical cell state owned by the DocumentController.
 *
 * `value` is the literal user-entered value (non-formula). For formula cells, `formula`
 * holds the canonical formula string (including the leading "=").
 *
 * `styleId` references a deduplicated style table owned by the DocumentController.
 *
 * @typedef {{
 *   value: CellValue,
 *   formula: string | null,
 *   styleId: number
 * }} CellState
 */

/**
 * @returns {CellState}
 */
export function emptyCellState() {
  return { value: null, formula: null, styleId: 0 };
}

/**
 * @param {CellState | undefined | null} cell
 * @returns {CellState}
 */
export function normalizeCellState(cell) {
  if (!cell) return emptyCellState();
  return {
    value: cell.value ?? null,
    formula: cell.formula ?? null,
    styleId: typeof cell.styleId === "number" ? cell.styleId : 0,
  };
}

/**
 * Deep-ish clone for history safety.
 *
 * @param {CellState} cell
 * @returns {CellState}
 */
export function cloneCellState(cell) {
  const normalized = normalizeCellState(cell);
  const structuredCloneFn =
    typeof globalThis.structuredClone === "function" ? globalThis.structuredClone : null;
  const clonedValue =
    normalized.value != null && typeof normalized.value === "object"
      ? structuredCloneFn
        ? structuredCloneFn(normalized.value)
        : JSON.parse(JSON.stringify(normalized.value))
      : normalized.value;
  return {
    value: clonedValue,
    formula: normalized.formula,
    styleId: normalized.styleId,
  };
}

/**
 * @param {CellState} cell
 * @returns {boolean}
 */
export function isCellStateEmpty(cell) {
  const normalized = normalizeCellState(cell);
  return normalized.value == null && normalized.formula == null && normalized.styleId === 0;
}

/**
 * @param {CellState} a
 * @param {CellState} b
 * @returns {boolean}
 */
export function cellStateEquals(a, b) {
  const an = normalizeCellState(a);
  const bn = normalizeCellState(b);
  const valuesEqual =
    an.value === bn.value ||
    (an.value != null &&
      bn.value != null &&
      typeof an.value === "object" &&
      typeof bn.value === "object" &&
      JSON.stringify(an.value) === JSON.stringify(bn.value));
  return valuesEqual && an.formula === bn.formula && an.styleId === bn.styleId;
}
