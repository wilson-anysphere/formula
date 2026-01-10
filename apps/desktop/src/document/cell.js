/**
 * @typedef {{
 *   text: string,
 *   runs?: Array<{ start: number, end: number, style?: Record<string, unknown> }>
 * }} RichText
 *
 * @typedef {string | number | boolean | RichText | null} CellValue
 *
 * @typedef {Record<string, unknown>} CellFormat
 *
 * Canonical cell state owned by the DocumentController.
 *
 * `value` is the literal user-entered value (non-formula). For formula cells, `formula`
 * holds the raw formula string (without leading "=" normalization).
 *
 * `format` is an optional style bag. It is intentionally un-opinionated so that a future
 * style system can grow without changing the undo/redo substrate.
 *
 * @typedef {{
 *   value: CellValue,
 *   formula: string | null,
 *   format: CellFormat | null
 * }} CellState
 */

/**
 * @returns {CellState}
 */
export function emptyCellState() {
  return { value: null, formula: null, format: null };
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
    format: cell.format ?? null,
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
    format:
      normalized.format == null
        ? null
        : structuredCloneFn
          ? structuredCloneFn(normalized.format)
          : JSON.parse(JSON.stringify(normalized.format)),
  };
}

/**
 * @param {CellState} cell
 * @returns {boolean}
 */
export function isCellStateEmpty(cell) {
  const normalized = normalizeCellState(cell);
  return normalized.value == null && normalized.formula == null && normalized.format == null;
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
  return (
    valuesEqual &&
    an.formula === bn.formula &&
    JSON.stringify(an.format) === JSON.stringify(bn.format)
  );
}
