/**
 * @param {any} raw
 * @returns {{ v?: any, f?: string }}
 */
export function normalizeCell(raw) {
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    const hasVF =
      Object.prototype.hasOwnProperty.call(raw, "v") || Object.prototype.hasOwnProperty.call(raw, "f");
    if (hasVF) {
      return raw;
    }

    const hasValueFormula =
      Object.prototype.hasOwnProperty.call(raw, "value") || Object.prototype.hasOwnProperty.call(raw, "formula");
    if (hasValueFormula) {
      const value = raw.value ?? null;
      const formulaRaw = raw.formula ?? null;
      const formula =
        typeof formulaRaw === "string" && formulaRaw.trim() !== ""
          ? formulaRaw.trim().startsWith("=")
            ? formulaRaw.trim()
            : `=${formulaRaw.trim()}`
          : null;

      if (formula == null && (value == null || value === "")) return {};
      /** @type {{ v?: any, f?: string }} */
      const out = {};
      if (value != null && value !== "") out.v = value;
      if (formula != null) out.f = formula;
      return out;
    }

    // Treat `{}` as an empty cell; it's a common sparse representation.
    if (raw.constructor === Object && Object.keys(raw).length === 0) return {};

    // Preserve rich object values (e.g. Date, structured types) as the cell value.
    if (raw instanceof Date) return { v: raw };
  }

  if (typeof raw === "string") {
    const trimmed = raw.trim();
    if (!trimmed) return {};
    if (trimmed.startsWith("=")) return { f: trimmed };
    return { v: trimmed };
  }

  if (raw == null) return {};
  return { v: raw };
}

/**
 * @param {any} sheet
 * @returns {any[][] | null}
 */
export function getSheetMatrix(sheet) {
  if (Array.isArray(sheet?.cells)) return sheet.cells;
  if (Array.isArray(sheet?.values)) return sheet.values;
  return null;
}

/**
 * @param {any} sheet
 * @returns {Map<string, any> | null}
 */
export function getSheetCellMap(sheet) {
  const cells = sheet?.cells;
  if (cells instanceof Map) return cells;
  return null;
}

/**
 * @param {any} sheet
 * @param {number} row
 * @param {number} col
 */
export function getCellRaw(sheet, row, col) {
  const matrix = getSheetMatrix(sheet);
  if (matrix) return matrix[row]?.[col];
  const map = getSheetCellMap(sheet);
  if (map) return map.get(`${row},${col}`) ?? map.get(`${row}:${col}`) ?? null;
  if (typeof sheet?.getCell === "function") return sheet.getCell(row, col);
  return null;
}
