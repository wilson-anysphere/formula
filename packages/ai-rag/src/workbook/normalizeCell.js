/**
 * @param {any} raw
 * @returns {{ v?: any, f?: string }}
 */
export function normalizeCell(raw) {
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    if (Object.prototype.hasOwnProperty.call(raw, "v") || Object.prototype.hasOwnProperty.call(raw, "f")) {
      return raw;
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
 * @returns {any[][]}
 */
export function getSheetMatrix(sheet) {
  return sheet?.cells ?? sheet?.values ?? [];
}
