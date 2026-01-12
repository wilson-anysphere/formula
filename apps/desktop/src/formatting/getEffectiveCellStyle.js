/**
 * Read the effective (layered) formatting for a single cell.
 *
 * The newer DocumentController API exposes `getCellFormat(sheetId, coord)`, which resolves
 * sheet/row/col/cell formatting layers. Older implementations only store per-cell `styleId`
 * and require a `styleTable` lookup.
 *
 * This helper prefers layered formatting when available, but remains compatible with legacy
 * documents/controllers.
 *
 * @param {any} doc
 * @param {string} sheetId
 * @param {import("../document/coords.js").CellCoord | string} coord
 * @returns {Record<string, any>}
 */
export function getEffectiveCellStyle(doc, sheetId, coord) {
  if (doc && typeof doc.getCellFormat === "function") {
    try {
      const style = doc.getCellFormat(sheetId, coord);
      return style && typeof style === "object" ? style : {};
    } catch {
      // Fall back to legacy behavior below.
    }
  }

  try {
    const cellState = doc?.getCell ? doc.getCell(sheetId, coord) : null;
    const styleId = cellState?.styleId ?? 0;
    const table = doc?.styleTable;
    if (table && typeof table.get === "function") return table.get(styleId) ?? {};
  } catch {
    // ignore
  }

  return {};
}

