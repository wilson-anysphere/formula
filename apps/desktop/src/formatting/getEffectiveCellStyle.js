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
  const id = String(sheetId ?? "").trim();
  if (!id) return {};
  // Avoid materializing "phantom" sheets when callers probe formatting with a stale sheet id.
  // `DocumentController.getCellFormat()` / `getCell()` create sheets lazily when referenced.
  // Prefer side-effect free reads when the controller exposes internal sheet existence checks.
  try {
    const model = doc?.model;
    const sheetMap = model?.sheets;
    const sheetMeta = doc?.sheetMeta;
    const canCheckExists =
      Boolean(sheetMap && typeof sheetMap.has === "function") ||
      Boolean(sheetMeta && typeof sheetMeta.has === "function") ||
      Boolean(doc && typeof doc.getSheetMeta === "function");
    if (canCheckExists) {
      const exists =
        (sheetMap && typeof sheetMap.has === "function" && sheetMap.has(id)) ||
        (sheetMeta && typeof sheetMeta.has === "function" && sheetMeta.has(id)) ||
        (doc && typeof doc.getSheetMeta === "function" && Boolean(doc.getSheetMeta(id)));
      if (!exists) return {};
    }
  } catch {
    // Best-effort: fall back to legacy behavior below.
  }

  if (doc && typeof doc.getCellFormat === "function") {
    try {
      const style = doc.getCellFormat(id, coord);
      return style && typeof style === "object" ? style : {};
    } catch {
      // Fall back to legacy behavior below.
    }
  }

  try {
    // Prefer side-effect free reads when available.
    const cellState = doc?.peekCell ? doc.peekCell(id, coord) : doc?.getCell ? doc.getCell(id, coord) : null;
    const styleId = cellState?.styleId ?? 0;
    const table = doc?.styleTable;
    if (table && typeof table.get === "function") return table.get(styleId) ?? {};
  } catch {
    // ignore
  }

  return {};
}
