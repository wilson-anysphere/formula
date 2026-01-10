/**
 * Minimal "Format Cells" implementation.
 *
 * The real UI will be a modal dialog; this module focuses on the style mutation
 * semantics so they can be unit tested and wired to shortcuts.
 */

/**
 * @param {import("../document/documentController.js").DocumentController} doc
 * @param {string} sheetId
 * @param {string | import("../document/coords.js").CellRange} range
 * @param {Record<string, any>} changes
 */
export function applyFormatCells(doc, sheetId, range, changes) {
  doc.setRangeFormat(sheetId, range, changes, { label: "Format Cells" });
}

