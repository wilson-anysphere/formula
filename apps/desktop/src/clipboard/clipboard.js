import { normalizeRange, parseA1, parseRangeA1 } from "../document/coords.js";
import { parseHtmlToCellGrid, serializeCellGridToHtml } from "./html.js";
import { parseTsvToCellGrid, serializeCellGridToTsv } from "./tsv.js";

/**
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("./types.js").ClipboardContent} ClipboardContent
 * @typedef {import("./types.js").ClipboardWritePayload} ClipboardWritePayload
 * @typedef {import("./types.js").PasteOptions} PasteOptions
 *
 * @typedef {import("../document/coords.js").CellCoord} CellCoord
 * @typedef {import("../document/coords.js").CellRange} CellRange
 * @typedef {import("../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../document/cell.js").CellState} CellState
 */

/**
 * @param {CellGrid} grid
 * @returns {ClipboardWritePayload}
 */
export function serializeCellGridToClipboardPayload(grid) {
  return {
    text: serializeCellGridToTsv(grid),
    html: serializeCellGridToHtml(grid),
  };
}

/**
 * Parse clipboard payloads in priority order: HTML > TSV/plain.
 * @param {ClipboardContent} content
 * @returns {CellGrid | null}
 */
export function parseClipboardContentToCellGrid(content) {
  const html = content.html?.trim();
  if (html) {
    const parsed = parseHtmlToCellGrid(html);
    if (parsed) return parsed;
  }

  const text = content.text;
  if (typeof text === "string") return parseTsvToCellGrid(text);

  return null;
}

/**
 * Extract a rectangular cell grid from the document controller.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellRange | string} range
 * @returns {CellGrid}
 */
export function getCellGridFromRange(doc, sheetId, range) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);

  /** @type {CellGrid} */
  const grid = [];

  for (let row = r.start.row; row <= r.end.row; row++) {
    /** @type {CellState[]} */
    const outRow = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      outRow.push(doc.getCell(sheetId, { row, col }));
    }
    grid.push(outRow);
  }

  return grid;
}

/**
 * Copy a document range into an Excel-compatible clipboard payload (TSV + HTML table).
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellRange | string} range
 * @returns {ClipboardWritePayload}
 */
export function copyRangeToClipboardPayload(doc, sheetId, range) {
  const grid = getCellGridFromRange(doc, sheetId, range);
  return serializeCellGridToClipboardPayload(grid);
}

/**
 * Paste clipboard content into the document controller at a start cell.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellCoord | string} start
 * @param {ClipboardContent} content
 * @param {PasteOptions} [options]
 * @returns {boolean} whether a paste occurred
 */
export function pasteClipboardContent(doc, sheetId, start, content, options = {}) {
  const grid = parseClipboardContentToCellGrid(content);
  if (!grid) return false;

  const mode = options.mode ?? "all";

  const values = grid.map((row) =>
    row.map((cell) => {
      if (mode === "values") return cell.value ?? null;
      if (mode === "formulas") return cell.formula != null ? { formula: cell.formula } : cell.value ?? null;
      if (mode === "formats") return { format: cell.format ?? null };

      // mode === "all"
      if (cell.formula != null) return { formula: cell.formula, format: cell.format ?? null };
      return { value: cell.value ?? null, format: cell.format ?? null };
    })
  );

  const startCoord = typeof start === "string" ? parseA1(start) : start;
  doc.setRangeValues(sheetId, startCoord, values, { label: "Paste" });
  return true;
}
