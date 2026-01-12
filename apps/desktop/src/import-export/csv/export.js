import { normalizeRange, parseRangeA1 } from "../../document/coords.js";
import { excelSerialToDate } from "../../shared/valueParsing.js";
import { stringifyCsv } from "./csv.js";
import { enforceExport } from "../../dlp/enforceExport.js";

/**
 * @typedef {import("../../document/cell.js").CellState} CellState
 * @typedef {CellState[][]} CellGrid
 * @typedef {import("../../document/coords.js").CellRange} CellRange
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 *
 * @typedef {{
 *  delimiter?: string,
 *  newline?: "\n" | "\r\n",
 *  dlp?: { documentId: string, classificationStore: any, policy: any }
 * }} CsvExportOptions
 */

function isLikelyDateNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  return fmt.toLowerCase().includes("yyyy-mm-dd");
}

/**
 * @param {CellState} cell
 * @returns {string}
 */
function cellToCsvField(cell) {
  const value = cell.value;
  if (value == null) return "";

  // DocumentController stores rich text as `{ text, runs }`. CSV exports should serialize
  // this as plain text rather than `[object Object]`.
  if (typeof value === "object" && typeof value.text === "string") {
    return value.text;
  }

  const numberFormat = cell.format?.numberFormat;
  if (typeof value === "number" && isLikelyDateNumberFormat(numberFormat)) {
    const date = excelSerialToDate(value);
    return numberFormat.includes("hh") ? date.toISOString() : date.toISOString().slice(0, 10);
  }

  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return String(value);
  return String(value);
}

/**
 * Export a cell grid to CSV.
 *
 * @param {CellGrid} grid
 * @param {CsvExportOptions} [options]
 * @returns {string}
 */
export function exportCellGridToCsv(grid, options = {}) {
  const delimiter = options.delimiter ?? ",";
  const rows = grid.map((row) => row.map(cellToCsvField));
  return stringifyCsv(rows, { delimiter, newline: options.newline });
}

/**
 * Export a document range to CSV.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellRange | string} range
 * @param {CsvExportOptions} [options]
 * @returns {string}
 */
export function exportDocumentRangeToCsv(doc, sheetId, range, options = {}) {
  const r = typeof range === "string" ? parseRangeA1(range) : normalizeRange(range);

  if (options?.dlp) {
    enforceExport({
      documentId: options.dlp.documentId,
      sheetId,
      range: r,
      format: "csv",
      classificationStore: options.dlp.classificationStore,
      policy: options.dlp.policy,
    });
  }

  /** @type {CellGrid} */
  const grid = [];
  /** @type {Map<string, any>} */
  const formatCache = new Map();
  const hasStyleIdTuple = typeof doc.getCellFormatStyleIds === "function";

  for (let row = r.start.row; row <= r.end.row; row++) {
    /** @type {CellState[]} */
    const outRow = [];
    for (let col = r.start.col; col <= r.end.col; col++) {
      const cell = doc.getCell(sheetId, { row, col });
      let format = null;
      if (hasStyleIdTuple) {
        const tuple = doc.getCellFormatStyleIds(sheetId, { row, col });
        if (Array.isArray(tuple) && tuple.length >= 4) {
          const key = tuple.join(",");
          if (formatCache.has(key)) {
            format = formatCache.get(key) ?? null;
          } else {
            const hasAnyFormat = tuple.some((id) => id !== 0);
            format = hasAnyFormat ? doc.getCellFormat(sheetId, { row, col }) : null;
            formatCache.set(key, format);
          }
        } else {
          format = doc.getCellFormat(sheetId, { row, col });
        }
      } else {
        format = doc.getCellFormat(sheetId, { row, col });
      }
      // Export paths expect `cell.format` (not `styleId`) so downstream serialization can
      // respect number formats (including layered formats inherited from row/col/sheet defaults).
      outRow.push({ ...cell, format });
    }
    grid.push(outRow);
  }

  return exportCellGridToCsv(grid, options);
}
