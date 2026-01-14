import { normalizeRange, parseRangeA1 } from "../../document/coords.js";
import { excelSerialToDate } from "../../shared/valueParsing.js";
import { parseImageCellValue } from "../../shared/imageCellValue.js";
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
 *  maxCells?: number,
 *  dlp?: { documentId: string, classificationStore: any, policy: any }
 * }} CsvExportOptions
 */

// Exporting large selections requires materializing a full 2D cell grid and building a
// potentially huge CSV string. Keep this bounded so Excel-scale sheet limits don't allow
// accidental multi-million-cell exports that would exhaust memory.
const DEFAULT_MAX_EXPORT_CELLS = 200_000;

/**
 * DocumentController creates sheets lazily when referenced by `getCell()` / `getCellFormat()` /
 * `getSheetView()`.
 *
 * Export is a read-only operation and should never recreate a deleted sheet if callers pass a
 * stale id (e.g. during sheet deletion/undo races). We treat a sheet as "known missing" when the
 * workbook already has *some* sheets, but the requested id is present in neither the materialized
 * `model.sheets` map nor the `sheetMeta` map.
 *
 * When the workbook has no sheets yet, we treat the id as "unknown" and preserve the historical
 * lazy-creation behavior.
 *
 * @param {any} doc
 * @param {string} sheetId
 */
function isSheetKnownMissing(doc, sheetId) {
  const id = String(sheetId ?? "").trim();
  if (!id) return true;

  const sheets = doc?.model?.sheets;
  const sheetMeta = doc?.sheetMeta;
  if (
    sheets &&
    typeof sheets.has === "function" &&
    typeof sheets.size === "number" &&
    sheetMeta &&
    typeof sheetMeta.has === "function" &&
    typeof sheetMeta.size === "number"
  ) {
    const workbookHasAnySheets = sheets.size > 0 || sheetMeta.size > 0;
    if (!workbookHasAnySheets) return false;
    return !sheets.has(id) && !sheetMeta.has(id);
  }
  return false;
}

function isLikelyDateNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  const lower = fmt.toLowerCase();
  return lower.includes("yyyy-mm-dd") || lower.includes("m/d/yyyy");
}

function isLikelyTimeNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  const compact = fmt.toLowerCase().replace(/\s+/g, "");
  return /^h{1,2}:m{1,2}(:s{1,2})?$/.test(compact);
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function formatExcelTime(serial, fmt) {
  const date = excelSerialToDate(serial);
  const hh = date.getUTCHours();
  const mm = date.getUTCMinutes();
  const ss = date.getUTCSeconds();
  const compact = String(fmt).toLowerCase().replace(/\s+/g, "");
  const hasSeconds = compact.includes(":s");
  return hasSeconds ? `${pad2(hh)}:${pad2(mm)}:${pad2(ss)}` : `${pad2(hh)}:${pad2(mm)}`;
}

/**
 * @param {CellState} cell
 * @returns {string}
 */
function cellToCsvField(cell) {
  const value = cell.value;
  if (value == null) {
    const formula = cell.formula;
    if (typeof formula === "string" && formula.trim() !== "") {
      // When we don't have a cached/display value, fall back to exporting the formula text.
      return formula;
    }
    return "";
  }

  // DocumentController stores rich text as `{ text, runs }`. CSV exports should serialize
  // this as plain text rather than `[object Object]`.
  if (typeof value === "object" && typeof value.text === "string") {
    const text = value.text;
    // Escape so CSV import doesn't treat literal rich text (e.g. "=literal") as a formula.
    if (text.trimStart().startsWith("=") || text.startsWith("'")) return `'${text}`;
    return text;
  }

  const image = parseImageCellValue(value);
  if (image) {
    const formula = cell.formula;
    if (typeof formula === "string" && formula.trim() !== "") {
      // Mirror the "formula fallback" behavior above, even though image payloads are represented
      // as objects rather than the normal `value=null` formula convention.
      return formula;
    }

    const text = image.altText ?? "[Image]";
    if (text.trimStart().startsWith("=") || text.startsWith("'")) return `'${text}`;
    return text;
  }

  const numberFormat = cell.format?.numberFormat;
  if (typeof value === "number" && isLikelyDateNumberFormat(numberFormat)) {
    const date = excelSerialToDate(value);
    const lower = numberFormat.toLowerCase();
    return lower.includes("h") ? date.toISOString() : date.toISOString().slice(0, 10);
  }

  if (typeof value === "number" && isLikelyTimeNumberFormat(numberFormat)) {
    return formatExcelTime(value, numberFormat);
  }

  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return String(value);
  if (typeof value === "string" && (value.trimStart().startsWith("=") || value.startsWith("'"))) {
    // Escape CSV so `importCsvToCellGrid` doesn't accidentally treat literal strings (e.g. "=literal")
    // as formulas when parsing back into a cell grid.
    return `'${value}`;
  }
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
  const sheetKey = String(sheetId ?? "").trim();
  if (!sheetKey) {
    throw new Error("Sheet id is required.");
  }
  if (isSheetKnownMissing(doc, sheetKey)) {
    throw new Error(`Unknown sheet: ${sheetKey}`);
  }

  if (options?.dlp) {
    enforceExport({
      documentId: options.dlp.documentId,
      sheetId: sheetKey,
      range: r,
      format: "csv",
      classificationStore: options.dlp.classificationStore,
      policy: options.dlp.policy,
    });
  }

  const rowCount = Math.max(0, r.end.row - r.start.row + 1);
  const colCount = Math.max(0, r.end.col - r.start.col + 1);
  const cellCount = rowCount * colCount;

  const maxCells = (() => {
    const raw = options?.maxCells;
    if (raw === Infinity) return Infinity;
    const n = Number(raw);
    if (Number.isFinite(n) && Number.isInteger(n) && n > 0) return n;
    return DEFAULT_MAX_EXPORT_CELLS;
  })();

  if (Number.isFinite(maxCells) && cellCount > maxCells) {
    throw new Error(
      `Selection too large to export (${cellCount.toLocaleString()} cells; max=${maxCells.toLocaleString()}). Select fewer cells and try again.`
    );
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
      const cell = doc.getCell(sheetKey, { row, col });
      let format = null;
      if (hasStyleIdTuple) {
        const tuple = doc.getCellFormatStyleIds(sheetKey, { row, col });
        if (Array.isArray(tuple) && tuple.length >= 4) {
          const key = tuple.join(",");
          if (formatCache.has(key)) {
            format = formatCache.get(key) ?? null;
          } else {
            const hasAnyFormat = tuple.some((id) => id !== 0);
            format = hasAnyFormat ? doc.getCellFormat(sheetKey, { row, col }) : null;
            formatCache.set(key, format);
          }
        } else {
          format = doc.getCellFormat(sheetKey, { row, col });
        }
      } else {
        format = doc.getCellFormat(sheetKey, { row, col });
      }
      // Export paths expect `cell.format` (not `styleId`) so downstream serialization can
      // respect number formats (including layered formats inherited from row/col/sheet defaults).
      outRow.push({ ...cell, format });
    }
    grid.push(outRow);
  }

  return exportCellGridToCsv(grid, options);
}
