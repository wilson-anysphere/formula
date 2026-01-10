import { parseA1 } from "../../document/coords.js";
import { t } from "../../i18n/index.js";
import { parseCsv } from "./csv.js";
import { inferColumnTypes, parseCellWithColumnType } from "./infer.js";

/**
 * @typedef {import("../../document/cell.js").CellState} CellState
 * @typedef {CellState[][]} CellGrid
 *
 * @typedef {{ delimiter?: string, sampleSize?: number }} CsvImportOptions
 *
 * @typedef {{ grid: CellGrid }} CsvImportResult
 *
 * @typedef {import("../../document/documentController.js").DocumentController} DocumentController
 * @typedef {import("../../document/coords.js").CellCoord} CellCoord
 */

/**
 * Parse CSV text into a typed cell grid (for insertion into a sheet).
 *
 * @param {string} csvText
 * @param {CsvImportOptions} [options]
 * @returns {CsvImportResult}
 */
export function importCsvToCellGrid(csvText, options = {}) {
  const delimiter = options.delimiter ?? ",";
  const rows = parseCsv(csvText, { delimiter });

  const columnTypes = inferColumnTypes(rows, options.sampleSize ?? 100);
  const columnCount = columnTypes.length;

  /** @type {CellGrid} */
  const grid = rows.map((row) => {
    return Array.from({ length: columnCount }, (_, col) => {
      const raw = row[col] ?? "";

      if (raw.startsWith("=") && raw.length > 1) {
        return { value: null, formula: raw, format: null };
      }

      const parsed = parseCellWithColumnType(raw, columnTypes[col] ?? "string");
      return { value: parsed.value ?? null, formula: null, format: parsed.format ?? null };
    });
  });

  return { grid };
}

/**
 * Convenience helper to import CSV directly into a document at a start cell.
 *
 * @param {DocumentController} doc
 * @param {string} sheetId
 * @param {CellCoord | string} start
 * @param {string} csvText
 * @param {CsvImportOptions} [options]
 * @returns {CellGrid}
 */
export function importCsvIntoDocument(doc, sheetId, start, csvText, options = {}) {
  const { grid } = importCsvToCellGrid(csvText, options);
  const startCoord = typeof start === "string" ? parseA1(start) : start;

  const values = grid.map((row) =>
    row.map((cell) => {
      if (cell.formula != null) return { formula: cell.formula, format: cell.format };
      return { value: cell.value, format: cell.format };
    })
  );

  doc.setRangeValues(sheetId, startCoord, values, { label: t("importExport.importCsv") });
  return grid;
}
