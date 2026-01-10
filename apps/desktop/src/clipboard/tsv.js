import { excelSerialToDate, parseScalar } from "../shared/valueParsing.js";

/**
 * @typedef {import("./types.js").CellGrid} CellGrid
 * @typedef {import("../document/cell.js").CellState} CellState
 */

function isLikelyDateNumberFormat(fmt) {
  if (typeof fmt !== "string") return false;
  return fmt.toLowerCase().includes("yyyy-mm-dd");
}

/**
 * @param {CellState} cell
 * @returns {string}
 */
function cellValueToPlainText(cell) {
  const value = cell.value;
  if (value == null) return "";

  const numberFormat = cell.format?.numberFormat;
  if (typeof value === "number" && isLikelyDateNumberFormat(numberFormat)) {
    const date = excelSerialToDate(value);
    return numberFormat.includes("hh") ? date.toISOString() : date.toISOString().slice(0, 10);
  }

  return String(value);
}

/**
 * @param {CellGrid} grid
 * @returns {string}
 */
export function serializeCellGridToTsv(grid) {
  return grid
    .map((row) => row.map(cellValueToPlainText).join("\t"))
    .join("\n");
}

/**
 * @param {string} tsv
 * @returns {CellGrid}
 */
export function parseTsvToCellGrid(tsv) {
  const normalized = tsv.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = normalized.split("\n");

  // Drop the final empty record when the clipboard payload ends with a newline.
  if (lines.length > 1 && lines.at(-1) === "") lines.pop();

  return lines.map((line) => {
    const parts = line.split("\t");
    return parts.map((raw) => {
      if (raw.startsWith("=") && raw.length > 1) {
        return { value: null, formula: raw, format: null };
      }

      const parsed = parseScalar(raw);
      const format = parsed.type === "datetime" ? { numberFormat: parsed.numberFormat } : null;
      return { value: parsed.value, formula: null, format };
    });
  });
}
