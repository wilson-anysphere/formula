import type { CellProvider, CellRange } from "../model/CellProvider.ts";
import { formatCellDisplayText } from "../rendering/CanvasGridRenderer.ts";

export { formatCellDisplayText } from "../rendering/CanvasGridRenderer.ts";

/**
 * Convert a 0-based column index to an Excel-style column name (A, B, ..., Z, AA, AB, ...).
 */
export function toColumnName(col0: number): string {
  let value = col0 + 1;
  let name = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    value = Math.floor((value - 1) / 26);
  }
  return name;
}

/**
 * Convert 0-based row/column indexes to an Excel-style A1 address.
 */
export function toA1Address(row0: number, col0: number): string {
  return `${toColumnName(col0)}${row0 + 1}`;
}

export function describeCell(
  selection: { row: number; col: number } | null,
  range: CellRange | null,
  provider: CellProvider,
  headerRows: number,
  headerCols: number
): string {
  if (!selection) return "No cell selected.";

  const row0 = selection.row - headerRows;
  const col0 = selection.col - headerCols;
  const address =
    row0 >= 0 && col0 >= 0 ? toA1Address(row0, col0) : `row ${selection.row + 1}, column ${selection.col + 1}`;

  const cell = provider.getCell(selection.row, selection.col);
  let valueText = formatCellDisplayText(cell?.value ?? null);
  if (valueText.trim() === "" && cell?.image) {
    valueText = cell.image.altText?.trim() ? cell.image.altText : "[Image]";
  }
  const valueDescription = valueText.trim() === "" ? "blank" : valueText;

  let selectionDescription = "none";
  if (range) {
    const startRow0 = range.startRow - headerRows;
    const startCol0 = range.startCol - headerCols;
    const endRow0 = range.endRow - headerRows - 1;
    const endCol0 = range.endCol - headerCols - 1;
    if (startRow0 >= 0 && startCol0 >= 0 && endRow0 >= 0 && endCol0 >= 0) {
      const start = toA1Address(startRow0, startCol0);
      const end = toA1Address(endRow0, endCol0);
      selectionDescription = start === end ? start : `${start}:${end}`;
    } else {
      selectionDescription = `row ${range.startRow + 1}, column ${range.startCol + 1}`;
    }
  }

  return `Active cell ${address}, value ${valueDescription}. Selection ${selectionDescription}.`;
}

export { SR_ONLY_STYLE, applySrOnlyStyle } from "./srOnlyStyle";
