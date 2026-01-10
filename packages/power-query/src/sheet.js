import { DataTable } from "./table.js";

/**
 * Very small "sheet" abstraction for tests / integration points.
 * The real product should implement this against the spreadsheet's data model.
 */
export class InMemorySheet {
  constructor() {
    /** @type {Map<string, unknown>} */
    this.cells = new Map();
  }

  /**
   * @param {number} row 1-based
   * @param {number} col 1-based
   * @param {unknown} value
   */
  setCell(row, col, value) {
    this.cells.set(`${row},${col}`, value);
  }

  /**
   * @param {number} row 1-based
   * @param {number} col 1-based
   * @returns {unknown}
   */
  getCell(row, col) {
    return this.cells.get(`${row},${col}`);
  }
}

/**
 * Write a table into a sheet-like interface.
 * @param {DataTable} table
 * @param {{ setCell: (row: number, col: number, value: unknown) => void }} sheet
 * @param {{ startRow?: number, startCol?: number }} [options]
 */
export function writeTableToSheet(table, sheet, options = {}) {
  const startRow = options.startRow ?? 1;
  const startCol = options.startCol ?? 1;
  const grid = table.toGrid({ includeHeader: true });

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r];
    for (let c = 0; c < row.length; c++) {
      sheet.setCell(startRow + r, startCol + c, row[c]);
    }
  }
}

