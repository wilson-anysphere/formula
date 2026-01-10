import { cloneCellState, emptyCellState, isCellStateEmpty } from "./cell.js";

/**
 * @typedef {import("./cell.js").CellState} CellState
 * @typedef {import("./coords.js").CellCoord} CellCoord
 */

/**
 * A minimal interface the DocumentController uses to keep the calc engine in sync.
 *
 * A future engine implementation can interpret `value`/`formula` however it likes.
 *
 * @typedef {{
 *   sheetId: string,
 *   row: number,
 *   col: number,
 *   cell: CellState
 * }} CellChange
 */

/**
 * @typedef {{
 *   applyChanges: (changes: readonly CellChange[]) => void,
 *   recalculate: () => void,
 *   beginBatch?: () => void,
 *   endBatch?: () => void,
 * }} Engine
 */

function cellKey(row, col) {
  return `${row},${col}`;
}

class MockSheet {
  constructor() {
    /** @type {Map<string, CellState>} */
    this.cells = new Map();
  }

  /**
   * @param {number} row
   * @param {number} col
   * @returns {CellState}
   */
  getCell(row, col) {
    return cloneCellState(this.cells.get(cellKey(row, col)) ?? emptyCellState());
  }

  /**
   * @param {number} row
   * @param {number} col
   * @param {CellState} cell
   */
  setCell(row, col, cell) {
    if (isCellStateEmpty(cell)) {
      this.cells.delete(cellKey(row, col));
      return;
    }
    this.cells.set(cellKey(row, col), cloneCellState(cell));
  }
}

/**
 * Simple in-memory engine used by tests.
 */
export class MockEngine {
  constructor() {
    /** @type {Map<string, MockSheet>} */
    this.sheets = new Map();
    this.recalcCount = 0;
    this.batchDepth = 0;
    /** @type {CellChange[]} */
    this.appliedChanges = [];
  }

  /**
   * @param {string} sheetId
   * @returns {MockSheet}
   */
  #sheet(sheetId) {
    let sheet = this.sheets.get(sheetId);
    if (!sheet) {
      sheet = new MockSheet();
      this.sheets.set(sheetId, sheet);
    }
    return sheet;
  }

  /**
   * @param {readonly CellChange[]} changes
   */
  applyChanges(changes) {
    for (const change of changes) {
      this.#sheet(change.sheetId).setCell(change.row, change.col, change.cell);
      this.appliedChanges.push({ ...change, cell: cloneCellState(change.cell) });
    }
  }

  recalculate() {
    this.recalcCount += 1;
  }

  beginBatch() {
    this.batchDepth += 1;
  }

  endBatch() {
    this.batchDepth = Math.max(0, this.batchDepth - 1);
  }

  /**
   * @param {string} sheetId
   * @param {number} row
   * @param {number} col
   * @returns {CellState}
   */
  getCell(sheetId, row, col) {
    return this.#sheet(sheetId).getCell(row, col);
  }
}

