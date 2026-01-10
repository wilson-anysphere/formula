export function normalizeName(name) {
  return String(name).trim().toUpperCase();
}

function cellKey(row, col) {
  return `${row},${col}`;
}

function parseCellKey(key) {
  const [r, c] = key.split(",");
  return { row: Number.parseInt(r, 10), col: Number.parseInt(c, 10) };
}

function isCellEmpty(cell) {
  if (cell == null) return true;
  if (cell.formula != null && cell.formula !== "") return false;
  if (cell.value != null && cell.value !== "") return false;
  if (cell.display != null && cell.display !== "") return false;
  return true;
}

export class InMemorySheet {
  constructor(name) {
    this.name = name;
    this._cells = new Map();
  }

  getCell(row, col) {
    return this._cells.get(cellKey(row, col)) ?? null;
  }

  /**
   * Set/replace a cell. Passing `null` clears the cell.
   */
  setCell(row, col, cell) {
    const key = cellKey(row, col);
    if (cell == null || isCellEmpty(cell)) {
      this._cells.delete(key);
      return;
    }
    this._cells.set(key, { ...cell });
  }

  setValue(row, col, value, { display } = {}) {
    this.setCell(row, col, { value, display });
  }

  setFormula(row, col, formula, { value, display } = {}) {
    this.setCell(row, col, { formula, value, display });
  }

  getUsedRange() {
    if (this._cells.size === 0) return null;

    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -Infinity;
    let maxCol = -Infinity;

    for (const key of this._cells.keys()) {
      const { row, col } = parseCellKey(key);
      if (row < minRow) minRow = row;
      if (col < minCol) minCol = col;
      if (row > maxRow) maxRow = row;
      if (col > maxCol) maxCol = col;
    }

    return { startRow: minRow, endRow: maxRow, startCol: minCol, endCol: maxCol };
  }

  *iterateCells(range, { order = "byRows" } = {}) {
    const results = [];

    for (const [key, cell] of this._cells.entries()) {
      const { row, col } = parseCellKey(key);
      if (
        row >= range.startRow &&
        row <= range.endRow &&
        col >= range.startCol &&
        col <= range.endCol
      ) {
        results.push({ row, col, cell });
      }
    }

    results.sort((a, b) => {
      if (order === "byColumns") {
        if (a.col !== b.col) return a.col - b.col;
        return a.row - b.row;
      }
      // byRows
      if (a.row !== b.row) return a.row - b.row;
      return a.col - b.col;
    });

    for (const entry of results) {
      yield entry;
    }
  }
}

export class InMemoryWorkbook {
  constructor() {
    this.sheets = [];
    this._sheetsByName = new Map();
    this.names = new Map();
    this.tables = new Map();
  }

  addSheet(name) {
    const sheet = new InMemorySheet(name);
    this.sheets.push(sheet);
    this._sheetsByName.set(normalizeName(name), sheet);
    return sheet;
  }

  getSheet(name) {
    const sheet = this._sheetsByName.get(normalizeName(name));
    if (!sheet) throw new Error(`Unknown sheet: ${name}`);
    return sheet;
  }

  defineName(name, ref) {
    this.names.set(normalizeName(name), ref);
  }

  getName(name) {
    return this.names.get(normalizeName(name)) ?? null;
  }

  addTable(table) {
    this.tables.set(normalizeName(table.name), table);
  }

  getTable(name) {
    return this.tables.get(normalizeName(name)) ?? null;
  }
}
