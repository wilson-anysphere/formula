function cellKey(row, col) {
  return `${row},${col}`;
}

class InMemorySpreadsheet {
  constructor() {
    this._cells = new Map();
    this._selection = { startRow: 0, startCol: 0, endRow: 0, endCol: 0 };
    this._selectionListeners = new Set();
    this._cellListeners = new Set();
  }

  setSelection(range) {
    this._selection = {
      startRow: range.startRow,
      startCol: range.startCol,
      endRow: range.endRow,
      endCol: range.endCol
    };

    const payload = { selection: this.getSelection() };
    for (const listener of this._selectionListeners) listener(payload);
  }

  getSelection() {
    const { startRow, startCol, endRow, endCol } = this._selection;
    return {
      startRow,
      startCol,
      endRow,
      endCol,
      values: this.getRangeValues(startRow, startCol, endRow, endCol)
    };
  }

  onSelectionChanged(callback) {
    this._selectionListeners.add(callback);
    return { dispose: () => this._selectionListeners.delete(callback) };
  }

  getCell(row, col) {
    const key = cellKey(row, col);
    return this._cells.has(key) ? this._cells.get(key) : null;
  }

  setCell(row, col, value) {
    this._cells.set(cellKey(row, col), value);
    const payload = { row, col, value };
    for (const listener of this._cellListeners) listener(payload);
  }

  onCellChanged(callback) {
    this._cellListeners.add(callback);
    return { dispose: () => this._cellListeners.delete(callback) };
  }

  getRangeValues(startRow, startCol, endRow, endCol) {
    const rows = [];
    for (let r = startRow; r <= endRow; r++) {
      const cols = [];
      for (let c = startCol; c <= endCol; c++) {
        cols.push(this.getCell(r, c));
      }
      rows.push(cols);
    }
    return rows;
  }
}

module.exports = {
  InMemorySpreadsheet
};
