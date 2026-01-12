function cellKey(row, col) {
  return `${row},${col}`;
}

// Extension APIs represent ranges as full 2D JS arrays. With Excel-scale sheets, unbounded
// ranges can allocate millions of entries and OOM the host process. Keep reads/writes bounded
// to match the desktop extension API guardrails.
const DEFAULT_EXTENSION_RANGE_CELL_LIMIT = 200000;

function normalizeRangeCoords(startRow, startCol, endRow, endCol) {
  const sRow = Number(startRow);
  const sCol = Number(startCol);
  const eRow = Number(endRow);
  const eCol = Number(endCol);
  return {
    startRow: Math.min(sRow, eRow),
    startCol: Math.min(sCol, eCol),
    endRow: Math.max(sRow, eRow),
    endCol: Math.max(sCol, eCol)
  };
}

function getRangeSize(startRow, startCol, endRow, endCol) {
  const r = normalizeRangeCoords(startRow, startCol, endRow, endCol);
  const rows = Math.max(0, r.endRow - r.startRow + 1);
  const cols = Math.max(0, r.endCol - r.startCol + 1);
  return { ...r, rows, cols, cellCount: rows * cols };
}

function assertRangeWithinLimit(startRow, startCol, endRow, endCol, { maxCells, label } = {}) {
  const size = getRangeSize(startRow, startCol, endRow, endCol);
  const limit = Number.isFinite(maxCells) ? maxCells : DEFAULT_EXTENSION_RANGE_CELL_LIMIT;
  if (size.cellCount > limit) {
    const name = String(label ?? "Range");
    throw new Error(
      `${name} is too large (${size.rows}x${size.cols}=${size.cellCount} cells). Limit is ${limit} cells.`
    );
  }
  return size;
}

function columnLettersToIndex(letters) {
  const cleaned = String(letters ?? "").trim().toUpperCase();
  if (!/^[A-Z]+$/.test(cleaned)) {
    throw new Error(`Invalid column letters: ${letters}`);
  }

  let index = 0;
  for (const ch of cleaned) {
    index = index * 26 + (ch.charCodeAt(0) - 64); // A=1
  }
  return index - 1; // 0-based
}

function parseA1CellRef(ref) {
  const match = /^\s*\$?([A-Za-z]+)\$?(\d+)\s*$/.exec(String(ref ?? ""));
  if (!match) throw new Error(`Invalid A1 cell reference: ${ref}`);
  const [, colLetters, rowDigits] = match;
  const row = Number.parseInt(rowDigits, 10) - 1; // 0-based
  const col = columnLettersToIndex(colLetters);
  if (!Number.isFinite(row) || row < 0) throw new Error(`Invalid row in A1 reference: ${ref}`);
  return { row, col };
}

function parseSheetPrefix(raw) {
  const str = String(raw ?? "");
  const bang = str.indexOf("!");
  if (bang === -1) return { sheetName: null, a1Ref: str };
  const sheetPart = str.slice(0, bang).trim();
  const rest = str.slice(bang + 1);
  if (sheetPart.length === 0) throw new Error(`Invalid sheet-qualified reference: ${raw}`);
  if (sheetPart.startsWith("'") && sheetPart.endsWith("'") && sheetPart.length >= 2) {
    const unquoted = sheetPart.slice(1, -1).replace(/''/g, "'");
    return { sheetName: unquoted, a1Ref: rest };
  }
  return { sheetName: sheetPart, a1Ref: rest };
}

function parseA1RangeRef(ref) {
  const { sheetName, a1Ref } = parseSheetPrefix(ref);
  const parts = String(a1Ref ?? "").split(":");
  if (parts.length > 2) throw new Error(`Invalid A1 range reference: ${ref}`);
  const start = parseA1CellRef(parts[0]);
  const end = parts.length === 2 ? parseA1CellRef(parts[1]) : start;
  return {
    sheetName,
    startRow: Math.min(start.row, end.row),
    startCol: Math.min(start.col, end.col),
    endRow: Math.max(start.row, end.row),
    endCol: Math.max(start.col, end.col)
  };
}

class InMemorySpreadsheet {
  constructor() {
    this._sheets = [
      {
        id: "sheet1",
        name: "Sheet1",
        cells: new Map(),
        selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }
      }
    ];
    this._nextSheetId = 2;
    this._activeSheetId = "sheet1";
    this._selectionListeners = new Set();
    this._cellListeners = new Set();
    this._sheetListeners = new Set();
  }

  _getActiveSheetRecord() {
    const sheet = this._sheets.find((s) => s.id === this._activeSheetId);
    if (!sheet) {
      // Should never happen, but keep behavior predictable for tests.
      throw new Error(`Active sheet not found: ${this._activeSheetId}`);
    }
    return sheet;
  }

  getActiveSheet() {
    const sheet = this._getActiveSheetRecord();
    return { id: sheet.id, name: sheet.name };
  }

  listSheets() {
    return this._sheets.map((sheet) => ({ id: sheet.id, name: sheet.name }));
  }

  _getSheetRecordByName(name) {
    const sheetName = String(name);
    const sheet = this._sheets.find((s) => s.name === sheetName);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
    return sheet;
  }

  getSheet(name) {
    const sheet = this._sheets.find((s) => s.name === String(name));
    if (!sheet) return undefined;
    return { id: sheet.id, name: sheet.name };
  }

  createSheet(name) {
    const sheetName = String(name);
    if (sheetName.trim().length === 0) throw new Error("Sheet name must be a non-empty string");
    if (this._sheets.some((s) => s.name === sheetName)) {
      throw new Error(`Sheet already exists: ${sheetName}`);
    }

    const sheet = {
      id: `sheet${this._nextSheetId++}`,
      name: sheetName,
      cells: new Map(),
      selection: { startRow: 0, startCol: 0, endRow: 0, endCol: 0 }
    };
    this._sheets.push(sheet);
    this._activeSheetId = sheet.id;

    const payload = { sheet: this.getActiveSheet() };
    for (const listener of this._sheetListeners) listener(payload);

    return { id: sheet.id, name: sheet.name };
  }

  renameSheet(oldName, newName) {
    const from = String(oldName);
    const to = String(newName);
    if (to.trim().length === 0) throw new Error("New sheet name must be a non-empty string");
    if (this._sheets.some((s) => s.name === to)) {
      throw new Error(`Sheet already exists: ${to}`);
    }

    const sheet = this._sheets.find((s) => s.name === from);
    if (!sheet) throw new Error(`Unknown sheet: ${from}`);
    sheet.name = to;
    return { id: sheet.id, name: sheet.name };
  }

  deleteSheet(name) {
    const sheetName = String(name);
    const idx = this._sheets.findIndex((s) => s.name === sheetName);
    if (idx === -1) throw new Error(`Unknown sheet: ${sheetName}`);
    if (this._sheets.length === 1) {
      throw new Error("Cannot delete the last remaining sheet");
    }

    const sheet = this._sheets[idx];
    const wasActive = sheet.id === this._activeSheetId;
    this._sheets.splice(idx, 1);

    if (wasActive) {
      // Mirror typical spreadsheet behavior: deleting the active sheet activates another.
      this._activeSheetId = this._sheets[0].id;
      const payload = { sheet: this.getActiveSheet() };
      for (const listener of this._sheetListeners) listener(payload);
    }
  }

  activateSheet(name) {
    const sheetName = String(name);
    const sheet = this._sheets.find((s) => s.name === sheetName);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
    if (sheet.id === this._activeSheetId) return this.getActiveSheet();

    this._activeSheetId = sheet.id;
    const payload = { sheet: this.getActiveSheet() };
    for (const listener of this._sheetListeners) listener(payload);
    return this.getActiveSheet();
  }

  onSheetActivated(callback) {
    this._sheetListeners.add(callback);
    return { dispose: () => this._sheetListeners.delete(callback) };
  }

  setSelection(range) {
    const sheet = this._getActiveSheetRecord();
    const { startRow, startCol, endRow, endCol } = range ?? {};
    sheet.selection = { startRow, startCol, endRow, endCol };

    let selection;
    try {
      selection = this.getSelection();
    } catch {
      // Best-effort: still emit the event so extensions can observe selection movement, but
      // do not materialize a huge values matrix for Excel-scale selections.
      selection = { startRow, startCol, endRow, endCol, values: [], truncated: true };
    }

    const payload = { sheetId: sheet.id, selection };
    for (const listener of this._selectionListeners) listener(payload);
  }

  getSelection() {
    const { startRow, startCol, endRow, endCol } = this._getActiveSheetRecord().selection;
    assertRangeWithinLimit(startRow, startCol, endRow, endCol, { label: "Selection" });
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
    const sheet = this._getActiveSheetRecord();
    const key = cellKey(row, col);
    return sheet.cells.has(key) ? sheet.cells.get(key) : null;
  }

  _getCellFromSheet(sheet, row, col) {
    const key = cellKey(row, col);
    return sheet.cells.has(key) ? sheet.cells.get(key) : null;
  }

  _setCellInSheet(sheet, row, col, value, { emitCellChanged = true } = {}) {
    sheet.cells.set(cellKey(row, col), value);
    if (!emitCellChanged) return;
    const payload = { sheetId: sheet.id, row, col, value };
    for (const listener of this._cellListeners) listener(payload);
  }

  setCell(row, col, value) {
    const sheet = this._getActiveSheetRecord();
    this._setCellInSheet(sheet, row, col, value, { emitCellChanged: true });
  }

  onCellChanged(callback) {
    this._cellListeners.add(callback);
    return { dispose: () => this._cellListeners.delete(callback) };
  }

  getRange(ref) {
    const { sheetName, startRow, startCol, endRow, endCol } = parseA1RangeRef(ref);
    assertRangeWithinLimit(startRow, startCol, endRow, endCol);
    const sheet =
      sheetName == null ? this._getActiveSheetRecord() : this._getSheetRecordByName(sheetName);
    return {
      startRow,
      startCol,
      endRow,
      endCol,
      values: this.getRangeValues(startRow, startCol, endRow, endCol, sheet)
    };
  }

  setRange(ref, values) {
    const { sheetName, startRow, startCol, endRow, endCol } = parseA1RangeRef(ref);
    assertRangeWithinLimit(startRow, startCol, endRow, endCol);
    const sheet =
      sheetName == null ? this._getActiveSheetRecord() : this._getSheetRecordByName(sheetName);
    const expectedRows = endRow - startRow + 1;
    const expectedCols = endCol - startCol + 1;
    const emitCellChanged = sheet.id === this._activeSheetId;

    if (!Array.isArray(values) || values.length !== expectedRows) {
      throw new Error(
        `Range values must be a ${expectedRows}x${expectedCols} array (got ${Array.isArray(values) ? values.length : 0} rows)`
      );
    }

    for (let r = 0; r < expectedRows; r++) {
      const rowValues = values[r];
      if (!Array.isArray(rowValues) || rowValues.length !== expectedCols) {
        throw new Error(
          `Range values must be a ${expectedRows}x${expectedCols} array (row ${r} has ${Array.isArray(rowValues) ? rowValues.length : 0} cols)`
        );
      }
      for (let c = 0; c < expectedCols; c++) {
        this._setCellInSheet(sheet, startRow + r, startCol + c, rowValues[c], { emitCellChanged });
      }
    }
  }

  getRangeValues(startRow, startCol, endRow, endCol, sheet = this._getActiveSheetRecord()) {
    const rows = [];
    for (let r = startRow; r <= endRow; r++) {
      const cols = [];
      for (let c = startCol; c <= endCol; c++) {
        cols.push(this._getCellFromSheet(sheet, r, c));
      }
      rows.push(cols);
    }
    return rows;
  }
}

module.exports = {
  InMemorySpreadsheet
};
