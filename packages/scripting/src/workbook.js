import { formatRangeAddress, parseRangeAddress } from "./a1.js";
import { TypedEventEmitter } from "./events.js";

function key(row, col) {
  return `${row},${col}`;
}

export class Range {
  constructor(
    sheet,
    coords,
  ) {
    this.sheet = sheet;
    this.coords = coords;
  }

  /** @type {Sheet} */
  sheet;

  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }} */
  coords;

  get address() {
    return formatRangeAddress(this.coords);
  }

  getValues() {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        row.push(this.sheet.getCellValue(this.coords.startRow + r, this.coords.startCol + c));
      }
      out.push(row);
    }
    return out;
  }

  setValues(values) {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    if (values.length !== rows || values.some((row) => row.length !== cols)) {
      throw new Error(
        `setValues expected ${rows}x${cols} matrix for range ${this.address}, got ${values.length}x${values[0]?.length ?? 0}`,
      );
    }

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        this.sheet.setCellValueInternal(this.coords.startRow + r, this.coords.startCol + c, values[r][c]);
      }
    }

    this.sheet.workbook.events.emit("cellChanged", {
      sheetName: this.sheet.name,
      address: this.address,
      values,
    });
  }

  getFormulas() {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        row.push(this.sheet.getCellFormula(this.coords.startRow + r, this.coords.startCol + c));
      }
      out.push(row);
    }
    return out;
  }

  setFormulas(formulas) {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    if (formulas.length !== rows || formulas.some((row) => row.length !== cols)) {
      throw new Error(
        `setFormulas expected ${rows}x${cols} matrix for range ${this.address}, got ${formulas.length}x${formulas[0]?.length ?? 0}`,
      );
    }

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        this.sheet.setCellFormulaInternal(this.coords.startRow + r, this.coords.startCol + c, formulas[r][c]);
      }
    }

    this.sheet.workbook.events.emit("formulaChanged", {
      sheetName: this.sheet.name,
      address: this.address,
      formulas,
    });
  }

  getValue() {
    const values = this.getValues();
    if (values.length !== 1 || values[0].length !== 1) {
      throw new Error(`getValue is only valid for a single cell, got range ${this.address}`);
    }
    return values[0][0];
  }

  setValue(value) {
    const range = this.coords;
    if (range.startRow !== range.endRow || range.startCol !== range.endCol) {
      throw new Error(`setValue is only valid for a single cell, got range ${this.address}`);
    }

    this.sheet.setCellValueInternal(range.startRow, range.startCol, value);
    this.sheet.workbook.events.emit("cellChanged", {
      sheetName: this.sheet.name,
      address: this.address,
      values: [[value]],
    });
  }

  setFormat(format) {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        this.sheet.setCellFormatInternal(this.coords.startRow + r, this.coords.startCol + c, format);
      }
    }

    this.sheet.workbook.events.emit("formatChanged", {
      sheetName: this.sheet.name,
      address: this.address,
      format,
    });
  }

  getFormat() {
    const range = this.coords;
    return this.sheet.getCellFormat(range.startRow, range.startCol);
  }

  getFormats() {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        row.push({ ...this.sheet.getCellFormat(this.coords.startRow + r, this.coords.startCol + c) });
      }
      out.push(row);
    }
    return out;
  }

  setFormats(formats) {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    if (formats.length !== rows || formats.some((row) => row.length !== cols)) {
      throw new Error(
        `setFormats expected ${rows}x${cols} matrix for range ${this.address}, got ${formats.length}x${formats[0]?.length ?? 0}`,
      );
    }

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        this.sheet.setCellFormatInternal(this.coords.startRow + r, this.coords.startCol + c, formats[r][c]);
      }
    }

    this.sheet.workbook.events.emit("formatChanged", {
      sheetName: this.sheet.name,
      address: this.address,
      format: null,
      formats,
    });
  }
}

export class Sheet {
  /**
   * @param {Workbook} workbook
   * @param {string} name
   */
  constructor(workbook, name) {
    this.workbook = workbook;
    this.name = name;
    /** @type {Map<string, { value: any, formula: string | null, format: Record<string, any> }>} */
    this.cells = new Map();
  }

  /** @type {Workbook} */
  workbook;

  /** @type {string} */
  name;

  /** @type {Map<string, { value: any, formula: string | null, format: Record<string, any> }>} */
  cells;

  getRange(address) {
    return new Range(this, parseRangeAddress(address));
  }

  getCell(row, col) {
    return new Range(this, { startRow: row, startCol: col, endRow: row, endCol: col });
  }

  setCellValue(address, value) {
    this.getRange(address).setValue(value);
  }

  setRangeValues(address, values) {
    this.getRange(address).setValues(values);
  }

  setCellFormula(address, formula) {
    this.getRange(address).setFormulas([[formula]]);
  }

  getUsedRange() {
    let minRow = Infinity;
    let minCol = Infinity;
    let maxRow = -1;
    let maxCol = -1;

    for (const [id, cell] of this.cells.entries()) {
      const [rowStr, colStr] = id.split(",");
      const row = Number.parseInt(rowStr, 10);
      const col = Number.parseInt(colStr, 10);
      const hasFormat = cell.format && Object.keys(cell.format).length > 0;
      if (cell.value === null && cell.formula === null && !hasFormat) continue;
      minRow = Math.min(minRow, row);
      minCol = Math.min(minCol, col);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }

    if (!Number.isFinite(minRow)) {
      return this.getRange("A1");
    }

    return new Range(this, { startRow: minRow, startCol: minCol, endRow: maxRow, endCol: maxCol });
  }

  getCellValue(row, col) {
    return this.getCellState(row, col).value;
  }

  getCellFormula(row, col) {
    return this.getCellState(row, col).formula;
  }

  getCellFormat(row, col) {
    return this.getCellState(row, col).format;
  }

  getCellState(row, col) {
    const id = key(row, col);
    const existing = this.cells.get(id);
    if (existing) return existing;
    const created = { value: null, formula: null, format: {} };
    this.cells.set(id, created);
    return created;
  }

  setCellValueInternal(row, col, value) {
    const cell = this.getCellState(row, col);
    cell.value = value;
    cell.formula = null;
  }

  setCellFormulaInternal(row, col, formula) {
    const cell = this.getCellState(row, col);
    cell.formula = formula ?? null;
    cell.value = null;
  }

  setCellFormatInternal(row, col, format) {
    const cell = this.getCellState(row, col);
    if (format == null) {
      cell.format = {};
      return;
    }
    cell.format = { ...cell.format, ...format };
  }
}

export class Workbook {
  constructor() {
    this.events = new TypedEventEmitter();
    /** @type {Map<string, Sheet>} */
    this.sheets = new Map();
    /** @type {string | null} */
    this.activeSheetName = null;
    /** @type {{ sheetName: string, address: string } | null} */
    this.selection = null;
  }

  /** @type {TypedEventEmitter} */
  events;

  /** @type {Map<string, Sheet>} */
  sheets;

  /** @type {string | null} */
  activeSheetName;

  /** @type {{ sheetName: string, address: string } | null} */
  selection;

  addSheet(name) {
    if (this.sheets.has(name)) {
      throw new Error(`Sheet already exists: ${name}`);
    }
    const sheet = new Sheet(this, name);
    this.sheets.set(name, sheet);
    if (!this.activeSheetName) {
      this.activeSheetName = name;
      this.selection = { sheetName: name, address: "A1" };
    }
    return sheet;
  }

  getSheet(name) {
    const sheet = this.sheets.get(name);
    if (!sheet) {
      throw new Error(`Unknown sheet: ${name}`);
    }
    return sheet;
  }

  getSheets() {
    return [...this.sheets.values()];
  }

  getActiveSheet() {
    if (!this.activeSheetName) {
      throw new Error("Workbook has no sheets");
    }
    return this.getSheet(this.activeSheetName);
  }

  setActiveSheet(name) {
    if (!this.sheets.has(name)) {
      throw new Error(`Unknown sheet: ${name}`);
    }
    this.activeSheetName = name;
  }

  getSelection() {
    if (!this.selection) {
      const active = this.getActiveSheet().name;
      this.selection = { sheetName: active, address: "A1" };
    }
    return this.selection;
  }

  setSelection(sheetName, address) {
    if (!this.sheets.has(sheetName)) {
      throw new Error(`Unknown sheet: ${sheetName}`);
    }

    // Validate address early to keep selection consistent.
    parseRangeAddress(address);
    this.selection = { sheetName, address };
    this.events.emit("selectionChanged", this.selection);
  }

  snapshot() {
    const out = {};
    for (const [name, sheet] of this.sheets.entries()) {
      const sheetOut = {};
      for (const [k, v] of sheet.cells.entries()) {
        sheetOut[k] = { value: v.value, formula: v.formula, format: { ...v.format } };
      }
      out[name] = sheetOut;
    }
    return out;
  }
}
