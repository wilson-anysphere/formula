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
}

export class Sheet {
  /**
   * @param {Workbook} workbook
   * @param {string} name
   */
  constructor(workbook, name) {
    this.workbook = workbook;
    this.name = name;
    /** @type {Map<string, { value: any, format: Record<string, any> }>} */
    this.cells = new Map();
  }

  /** @type {Workbook} */
  workbook;

  /** @type {string} */
  name;

  /** @type {Map<string, { value: any, format: Record<string, any> }>} */
  cells;

  getRange(address) {
    return new Range(this, parseRangeAddress(address));
  }

  setCellValue(address, value) {
    this.getRange(address).setValue(value);
  }

  setRangeValues(address, values) {
    this.getRange(address).setValues(values);
  }

  getCellValue(row, col) {
    return this.getCell(row, col).value;
  }

  getCellFormat(row, col) {
    return this.getCell(row, col).format;
  }

  getCell(row, col) {
    const id = key(row, col);
    const existing = this.cells.get(id);
    if (existing) return existing;
    const created = { value: null, format: {} };
    this.cells.set(id, created);
    return created;
  }

  setCellValueInternal(row, col, value) {
    this.getCell(row, col).value = value;
  }

  setCellFormatInternal(row, col, format) {
    const cell = this.getCell(row, col);
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
        sheetOut[k] = { value: v.value, format: { ...v.format } };
      }
      out[name] = sheetOut;
    }
    return out;
  }
}
