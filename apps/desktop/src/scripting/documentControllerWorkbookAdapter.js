import { formatCellAddress, parseRangeAddress } from "../../../../packages/scripting/src/a1.js";
import { TypedEventEmitter } from "../../../../packages/scripting/src/events.js";

function valueEquals(a, b) {
  if (a === b) return true;
  if (a == null || b == null) return false;
  if (typeof a === "object" && typeof b === "object") {
    try {
      return JSON.stringify(a) === JSON.stringify(b);
    } catch {
      return false;
    }
  }
  return false;
}

function cellInputFromState(state) {
  if (state.formula != null) return state.formula;
  return state.value ?? null;
}

function isFormulaString(input) {
  if (typeof input !== "string") return false;
  const trimmed = input.trimStart();
  return trimmed.startsWith("=") && trimmed.length > 1;
}

/**
 * Exposes a `DocumentController` instance through the `@formula/scripting` Workbook/Sheet/Range
 * surface area.
 *
 * This adapter is intentionally minimal: it focuses on the RPC methods used by `ScriptRuntime`
 * and the events consumed by `MacroRecorder`.
 */
export class DocumentControllerWorkbookAdapter {
  /**
   * @param {import("../document/documentController.js").DocumentController} documentController
   * @param {{ activeSheetName?: string }} [options]
   */
  constructor(documentController, options = {}) {
    this.documentController = documentController;
    this.events = new TypedEventEmitter();
    /** @type {Map<string, DocumentControllerSheetAdapter>} */
    this.sheets = new Map();
    this.activeSheetName = options.activeSheetName ?? "Sheet1";
    /** @type {{ sheetName: string, address: string } | null} */
    this.selection = null;

    this.unsubscribes = [
      this.documentController.on("change", (payload) => this.#handleDocumentChange(payload)),
    ];
  }

  /** @type {import("../document/documentController.js").DocumentController} */
  documentController;

  /** @type {TypedEventEmitter} */
  events;

  /** @type {Map<string, DocumentControllerSheetAdapter>} */
  sheets;

  /** @type {string} */
  activeSheetName;

  /** @type {{ sheetName: string, address: string } | null} */
  selection;

  /** @type {Array<() => void>} */
  unsubscribes;

  dispose() {
    for (const unsub of this.unsubscribes) unsub();
    this.unsubscribes = [];
  }

  getSheet(name) {
    const sheetName = String(name);
    let sheet = this.sheets.get(sheetName);
    if (!sheet) {
      sheet = new DocumentControllerSheetAdapter(this, sheetName);
      this.sheets.set(sheetName, sheet);
    }
    return sheet;
  }

  getActiveSheet() {
    return this.getSheet(this.activeSheetName);
  }

  getActiveSheetName() {
    return this.activeSheetName;
  }

  setActiveSheet(name) {
    this.activeSheetName = String(name);
  }

  getSelection() {
    if (!this.selection) {
      this.selection = { sheetName: this.getActiveSheet().name, address: "A1" };
    }
    return this.selection;
  }

  setSelection(sheetName, address) {
    const normalizedSheetName = String(sheetName);
    // Validate address early to keep selection consistent.
    parseRangeAddress(String(address));

    this.selection = { sheetName: normalizedSheetName, address: String(address) };
    this.events.emit("selectionChanged", this.selection);
  }

  /**
   * @param {{ deltas: Array<any> }} payload
   */
  #handleDocumentChange(payload) {
    const deltas = Array.isArray(payload?.deltas) ? payload.deltas : [];
    for (const delta of deltas) {
      if (!delta) continue;
      const sheetName = delta.sheetId;
      const address = formatCellAddress({ row: delta.row, col: delta.col });

      const valueChanged =
        !valueEquals(delta.before?.value ?? null, delta.after?.value ?? null) ||
        (delta.before?.formula ?? null) !== (delta.after?.formula ?? null);

      const formatChanged = (delta.before?.styleId ?? 0) !== (delta.after?.styleId ?? 0);

      if (valueChanged) {
        const value = cellInputFromState(delta.after ?? {});
        this.events.emit("cellChanged", {
          sheetName,
          address,
          values: [[value]],
        });
      }

      if (formatChanged) {
        const styleId = delta.after?.styleId ?? 0;
        const format = { ...this.documentController.styleTable.get(styleId) };
        this.events.emit("formatChanged", {
          sheetName,
          address,
          format,
        });
      }
    }
  }
}

class DocumentControllerSheetAdapter {
  /**
   * @param {DocumentControllerWorkbookAdapter} workbook
   * @param {string} name
   */
  constructor(workbook, name) {
    this.workbook = workbook;
    this.name = name;
  }

  /** @type {DocumentControllerWorkbookAdapter} */
  workbook;

  /** @type {string} */
  name;

  getRange(address) {
    return new DocumentControllerRangeAdapter(this, parseRangeAddress(String(address)));
  }

  setCellValue(address, value) {
    return this.getRange(address).setValue(value);
  }

  setRangeValues(address, values) {
    return this.getRange(address).setValues(values);
  }
}

class DocumentControllerRangeAdapter {
  /**
   * @param {DocumentControllerSheetAdapter} sheet
   * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} coords
   */
  constructor(sheet, coords) {
    this.sheet = sheet;
    this.coords = coords;
  }

  /** @type {DocumentControllerSheetAdapter} */
  sheet;

  /** @type {{ startRow: number, startCol: number, endRow: number, endCol: number }} */
  coords;

  get address() {
    const start = formatCellAddress({ row: this.coords.startRow, col: this.coords.startCol });
    const end = formatCellAddress({ row: this.coords.endRow, col: this.coords.endCol });
    return start === end ? start : `${start}:${end}`;
  }

  getValues() {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    const out = [];
    for (let r = 0; r < rows; r++) {
      const row = [];
      for (let c = 0; c < cols; c++) {
        const cell = this.sheet.workbook.documentController.getCell(this.sheet.name, {
          row: this.coords.startRow + r,
          col: this.coords.startCol + c,
        });
        row.push(cellInputFromState(cell));
      }
      out.push(row);
    }
    return out;
  }

  setValues(values) {
    const rows = this.coords.endRow - this.coords.startRow + 1;
    const cols = this.coords.endCol - this.coords.startCol + 1;
    if (!Array.isArray(values) || values.length !== rows || values.some((row) => row.length !== cols)) {
      throw new Error(
        `setValues expected ${rows}x${cols} matrix for range ${this.address}, got ${values.length}x${values[0]?.length ?? 0}`,
      );
    }

    this.sheet.workbook.documentController.setRangeValues(this.sheet.name, this.address, values);
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

    const coord = { row: range.startRow, col: range.startCol };

    if (typeof value === "string") {
      if (value.startsWith("'")) {
        this.sheet.workbook.documentController.setCellValue(this.sheet.name, coord, value.slice(1));
        return;
      }
      if (isFormulaString(value)) {
        this.sheet.workbook.documentController.setCellFormula(this.sheet.name, coord, value);
        return;
      }
    }

    this.sheet.workbook.documentController.setCellValue(this.sheet.name, coord, value ?? null);
  }

  getFormat() {
    const cell = this.sheet.workbook.documentController.getCell(this.sheet.name, {
      row: this.coords.startRow,
      col: this.coords.startCol,
    });
    return { ...this.sheet.workbook.documentController.styleTable.get(cell.styleId) };
  }

  setFormat(format) {
    this.sheet.workbook.documentController.setRangeFormat(this.sheet.name, this.address, format);
  }
}
