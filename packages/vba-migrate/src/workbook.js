import { normalizeA1Address, rowColToA1 } from "./a1.js";

export class Sheet {
  #name;
  #cells;

  constructor(name) {
    this.#name = String(name || "");
    if (!this.#name) throw new Error("Sheet name is required");
    this.#cells = new Map();
  }

  get name() {
    return this.#name;
  }

  clone() {
    const cloned = new Sheet(this.#name);
    for (const [addr, cell] of this.#cells.entries()) {
      cloned.#cells.set(addr, { value: cell.value, formula: cell.formula, format: cell.format ?? null });
    }
    return cloned;
  }

  getCell(address) {
    const a1 = normalizeA1Address(address);
    const cell = this.#cells.get(a1);
    if (!cell) return { value: undefined, formula: null, format: null };
    return { value: cell.value, formula: cell.formula, format: cell.format ?? null };
  }

  setCellValue(address, value) {
    const a1 = normalizeA1Address(address);
    const prev = this.#cells.get(a1);
    this.#cells.set(a1, { value, formula: null, format: prev?.format ?? null });
  }

  setCellFormula(address, formula) {
    const a1 = normalizeA1Address(address);
    const prev = this.#cells.get(a1);
    this.#cells.set(a1, { value: undefined, formula: String(formula), format: prev?.format ?? null });
  }

  setCellFormat(address, format) {
    const a1 = normalizeA1Address(address);
    const prev = this.#cells.get(a1);
    this.#cells.set(a1, { value: prev?.value, formula: prev?.formula ?? null, format: format ?? null });
  }

  setCellValueByRowCol(row, col, value) {
    const addr = rowColToA1(row, col);
    this.setCellValue(addr, value);
  }

  setCellFormulaByRowCol(row, col, formula) {
    const addr = rowColToA1(row, col);
    this.setCellFormula(addr, formula);
  }

  /**
   * Returns a Map of A1 address => cell payload (value/formula).
   * Intended for diffing + diagnostics, not for mutation.
   */
  snapshotCells() {
    return new Map(this.#cells.entries());
  }
}

export class Workbook {
  #sheetsByName;
  #activeSheetName;

  constructor() {
    this.#sheetsByName = new Map();
    this.#activeSheetName = null;
  }

  addSheet(name, { makeActive = false } = {}) {
    const sheet = new Sheet(name);
    this.#sheetsByName.set(sheet.name, sheet);
    if (makeActive || !this.#activeSheetName) this.#activeSheetName = sheet.name;
    return sheet;
  }

  getSheet(name) {
    const sheetName = String(name || "");
    const sheet = this.#sheetsByName.get(sheetName);
    if (!sheet) throw new Error(`Unknown sheet: ${sheetName}`);
    return sheet;
  }

  hasSheet(name) {
    return this.#sheetsByName.has(String(name || ""));
  }

  get sheetNames() {
    return [...this.#sheetsByName.keys()];
  }

  get activeSheet() {
    if (!this.#activeSheetName) throw new Error("Workbook has no active sheet");
    return this.getSheet(this.#activeSheetName);
  }

  set activeSheetName(name) {
    const sheetName = String(name || "");
    if (!this.#sheetsByName.has(sheetName)) throw new Error(`Unknown sheet: ${sheetName}`);
    this.#activeSheetName = sheetName;
  }

  clone() {
    const cloned = new Workbook();
    for (const [name, sheet] of this.#sheetsByName.entries()) {
      cloned.#sheetsByName.set(name, sheet.clone());
    }
    cloned.#activeSheetName = this.#activeSheetName;
    return cloned;
  }

  snapshot() {
    const snapshot = new Map();
    for (const [name, sheet] of this.#sheetsByName.entries()) {
      snapshot.set(name, sheet.snapshotCells());
    }
    return snapshot;
  }

  /**
   * Serialize to the canonical oracle workbook JSON representation.
   * This format is consumed by the Rust VBA oracle CLI.
   */
  toOracleWorkbook({ vbaModules = [] } = {}) {
    const sheets = [...this.#sheetsByName.entries()]
      .map(([name, sheet]) => {
        const cellsSnapshot = sheet.snapshotCells();
        const cells = {};
        const addresses = [...cellsSnapshot.keys()].sort((a, b) => a.localeCompare(b));
        for (const address of addresses) {
          const cell = cellsSnapshot.get(address);
          if (!cell) continue;
          const payload = {};
          if (cell.value !== undefined && cell.value !== null) payload.value = cell.value;
          if (cell.formula !== null && cell.formula !== undefined) payload.formula = cell.formula;
          if (cell.format !== null && cell.format !== undefined) payload.format = cell.format;
          // Only serialize non-empty cells.
          if (Object.keys(payload).length) cells[address] = payload;
        }
        return { name, cells };
      })
      .sort((a, b) => a.name.localeCompare(b.name));

    const activeSheet = this.#activeSheetName ?? null;
    return {
      schemaVersion: 1,
      activeSheet,
      sheets,
      vbaModules
    };
  }

  static fromOracleWorkbook(payload) {
    if (!payload || typeof payload !== "object") {
      throw new Error("Invalid oracle workbook payload");
    }
    const workbook = new Workbook();
    const sheets = Array.isArray(payload.sheets) ? payload.sheets : [];
    for (const sheet of sheets) {
      const name = sheet?.name;
      if (!name) continue;
      workbook.addSheet(name);
      const worksheet = workbook.getSheet(name);
      const cells = sheet?.cells ?? {};
      for (const [addr, cell] of Object.entries(cells)) {
        if (cell?.value !== undefined && cell?.value !== null) {
          worksheet.setCellValue(addr, cell.value);
        }
        if (cell?.formula !== undefined && cell?.formula !== null) {
          worksheet.setCellFormula(addr, cell.formula);
        }
        if (cell?.format !== undefined && cell?.format !== null) {
          worksheet.setCellFormat(addr, cell.format);
        }
      }
    }
    if (payload.activeSheet && workbook.hasSheet(payload.activeSheet)) {
      workbook.activeSheetName = payload.activeSheet;
    }
    return workbook;
  }

  toBytes({ vbaModules = [] } = {}) {
    const payload = this.toOracleWorkbook({ vbaModules });
    return Buffer.from(JSON.stringify(payload), "utf8");
  }

  static fromBytes(bytes) {
    const text = Buffer.isBuffer(bytes) ? bytes.toString("utf8") : String(bytes ?? "");
    const payload = JSON.parse(text);
    return Workbook.fromOracleWorkbook(payload);
  }
}

function normalizeEmpty(value) {
  return value === null ? undefined : value;
}

function valuesEqual(a, b, { floatTolerance = 0 } = {}) {
  const av = normalizeEmpty(a);
  const bv = normalizeEmpty(b);
  if (av === undefined && bv === undefined) return true;
  if (typeof av === "number" && typeof bv === "number" && floatTolerance > 0) {
    return Math.abs(av - bv) <= floatTolerance;
  }
  return Object.is(av, bv);
}

function formulasEqual(a, b) {
  const af = a ?? null;
  const bf = b ?? null;
  return af === bf;
}

function formatsEqual(a, b) {
  const af = a ?? null;
  const bf = b ?? null;
  return af === bf;
}

export function diffWorkbooks(before, after, options = {}) {
  const beforeSnapshot = before.snapshot();
  const afterSnapshot = after.snapshot();

  const diffs = [];
  const sheetNames = new Set([...beforeSnapshot.keys(), ...afterSnapshot.keys()]);
  for (const sheetName of sheetNames) {
    const beforeCells = beforeSnapshot.get(sheetName) ?? new Map();
    const afterCells = afterSnapshot.get(sheetName) ?? new Map();

    const addresses = new Set([...beforeCells.keys(), ...afterCells.keys()]);
    for (const address of addresses) {
      const beforeCell = beforeCells.get(address) ?? { value: undefined, formula: null };
      const afterCell = afterCells.get(address) ?? { value: undefined, formula: null };

      if (
        valuesEqual(beforeCell.value, afterCell.value, options) &&
        formulasEqual(beforeCell.formula, afterCell.formula) &&
        formatsEqual(beforeCell.format, afterCell.format)
      )
        continue;
      diffs.push({
        sheet: sheetName,
        address,
        before: { value: beforeCell.value, formula: beforeCell.formula, format: beforeCell.format ?? null },
        after: { value: afterCell.value, formula: afterCell.formula, format: afterCell.format ?? null }
      });
    }
  }

  diffs.sort((a, b) => {
    if (a.sheet !== b.sheet) return a.sheet.localeCompare(b.sheet);
    return a.address.localeCompare(b.address);
  });
  return diffs;
}

export function compareWorkbooks(expected, actual, options = {}) {
  const expectedSnapshot = expected.snapshot();
  const actualSnapshot = actual.snapshot();

  const mismatches = [];
  const sheetNames = new Set([...expectedSnapshot.keys(), ...actualSnapshot.keys()]);
  for (const sheetName of sheetNames) {
    const expectedCells = expectedSnapshot.get(sheetName) ?? new Map();
    const actualCells = actualSnapshot.get(sheetName) ?? new Map();

    const addresses = new Set([...expectedCells.keys(), ...actualCells.keys()]);
    for (const address of addresses) {
      const expectedCell = expectedCells.get(address) ?? { value: undefined, formula: null };
      const actualCell = actualCells.get(address) ?? { value: undefined, formula: null };

      if (
        valuesEqual(expectedCell.value, actualCell.value, options) &&
        formulasEqual(expectedCell.formula, actualCell.formula) &&
        formatsEqual(expectedCell.format, actualCell.format)
      )
        continue;
      mismatches.push({
        sheet: sheetName,
        address,
        expected: { value: expectedCell.value, formula: expectedCell.formula, format: expectedCell.format ?? null },
        actual: { value: actualCell.value, formula: actualCell.formula, format: actualCell.format ?? null }
      });
    }
  }

  mismatches.sort((a, b) => {
    if (a.sheet !== b.sheet) return a.sheet.localeCompare(b.sheet);
    return a.address.localeCompare(b.address);
  });
  return mismatches;
}
