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
      cloned.#cells.set(addr, { value: cell.value, formula: cell.formula });
    }
    return cloned;
  }

  getCell(address) {
    const a1 = normalizeA1Address(address);
    const cell = this.#cells.get(a1);
    if (!cell) return { value: undefined, formula: null };
    return { value: cell.value, formula: cell.formula };
  }

  setCellValue(address, value) {
    const a1 = normalizeA1Address(address);
    this.#cells.set(a1, { value, formula: null });
  }

  setCellFormula(address, formula) {
    const a1 = normalizeA1Address(address);
    this.#cells.set(a1, { value: undefined, formula: String(formula) });
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
}

export function diffWorkbooks(before, after) {
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

      if (Object.is(beforeCell.value, afterCell.value) && beforeCell.formula === afterCell.formula) continue;
      diffs.push({
        sheet: sheetName,
        address,
        before: { value: beforeCell.value, formula: beforeCell.formula },
        after: { value: afterCell.value, formula: afterCell.formula }
      });
    }
  }

  diffs.sort((a, b) => {
    if (a.sheet !== b.sheet) return a.sheet.localeCompare(b.sheet);
    return a.address.localeCompare(b.address);
  });
  return diffs;
}

export function compareWorkbooks(expected, actual) {
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

      if (Object.is(expectedCell.value, actualCell.value) && expectedCell.formula === actualCell.formula) continue;
      mismatches.push({
        sheet: sheetName,
        address,
        expected: { value: expectedCell.value, formula: expectedCell.formula },
        actual: { value: actualCell.value, formula: actualCell.formula }
      });
    }
  }

  mismatches.sort((a, b) => {
    if (a.sheet !== b.sheet) return a.sheet.localeCompare(b.sheet);
    return a.address.localeCompare(b.address);
  });
  return mismatches;
}

