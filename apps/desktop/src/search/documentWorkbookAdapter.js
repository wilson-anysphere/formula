import { normalizeName } from "../../../../packages/search/index.js";

function parseCellKey(key) {
  const [r, c] = String(key).split(",");
  const row = Number.parseInt(r, 10);
  const col = Number.parseInt(c, 10);
  if (!Number.isInteger(row) || row < 0 || !Number.isInteger(col) || col < 0) {
    return null;
  }
  return { row, col };
}

function normalizeFormula(formula) {
  if (formula == null) return null;
  const text = String(formula);
  if (text === "") return "";
  return text.startsWith("=") ? text : `=${text}`;
}

/**
 * Adapter that exposes a DocumentController-like model through the interface expected
 * by `packages/search` (workbook -> sheets -> cells).
 *
 * This keeps `packages/search` UI-agnostic while still allowing the desktop app
 * to reuse the same search/replace implementation.
 */
export class DocumentWorkbookAdapter {
  /**
   * @param {{
   *   document: import("../document/documentController.js").DocumentController,
   * }} params
   */
  constructor({ document }) {
    if (!document) throw new Error("DocumentWorkbookAdapter: document is required");
    this.document = document;
    /** @type {Map<string, DocumentSheetAdapter>} */
    this.#sheetsById = new Map();

    /** @type {Map<string, any>} */
    this.names = new Map();
    /** @type {Map<string, any>} */
    this.tables = new Map();
  }

  /** @type {Map<string, DocumentSheetAdapter>} */
  #sheetsById;

  get sheets() {
    const ids = typeof this.document.getSheetIds === "function" ? this.document.getSheetIds() : [];
    if (ids.length === 0) return [];

    return ids.map((id) => this.getSheet(id));
  }

  getSheet(sheetName) {
    const key = String(sheetName);
    let sheet = this.#sheetsById.get(key);
    if (!sheet) {
      sheet = new DocumentSheetAdapter(this.document, key);
      this.#sheetsById.set(key, sheet);
    }
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

class DocumentSheetAdapter {
  /**
   * @param {import("../document/documentController.js").DocumentController} document
   * @param {string} sheetId
   */
  constructor(document, sheetId) {
    this.document = document;
    this.name = sheetId;
  }

  getUsedRange() {
    if (typeof this.document.getUsedRange === "function") {
      return this.document.getUsedRange(this.name);
    }
    return null;
  }

  getCell(row, col) {
    const state = this.document.getCell(this.name, { row, col });
    const formula = normalizeFormula(state.formula);
    const value = state.value ?? null;

    if (value == null && formula == null) return null;

    const display = value != null ? String(value) : formula ?? "";
    return { value, formula, display };
  }

  setCell(row, col, cell) {
    if (!cell || (cell.value == null && (cell.formula == null || cell.formula === ""))) {
      this.document.clearCell(this.name, { row, col });
      return;
    }

    if (cell.formula != null && cell.formula !== "") {
      const formula = String(cell.formula);
      const normalized = normalizeFormula(formula);
      // Keep the controller storage flexible: allow formulas with or without leading "=",
      // but normalize for user-facing semantics.
      this.document.setCellFormula(this.name, { row, col }, normalized);
      return;
    }

    this.document.setCellValue(this.name, { row, col }, cell.value);
  }

  *iterateCells(range, { order = "byRows" } = {}) {
    const used = this.getUsedRange();
    if (!used) return;

    const startRow = Math.max(used.startRow, range.startRow);
    const endRow = Math.min(used.endRow, range.endRow);
    const startCol = Math.max(used.startCol, range.startCol);
    const endCol = Math.min(used.endCol, range.endCol);
    if (startRow > endRow || startCol > endCol) return;

    const sheet = this.document.model?.sheets?.get(this.name);
    const entries = sheet?.cells ? Array.from(sheet.cells.entries()) : [];

    const results = [];
    for (const [key, cell] of entries) {
      if (!cell) continue;
      if (cell.value == null && cell.formula == null) continue;
      const parsed = parseCellKey(key);
      if (!parsed) continue;
      if (
        parsed.row < startRow ||
        parsed.row > endRow ||
        parsed.col < startCol ||
        parsed.col > endCol
      ) {
        continue;
      }

      const formula = normalizeFormula(cell.formula);
      const value = cell.value ?? null;
      const display = value != null ? String(value) : formula ?? "";
      results.push({ row: parsed.row, col: parsed.col, cell: { value, formula, display } });
    }

    results.sort((a, b) => {
      if (order === "byColumns") {
        if (a.col !== b.col) return a.col - b.col;
        return a.row - b.row;
      }
      if (a.row !== b.row) return a.row - b.row;
      return a.col - b.col;
    });

    for (const entry of results) yield entry;
  }
}

