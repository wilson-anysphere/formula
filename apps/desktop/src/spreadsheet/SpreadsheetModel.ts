import { FormulaBarModel } from "../formula-bar/FormulaBarModel.js";
import { evaluateFormula, type SpreadsheetValue } from "./evaluateFormula.js";
import { normalizeRange, parseA1, rangeToA1, type CellAddress, type RangeAddress } from "./a1.js";

export type Cell = { input: string; value: SpreadsheetValue };

export class SpreadsheetModel {
  readonly formulaBar = new FormulaBarModel();
  readonly #cells = new Map<string, Cell>();
  #activeCell = "A1";
  #selection: RangeAddress = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

  constructor(initial?: Record<string, string | number>) {
    if (initial) {
      for (const [addr, input] of Object.entries(initial)) {
        this.setCellInput(addr, String(input));
      }
    }
    this.selectCell(this.#activeCell);
  }

  get activeCell(): string {
    return this.#activeCell;
  }

  get selection(): RangeAddress {
    return this.#selection;
  }

  getCell(address: string): Cell {
    return this.#cells.get(address) ?? { input: "", value: null };
  }

  getCellValue(address: string): SpreadsheetValue {
    return this.getCell(address).value;
  }

  selectCell(address: string): void {
    this.#activeCell = address;
    const cell = this.getCell(address);
    this.formulaBar.setActiveCell({ address, input: cell.input, value: cell.value });
  }

  setCellInput(address: string, input: string): void {
    const value = evaluateFormula(input, (ref) => this.getCellValue(ref));
    this.#cells.set(address, { input, value });
    if (address === this.#activeCell) {
      this.formulaBar.setActiveCell({ address, input, value });
    }
  }

  beginFormulaEdit(): void {
    this.formulaBar.beginEdit();
  }

  typeInFormulaBar(newText: string, cursorIndex: number = newText.length): void {
    this.formulaBar.updateDraft(newText, cursorIndex, cursorIndex);
  }

  commitFormulaBar(): void {
    const committed = this.formulaBar.commit();
    this.setCellInput(this.#activeCell, committed);
  }

  cancelFormulaBar(): void {
    this.formulaBar.cancel();
  }

  beginRangeSelection(startCell: string): void {
    const start = parseA1(startCell);
    if (!start) throw new Error(`beginRangeSelection: invalid start cell ${startCell}`);
    this.#selection = { start, end: start };
    this.formulaBar.beginRangeSelection(this.#selection);
  }

  updateRangeSelection(endCell: string): void {
    const end = parseA1(endCell);
    if (!end) throw new Error(`updateRangeSelection: invalid end cell ${endCell}`);
    this.#selection = normalizeRange(this.#selection.start, end);
    this.formulaBar.updateRangeSelection(this.#selection);
  }

  endRangeSelection(): void {
    this.formulaBar.endRangeSelection();
  }

  selectionA1(): string {
    return rangeToA1(this.#selection);
  }
}
