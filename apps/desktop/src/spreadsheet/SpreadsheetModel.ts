import { FormulaBarModel } from "../formula-bar/FormulaBarModel.js";
import { evaluateFormula, type SpreadsheetValue } from "./evaluateFormula.js";
import { AiCellFunctionEngine } from "./AiCellFunctionEngine.js";
import { normalizeRange, parseA1, rangeToA1, toA1, type CellAddress, type RangeAddress } from "./a1.js";
import { TabCompletionEngine } from "@formula/ai-completion";
import {
  createLocaleAwareFunctionRegistry,
  createLocaleAwarePartialFormulaParser,
  createLocaleAwareStarterFunctions,
} from "../ai/completion/parsePartialFormula.js";
import { getLocale } from "../i18n/index.js";

export type Cell = { input: string; value: SpreadsheetValue };

const DEFAULT_SHEET_ID = "Sheet1";

function currentFormulaLocaleId(): string {
  // SpreadsheetModel is primarily used in tests, where the i18n locale is not
  // always wired, but callers may still set `<html lang>`.
  try {
    const raw = typeof document !== "undefined" ? document.documentElement?.lang : "";
    const trimmed = String(raw ?? "").trim();
    if (trimmed) return trimmed;
  } catch {
    // ignore
  }
  try {
    return getLocale();
  } catch {
    return "en-US";
  }
}

export class SpreadsheetModel {
  readonly formulaBar = new FormulaBarModel();
  readonly #cells = new Map<string, Cell>();
  readonly #aiCellFunctions: AiCellFunctionEngine;
  readonly #completion = new TabCompletionEngine({
    functionRegistry: createLocaleAwareFunctionRegistry({ getLocaleId: currentFormulaLocaleId }),
    starterFunctions: createLocaleAwareStarterFunctions({ getLocaleId: currentFormulaLocaleId }),
    // Mirror the production formula bar adapter: canonicalize localized function names so
    // range-arg heuristics can work in non-en-US locales, and optionally use the WASM
    // engine when available (SpreadsheetModel doesn't provide one, so this remains sync).
    parsePartialFormula: createLocaleAwarePartialFormulaParser({ timeoutMs: 10, getLocaleId: currentFormulaLocaleId }),
  });
  #completionRequest = 0;
  #pendingCompletion: Promise<void> | null = null;
  #cellsVersion = 0;
  #activeCell = "A1";
  #selection: RangeAddress = { start: { row: 0, col: 0 }, end: { row: 0, col: 0 } };

  constructor(initial?: Record<string, string | number>) {
    const workbookId = "local-workbook";
    this.#aiCellFunctions = new AiCellFunctionEngine({
      onUpdate: () => this.#recomputeAiCells(),
      workbookId,
    });

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
    const value = evaluateFormula(input, (ref) => this.getCellValue(ref), {
      ai: this.#aiCellFunctions,
      cellAddress: `${DEFAULT_SHEET_ID}!${address}`,
      localeId: currentFormulaLocaleId(),
    });
    this.#cells.set(address, { input, value });
    this.#cellsVersion += 1;
    if (address === this.#activeCell) {
      this.formulaBar.setActiveCell({ address, input, value });
    }
  }

  beginFormulaEdit(): void {
    this.formulaBar.beginEdit();
  }

  typeInFormulaBar(newText: string, cursorIndex: number = newText.length): void {
    this.formulaBar.updateDraft(newText, cursorIndex, cursorIndex);
    this.#updateCompletions();
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

  /**
   * Wait for the most recently scheduled tab-completion request (if any).
   *
   * This is primarily used by tests to avoid asserting while completion is
   * still in-flight.
   */
  async flushTabCompletion(): Promise<void> {
    await (this.#pendingCompletion ?? Promise.resolve());
  }

  selectionA1(): string {
    return rangeToA1(this.#selection);
  }

  #updateCompletions(): void {
    if (!this.formulaBar.isEditing) {
      this.formulaBar.setAiSuggestion(null);
      return;
    }

    if (this.formulaBar.cursorStart !== this.formulaBar.cursorEnd) {
      this.formulaBar.setAiSuggestion(null);
      return;
    }

    const requestId = ++this.#completionRequest;
    const draft = this.formulaBar.draft;
    const cursor = this.formulaBar.cursorStart;
    const activeCell = this.#activeCell;

    const surroundingCells = {
      getCellValue: (row: number, col: number): SpreadsheetValue => {
        const addr = toA1({ row, col });
        const cell = this.getCell(addr);
        // Treat formulas as non-empty even if their computed value is blank. This
        // matches the production adapter that uses DocumentController's `formula`
        // field to ensure range detection considers formula-filled tables.
        if (cell.input.trimStart().startsWith("=") && (cell.value == null || cell.value === "")) {
          return cell.input;
        }
        return cell.value;
      },
      // Include locale because parsing is locale-aware (argument separators, localized function names).
      getCacheKey: () => `${this.#cellsVersion}:${this.#cells.size}:locale:${currentFormulaLocaleId()}`,
    };

    this.#pendingCompletion = this.#completion
      .getSuggestions({
        currentInput: draft,
        cursorPosition: cursor,
        cellRef: activeCell,
        surroundingCells,
      })
      .then((suggestions) => {
        if (requestId !== this.#completionRequest) return;
        const prefix = draft.slice(0, cursor);
        const suffix = draft.slice(cursor);
        const best = suggestions.find((s) => {
          if (!s || typeof s.text !== "string") return false;
          if (s.text === draft) return false;
          if (!s.text.startsWith(prefix)) return false;
          if (suffix && !s.text.endsWith(suffix)) return false;
          const ghostLength = s.text.length - prefix.length - suffix.length;
          return ghostLength > 0;
        });
        this.formulaBar.setAiSuggestion(best ? best.text : null);
      })
      .catch(() => {
        if (requestId !== this.#completionRequest) return;
        this.formulaBar.setAiSuggestion(null);
      });
  }

  #recomputeAiCells(): void {
    let changed = false;
    for (const [address, cell] of this.#cells) {
      if (!isAiFormula(cell.input)) continue;
      const value = evaluateFormula(cell.input, (ref) => this.getCellValue(ref), {
        ai: this.#aiCellFunctions,
        cellAddress: `${DEFAULT_SHEET_ID}!${address}`,
        localeId: currentFormulaLocaleId(),
      });
      if (value === cell.value) continue;
      this.#cells.set(address, { input: cell.input, value });
      changed = true;
    }

    if (!changed) return;
    this.#cellsVersion += 1;

    if (this.#activeCell && !this.formulaBar.isEditing) {
      const active = this.getCell(this.#activeCell);
      this.formulaBar.setActiveCell({ address: this.#activeCell, input: active.input, value: active.value });
    }
  }
}

function isAiFormula(input: string): boolean {
  const trimmed = input.trimStart();
  if (!trimmed.startsWith("=")) return false;
  const upper = trimmed.slice(1).trimStart().toUpperCase();
  return upper.startsWith("AI(") || upper.startsWith("AI.");
}
