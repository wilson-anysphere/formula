import { TabCompletionEngine, type Suggestion } from "@formula/ai-completion";

import type { DocumentController } from "../../document/documentController.js";
import type { FormulaBarView } from "../../formula-bar/FormulaBarView.js";

export interface FormulaBarTabCompletionControllerOptions {
  formulaBar: FormulaBarView;
  document: DocumentController;
  getSheetId: () => string;
  limits?: { maxRows: number; maxCols: number };
}

export class FormulaBarTabCompletionController {
  readonly #completion = new TabCompletionEngine();
  readonly #formulaBar: FormulaBarView;
  readonly #document: DocumentController;
  readonly #getSheetId: () => string;
  readonly #limits: { maxRows: number; maxCols: number } | null;

  #cellsVersion = 0;
  #completionRequest = 0;
  #pendingCompletion: Promise<void> | null = null;

  readonly #unsubscribe: Array<() => void> = [];

  constructor(opts: FormulaBarTabCompletionControllerOptions) {
    this.#formulaBar = opts.formulaBar;
    this.#document = opts.document;
    this.#getSheetId = opts.getSheetId;
    this.#limits = opts.limits ?? null;

    const textarea = this.#formulaBar.textarea;

    const updateNow = () => this.update();
    const updateNextMicrotask = () => queueMicrotask(() => this.update());

    textarea.addEventListener("input", updateNow);
    textarea.addEventListener("click", updateNow);
    textarea.addEventListener("keyup", updateNow);
    textarea.addEventListener("focus", updateNow);
    textarea.addEventListener("keydown", updateNextMicrotask);

    this.#unsubscribe.push(() => textarea.removeEventListener("input", updateNow));
    this.#unsubscribe.push(() => textarea.removeEventListener("click", updateNow));
    this.#unsubscribe.push(() => textarea.removeEventListener("keyup", updateNow));
    this.#unsubscribe.push(() => textarea.removeEventListener("focus", updateNow));
    this.#unsubscribe.push(() => textarea.removeEventListener("keydown", updateNextMicrotask));

    const stopDocUpdates = this.#document.on?.("update", () => {
      this.#cellsVersion += 1;
      if (this.#formulaBar.isEditing()) {
        this.update();
      }
    });
    if (typeof stopDocUpdates === "function") {
      this.#unsubscribe.push(stopDocUpdates);
    }
  }

  destroy(): void {
    for (const stop of this.#unsubscribe.splice(0)) stop();
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

  update(): void {
    const model = this.#formulaBar.model;

    // Invalidate any in-flight request so stale results can't re-apply a ghost
    // after the user changes selection, commits, cancels, etc.
    const requestId = ++this.#completionRequest;

    if (!model.isEditing) {
      this.#formulaBar.setAiSuggestion(null);
      this.#pendingCompletion = null;
      return;
    }

    if (model.cursorStart !== model.cursorEnd) {
      this.#formulaBar.setAiSuggestion(null);
      this.#pendingCompletion = null;
      return;
    }

    const draft = model.draft;
    const cursor = model.cursorStart;
    const activeCell = model.activeCell.address;
    const sheetId = this.#getSheetId();
    const cellsVersion = this.#cellsVersion;

    const surroundingCells = {
      getCellValue: (row: number, col: number): unknown => {
        if (row < 0 || col < 0) return null;
        if (this.#limits && (row >= this.#limits.maxRows || col >= this.#limits.maxCols)) return null;

        const state = this.#document.getCell(sheetId, { row, col }) as { value: unknown; formula: string | null };
        if (state?.value != null) return state.value;
        if (typeof state?.formula === "string" && state.formula.length > 0) return state.formula;
        return null;
      },
      getCacheKey: () => `${sheetId}:${cellsVersion}`,
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
        if (!model.isEditing) return;
        if (model.cursorStart !== model.cursorEnd) return;
        if (model.draft !== draft) return;
        if (model.cursorStart !== cursor) return;
        if (model.activeCell.address !== activeCell) return;

        const prefix = draft.slice(0, cursor);
        const suffix = draft.slice(cursor);

        const best = bestPureInsertionSuggestion({ draft, cursor, prefix, suffix, suggestions });
        this.#formulaBar.setAiSuggestion(best ? best.text : null);
      })
      .catch(() => {
        if (requestId !== this.#completionRequest) return;
        this.#formulaBar.setAiSuggestion(null);
      });
  }
}

function bestPureInsertionSuggestion({
  draft,
  cursor,
  prefix,
  suffix,
  suggestions,
}: {
  draft: string;
  cursor: number;
  prefix: string;
  suffix: string;
  suggestions: Suggestion[];
}): Suggestion | null {
  for (const s of suggestions) {
    if (!s || typeof s.text !== "string") continue;
    if (s.text === draft) continue;
    if (!s.text.startsWith(prefix)) continue;
    if (suffix && !s.text.endsWith(suffix)) continue;

    const ghostLength = s.text.length - prefix.length - suffix.length;
    if (ghostLength <= 0) continue;

    // Ensure the suggested text actually represents an insertion at the caret.
    if (s.text.slice(cursor, s.text.length - suffix.length).length !== ghostLength) continue;

    return s;
  }

  return null;
}

