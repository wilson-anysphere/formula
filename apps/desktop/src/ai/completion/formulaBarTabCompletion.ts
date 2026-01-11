import {
  LocalModelManager,
  OllamaClient,
  TabCompletionEngine,
  type CompletionContext,
  type SchemaProvider,
  type Suggestion
} from "@formula/ai-completion";

import type { DocumentController } from "../../document/documentController.js";
import type { FormulaBarView } from "../../formula-bar/FormulaBarView.js";
import { evaluateFormula } from "../../spreadsheet/evaluateFormula.js";

export interface FormulaBarTabCompletionControllerOptions {
  formulaBar: FormulaBarView;
  document: DocumentController;
  getSheetId: () => string;
  limits?: { maxRows: number; maxCols: number };
  schemaProvider?: SchemaProvider | null;
}

const LOCAL_MODEL_ENABLED_KEY = "formula:aiCompletion:localModelEnabled";
const LOCAL_MODEL_NAME_KEY = "formula:aiCompletion:localModelName";
const LOCAL_MODEL_BASE_URL_KEY = "formula:aiCompletion:localModelBaseUrl";

export class FormulaBarTabCompletionController {
  readonly #completion: TabCompletionEngine;
  readonly #formulaBar: FormulaBarView;
  readonly #document: DocumentController;
  readonly #getSheetId: () => string;
  readonly #limits: { maxRows: number; maxCols: number } | null;
  readonly #schemaProvider: SchemaProvider;

  #cellsVersion = 0;
  #completionRequest = 0;
  #pendingCompletion: Promise<void> | null = null;

  readonly #unsubscribe: Array<() => void> = [];

  constructor(opts: FormulaBarTabCompletionControllerOptions) {
    this.#formulaBar = opts.formulaBar;
    this.#document = opts.document;
    this.#getSheetId = opts.getSheetId;
    this.#limits = opts.limits ?? null;

    const defaultSchemaProvider: SchemaProvider = {
      getSheetNames: () => {
        const ids = this.#document.getSheetIds();
        return ids.length > 0 ? ids : ["Sheet1"];
      },
      getNamedRanges: () => [],
      getTables: () => [],
      // Include the sheet list in the cache key so suggestions refresh when new
      // sheets are created (DocumentController materializes sheets lazily).
      getCacheKey: () => (this.#document.getSheetIds().length > 0 ? this.#document.getSheetIds().join("|") : "Sheet1"),
    };

    const externalSchemaProvider = opts.schemaProvider ?? null;
    const schemaProvider: SchemaProvider = {
      getSheetNames: externalSchemaProvider?.getSheetNames ?? defaultSchemaProvider.getSheetNames,
      getNamedRanges: externalSchemaProvider?.getNamedRanges ?? defaultSchemaProvider.getNamedRanges,
      getTables: externalSchemaProvider?.getTables ?? defaultSchemaProvider.getTables,
      getCacheKey: () => {
        const base = defaultSchemaProvider.getCacheKey?.() ?? "";
        const extra = externalSchemaProvider?.getCacheKey?.() ?? "";
        return extra ? `${base}|${extra}` : base;
      },
    };
    this.#schemaProvider = schemaProvider;

    const localModel = createLocalModelFromSettings();
    // Initialize local models opportunistically in the background (pulling models
    // can take time). Completion requests themselves are guarded by a strict
    // timeout in `TabCompletionEngine`.
    void localModel?.initialize?.().catch(() => {});

    this.#completion = new TabCompletionEngine({
      localModel,
      schemaProvider,
      // Keep tab completion responsive even if the local model is slow/unavailable.
      localModelTimeoutMs: 200,
    });

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
      getCellValue: (row: number, col: number, sheetName?: string): unknown => {
        if (row < 0 || col < 0) return null;
        if (this.#limits && (row >= this.#limits.maxRows || col >= this.#limits.maxCols)) return null;

        const targetSheet = typeof sheetName === "string" && sheetName.length > 0 ? sheetName : sheetId;
        const state = this.#document.getCell(targetSheet, { row, col }) as { value: unknown; formula: string | null };
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
      }, { previewEvaluator: createPreviewEvaluator({ document: this.#document, sheetId, cellAddress: activeCell, schemaProvider: this.#schemaProvider }) })
      .then((suggestions) => {
        if (requestId !== this.#completionRequest) return;
        if (this.#getSheetId() !== sheetId) return;
        if (!model.isEditing) return;
        if (model.cursorStart !== model.cursorEnd) return;
        if (model.draft !== draft) return;
        if (model.cursorStart !== cursor) return;
        if (model.activeCell.address !== activeCell) return;

        const prefix = draft.slice(0, cursor);
        const suffix = draft.slice(cursor);

        const best = bestPureInsertionSuggestion({ draft, cursor, prefix, suffix, suggestions });
        this.#formulaBar.setAiSuggestion(best ? { text: best.text, preview: best.preview } : null);
      })
      .catch(() => {
        if (requestId !== this.#completionRequest) return;
        this.#formulaBar.setAiSuggestion(null);
      });
  }
}

function readLocalStorage(key: string): string | null {
  try {
    const raw = globalThis.localStorage?.getItem(key);
    return raw == null ? null : raw;
  } catch {
    return null;
  }
}

function localStorageFlagEnabled(key: string): boolean {
  const raw = readLocalStorage(key);
  if (!raw) return false;
  const normalized = raw.trim().toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "on";
}

function createLocalModelFromSettings(): LocalModelManager | null {
  if (!localStorageFlagEnabled(LOCAL_MODEL_ENABLED_KEY)) return null;

  const modelName = readLocalStorage(LOCAL_MODEL_NAME_KEY) ?? "formula-completion";
  const baseUrl = readLocalStorage(LOCAL_MODEL_BASE_URL_KEY) ?? "http://localhost:11434";

  try {
    const ollamaClient = new OllamaClient({ baseUrl, timeoutMs: 2_000 });
    return new LocalModelManager({
      ollamaClient,
      requiredModels: [modelName],
      defaultModel: modelName,
      cacheSize: 200,
    });
  } catch {
    return null;
  }
}

function createPreviewEvaluator(params: {
  document: DocumentController;
  sheetId: string;
  cellAddress: string;
  schemaProvider?: SchemaProvider | null;
}): (args: { suggestion: Suggestion; context: CompletionContext }) => unknown | Promise<unknown> {
  const { document, sheetId, cellAddress } = params;
  const schemaProvider = params.schemaProvider ?? null;

  // Hard cap on the number of cell reads we allow for preview. This keeps
  // completion responsive even when the suggested formula references a large
  // range.
  const MAX_CELL_READS = 5_000;

  let namedRangesPromise: Promise<Map<string, { sheetName: string | null; ref: string }>> | null = null;
  const getNamedRanges = () => {
    if (!schemaProvider?.getNamedRanges) return Promise.resolve(new Map());
    if (!namedRangesPromise) {
      namedRangesPromise = Promise.resolve()
        .then(() => schemaProvider.getNamedRanges?.())
        .then((items) => (Array.isArray(items) ? items : []))
        .then((items) => {
          const map = new Map<string, { sheetName: string | null; ref: string }>();
          for (const item of items) {
            const name = typeof (item as any)?.name === "string" ? String((item as any).name) : "";
            const range = typeof (item as any)?.range === "string" ? String((item as any).range) : "";
            if (!name || !range) continue;

            const bang = range.indexOf("!");
            const sheetName = bang >= 0 ? range.slice(0, bang) : null;
            const ref = bang >= 0 ? range.slice(bang + 1) : range;
            if (!ref) continue;
            map.set(name.trim().toUpperCase(), { sheetName: sheetName ? sheetName.trim() : null, ref: ref.trim() });
          }
          return map;
        })
        .catch(() => new Map());
    }
    return namedRangesPromise;
  };

  return async ({ suggestion }: { suggestion: Suggestion; context: CompletionContext }): Promise<unknown> => {
    const text = suggestion?.text ?? "";
    if (typeof text !== "string" || text.trim() === "") return undefined;

    // The lightweight evaluator can't resolve sheet-qualified references or
    // structured references yet; don't show misleading errors.
    if (text.includes("!") || text.includes("[")) return "(preview unavailable)";

    const namedRanges = await getNamedRanges();
    const resolveNameToReference = (name: string): string | null => {
      if (!name) return null;
      const entry = namedRanges.get(name.trim().toUpperCase());
      if (!entry) return null;
      if (entry.sheetName && entry.sheetName.toLowerCase() !== sheetId.toLowerCase()) return null;
      return entry.ref;
    };

    let reads = 0;
    const memo = new Map<string, unknown>();
    const stack = new Set<string>();

    const getCellValue = (ref: string): unknown => {
      reads += 1;
      if (reads > MAX_CELL_READS) {
        throw new Error("preview too large");
      }

      const normalized = ref.replaceAll("$", "").toUpperCase();
      const key = `${sheetId}:${normalized}`;
      if (memo.has(key)) return memo.get(key) as unknown;
      if (stack.has(key)) return "#REF!";

      stack.add(key);
      const state = document.getCell(sheetId, normalized) as { value: unknown; formula: string | null };
      let value: unknown;
      if (state?.formula) {
        value = evaluateFormula(state.formula, getCellValue, {
          cellAddress: `${sheetId}!${normalized}`,
          resolveNameToReference,
        });
      } else {
        value = state?.value ?? null;
      }
      stack.delete(key);
      memo.set(key, value);
      return value;
    };

    try {
      const value = evaluateFormula(text, getCellValue, {
        cellAddress: `${sheetId}!${cellAddress}`,
        resolveNameToReference,
      });
      // Errors from the lightweight evaluator usually mean unsupported syntax.
      if (typeof value === "string" && (value === "#NAME?" || value === "#VALUE!")) return "(preview unavailable)";
      return value;
    } catch {
      return "(preview unavailable)";
    }
  };
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
