import {
  CursorTabCompletionClient,
  TabCompletionEngine,
  type TabCompletionClient,
  type CompletionContext,
  type SchemaProvider,
  type Suggestion
} from "@formula/ai-completion";
import type { EngineClient } from "@formula/engine";
import { extractFormulaReferences, fromA1, type FormulaReferenceRange } from "@formula/spreadsheet-frontend";

import type { DocumentController } from "../../document/documentController.js";
import type { FormulaBarView } from "../../formula-bar/FormulaBarView.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";
import { evaluateFormula, type SpreadsheetValue } from "../../spreadsheet/evaluateFormula.js";
import { normalizeFormulaLocaleId } from "../../spreadsheet/formulaLocale.js";
import {
  createLocaleAwareFunctionRegistry,
  createLocaleAwarePartialFormulaParser,
  createLocaleAwareStarterFunctions,
} from "./parsePartialFormula.js";

function normalizeSheetNameToken(sheetName: string): string {
  const raw = String(sheetName ?? "").trim();
  const quoted = /^'((?:[^']|'')+)'$/.exec(raw);
  if (quoted) return quoted[1]!.replace(/''/g, "'").trim();
  return raw;
}

interface FormulaBarTabCompletionControllerOptions {
  formulaBar: FormulaBarView;
  document: DocumentController;
  getSheetId: () => string;
  limits?: { maxRows: number; maxCols: number };
  schemaProvider?: SchemaProvider | null;
  /**
   * Optional getter for workbook file metadata (directory + filename).
   *
   * This is used by preview evaluation so functions like `CELL("filename")` and
   * `INFO("directory")` can return accurate values while the user is typing.
   */
  getWorkbookFileMetadata?: (() => { directory: string | null; filename: string | null } | null) | null;
  /**
   * Optional getter for the WASM engine client (when available).
   *
   * Tab completion should remain responsive while the engine is booting; callers
   * should return `null` until the engine has finished initializing.
   */
  getEngineClient?: (() => EngineClient | null) | null;
  /**
   * Optional sheet name <-> id resolver.
   *
   * When provided, sheet-qualified references in completion previews (and the sheet list in the
   * schema provider) use the user-facing sheet display name while resolving back to stable ids
   * for DocumentController reads.
   */
  sheetNameResolver?: SheetNameResolver | null;
  /**
   * Optional Cursor backend completion client.
   *
   * Tests can inject a stub implementation to keep the suite network-safe and
   * deterministic.
   */
  completionClient?: TabCompletionClient | null;
}

export class FormulaBarTabCompletionController {
  readonly #completion: TabCompletionEngine;
  readonly #formulaBar: FormulaBarView;
  readonly #document: DocumentController;
  readonly #getSheetId: () => string;
  readonly #sheetNameResolver: SheetNameResolver | null;
  readonly #getWorkbookFileMetadata: (() => { directory: string | null; filename: string | null } | null) | null;
  readonly #limits: { maxRows: number; maxCols: number } | null;
  readonly #schemaProvider: SchemaProvider;

  #cellsVersion = 0;
  #completionRequest = 0;
  #pendingCompletion: Promise<void> | null = null;
  #abort: AbortController | null = null;
  #destroyed = false;

  readonly #unsubscribe: Array<() => void> = [];

  constructor(opts: FormulaBarTabCompletionControllerOptions) {
    this.#formulaBar = opts.formulaBar;
    this.#document = opts.document;
    this.#getSheetId = opts.getSheetId;
    this.#sheetNameResolver = opts.sheetNameResolver ?? null;
    this.#getWorkbookFileMetadata = typeof opts.getWorkbookFileMetadata === "function" ? opts.getWorkbookFileMetadata : null;
    this.#limits = opts.limits ?? null;

    const defaultSchemaProvider: SchemaProvider = {
      getSheetNames: () => {
        const ids = this.#document.getSheetIds();
        if (ids.length > 0) {
          if (!this.#sheetNameResolver) return ids;
          return ids.map((id) => this.#sheetNameResolver?.getSheetNameById(id) ?? id);
        }

        // DocumentController creates sheets lazily; fall back to the current sheet even if it
        // hasn't been materialized yet.
        const currentSheetId = this.#getSheetId();
        const fallbackName = this.#sheetNameResolver?.getSheetNameById(currentSheetId) ?? currentSheetId ?? "Sheet1";
        return [fallbackName];
      },
      getNamedRanges: () => [],
      getTables: () => [],
      // Include the sheet list in the cache key so suggestions refresh when new
      // sheets are created (DocumentController materializes sheets lazily).
      getCacheKey: () => {
        const ids = this.#document.getSheetIds();
        if (ids.length > 0) {
          if (!this.#sheetNameResolver) return ids.join("|");
          return ids.map((id) => this.#sheetNameResolver?.getSheetNameById(id) ?? id).join("|");
        }
        const currentSheetId = this.#getSheetId();
        return this.#sheetNameResolver?.getSheetNameById(currentSheetId) ?? currentSheetId ?? "Sheet1";
      },
    };

    const externalSchemaProvider = opts.schemaProvider ?? null;
    const schemaProvider: SchemaProvider = {
      getSheetNames: externalSchemaProvider?.getSheetNames ?? defaultSchemaProvider.getSheetNames,
      getNamedRanges: externalSchemaProvider?.getNamedRanges ?? defaultSchemaProvider.getNamedRanges,
      getTables: externalSchemaProvider?.getTables ?? defaultSchemaProvider.getTables,
       getCacheKey: () => {
         const base = defaultSchemaProvider.getCacheKey?.() ?? "";
         const extra = externalSchemaProvider?.getCacheKey?.() ?? "";
         // The partial parser is locale-aware (argument separators, etc). Include locale in the cache key so
         // tab completion recomputes suggestions if the UI locale changes at runtime.
         const locale = normalizeFormulaLocaleId(this.#formulaBar.currentLocaleId()) ?? "en-US";
         const combined = extra ? `${base}|${extra}` : base;
         return locale ? `${combined}|locale:${locale}` : combined;
       },
     };
    this.#schemaProvider = schemaProvider;

    const completionClient =
      opts.completionClient ?? createCursorTabCompletionClientFromEnv() ?? new CursorTabCompletionClient();
    const getLocaleId = () => this.#formulaBar.currentLocaleId();
    this.#completion = new TabCompletionEngine({
      functionRegistry: createLocaleAwareFunctionRegistry({ getLocaleId }),
      starterFunctions: createLocaleAwareStarterFunctions({ getLocaleId }),
      completionClient,
      schemaProvider,
      parsePartialFormula: createLocaleAwarePartialFormulaParser({
        getEngineClient: typeof opts.getEngineClient === "function" ? opts.getEngineClient : undefined,
        getLocaleId,
        timeoutMs: 10,
      }),
      // Keep room for a backend (Cursor) suggestion even when the rule-based engine
      // returns a full set of top-level starters (which otherwise consumes the
      // default `maxSuggestions=5` budget).
      maxSuggestions: 6,
      // Keep tab completion responsive even if the backend is slow/unavailable.
      completionTimeoutMs: 100,
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
    if (this.#destroyed) return;
    this.#destroyed = true;
    // Invalidate any in-flight request so stale results can't apply after teardown.
    this.#completionRequest += 1;
    try {
      this.#abort?.abort();
    } catch {
      // ignore
    }
    this.#abort = null;
    this.#pendingCompletion = null;
    try {
      this.#formulaBar.setAiSuggestion(null);
    } catch {
      // ignore
    }
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
    if (this.#destroyed) return;
    const model = this.#formulaBar.model;

    // Invalidate any in-flight request so stale results can't re-apply a ghost
    // after the user changes selection, commits, cancels, etc.
    const requestId = ++this.#completionRequest;
    if (this.#abort) {
      try {
        this.#abort.abort();
      } catch {
        // ignore
      }
      this.#abort = null;
    }

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
    const sheetNameResolver = this.#sheetNameResolver;
    const knownSheets =
      typeof this.#document.getSheetIds === "function"
        ? (this.#document.getSheetIds() as string[]).filter((s) => typeof s === "string" && s.length > 0)
        : [];
    const hasCurrentSheet = knownSheets.some((id) => id.toLowerCase() === sheetId.toLowerCase());

    const isKnownSheet = (id: string): boolean => {
      const candidate = String(id ?? "").trim();
      if (!candidate) return false;
      return knownSheets.some((known) => known.toLowerCase() === candidate.toLowerCase());
    };

    const resolveSheetId = (name: string): string | null => {
      const trimmed = normalizeSheetNameToken(name);
      if (!trimmed) return null;

      const resolved = sheetNameResolver?.getSheetIdByName(trimmed) ?? null;
      const candidate = resolved ?? trimmed;

      // Avoid creating phantom sheets during completion (DocumentController lazily
      // materializes sheets on read). If we don't have any known sheets yet,
      // don't perform any document reads during completion.
      if (candidate.toLowerCase() === sheetId.toLowerCase()) return hasCurrentSheet ? sheetId : null;
      if (knownSheets.length === 0) return null;
      return knownSheets.find((id) => id.toLowerCase() === candidate.toLowerCase()) ?? null;
    };

    const surroundingCells = {
      getCellValue: (row: number, col: number, sheetName?: string): unknown => {
        if (row < 0 || col < 0) return null;
        if (this.#limits && (row >= this.#limits.maxRows || col >= this.#limits.maxCols)) return null;

        const targetSheet =
          typeof sheetName === "string" && sheetName.length > 0 ? resolveSheetId(sheetName) : hasCurrentSheet ? sheetId : null;
        if (!targetSheet) return null;
        // Avoid creating phantom sheets during completion: `DocumentController.getCell()` lazily
        // materializes sheets on read. Prefer `peekCell()` which returns an empty cell state when
        // the sheet doesn't exist yet.
        let state: { value: unknown; formula: string | null };
        if (typeof (this.#document as any).peekCell === "function") {
          state = (this.#document as any).peekCell(targetSheet, { row, col }) as { value: unknown; formula: string | null };
        } else {
          if (!isKnownSheet(targetSheet)) return null;
          state = this.#document.getCell(targetSheet, { row, col }) as { value: unknown; formula: string | null };
        }
        if (state?.value != null) return state.value;
        if (typeof state?.formula === "string" && state.formula.length > 0) return state.formula;
        return null;
      },
      getCacheKey: () => `${sheetId}:${cellsVersion}`,
    };

    const abortController = typeof AbortController !== "undefined" ? new AbortController() : null;
    this.#abort = abortController;

    this.#pendingCompletion = this.#completion
      .getSuggestions(
        {
          currentInput: draft,
          cursorPosition: cursor,
          cellRef: activeCell,
          surroundingCells,
        },
        {
          previewEvaluator: createPreviewEvaluator({
            document: this.#document,
            sheetId,
            cellAddress: activeCell,
            localeId: this.#formulaBar.currentLocaleId(),
            schemaProvider: this.#schemaProvider,
            sheetNameResolver: this.#sheetNameResolver,
            getWorkbookFileMetadata: this.#getWorkbookFileMetadata,
          }),
          ...(abortController ? { signal: abortController.signal } : {}),
        },
      )
      .then((suggestions) => {
        if (this.#destroyed) return;
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
        if (this.#destroyed) return;
        if (requestId !== this.#completionRequest) return;
        this.#formulaBar.setAiSuggestion(null);
      });
  }
}

function createCursorTabCompletionClientFromEnv(): CursorTabCompletionClient | null {
  const viteUrl = readViteEnv("VITE_CURSOR_AI_COMPLETION_URL");
  const nodeUrl = readNodeEnv("CURSOR_AI_COMPLETION_URL");
  const raw = viteUrl ?? nodeUrl;
  if (!raw) return null;

  try {
    return new CursorTabCompletionClient({ baseUrl: raw, timeoutMs: 100 });
  } catch {
    return null;
  }
}

function readViteEnv(key: string): string | null {
  try {
    const env = (import.meta as any)?.env;
    const value = env?.[key];
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (trimmed) return trimmed;
    }
  } catch {
    // ignore
  }
  return null;
}

function readNodeEnv(key: string): string | null {
  try {
    const value = (globalThis as any).process?.env?.[key];
    if (typeof value === "string") {
      const trimmed = value.trim();
      if (trimmed) return trimmed;
    }
  } catch {
    // ignore
  }
  return null;
}

function createPreviewEvaluator(params: {
  document: DocumentController;
  sheetId: string;
  cellAddress: string;
  localeId?: string;
  schemaProvider?: SchemaProvider | null;
  sheetNameResolver?: SheetNameResolver | null;
  getWorkbookFileMetadata?: (() => { directory: string | null; filename: string | null } | null) | null;
}): (args: { suggestion: Suggestion; context: CompletionContext }) => unknown | Promise<unknown> {
  const { document, sheetId, cellAddress } = params;
  const schemaProvider = params.schemaProvider ?? null;
  const sheetNameResolver = params.sheetNameResolver ?? null;
  const getWorkbookFileMetadata = params.getWorkbookFileMetadata ?? null;

  const resolveSheetDisplayNameById = (id: string): string => {
    const resolved = sheetNameResolver?.getSheetNameById(id) ?? null;
    const resolvedTrimmed = typeof resolved === "string" ? resolved.trim() : "";
    if (resolvedTrimmed) return resolvedTrimmed;
    const meta = typeof (document as any)?.getSheetMeta === "function" ? (document as any).getSheetMeta(id) : null;
    const metaName = typeof meta?.name === "string" ? meta.name.trim() : "";
    return metaName || id;
  };

  // Hard cap on the number of cell reads we allow for preview. This keeps
  // completion responsive even when the suggested formula references a large
  // range.
  const MAX_CELL_READS = 5_000;

  let namedRangesPromise: Promise<Map<string, { sheetName: string | null; ref: string }>> | null = null;
  let tablesPromise: Promise<Map<string, TablePreviewInfo>> | null = null;
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

  const getTables = () => {
    if (!schemaProvider?.getTables) return Promise.resolve(new Map());
    if (!tablesPromise) {
      tablesPromise = Promise.resolve()
        .then(() => schemaProvider.getTables?.())
        .then((items) => (Array.isArray(items) ? items : []))
        .then((items) => {
          const map = new Map<string, TablePreviewInfo>();
          for (const item of items) {
            const name = typeof (item as any)?.name === "string" ? String((item as any).name) : "";
            if (!name) continue;

            const columns = Array.isArray((item as any)?.columns) ? (item as any).columns.map(String) : [];
            if (columns.length === 0) continue;

            const sheetName = typeof (item as any)?.sheetName === "string" ? String((item as any).sheetName) : null;
            const startRow = parseIntLike((item as any)?.startRow);
            const startCol = parseIntLike((item as any)?.startCol);
            const endRow = parseIntLike((item as any)?.endRow);
            const endCol = parseIntLike((item as any)?.endCol);

            map.set(name.trim().toUpperCase(), {
              sheetName: sheetName ? sheetName.trim() : null,
              startRow,
              startCol,
              endRow,
              endCol,
              columns,
            });
          }
          return map;
        })
        .catch(() => new Map());
    }
    return tablesPromise;
  };

  return async ({ suggestion }: { suggestion: Suggestion; context: CompletionContext }): Promise<unknown> => {
    const text = suggestion?.text ?? "";
    if (typeof text !== "string" || text.trim() === "") return undefined;

    const localeId = (() => {
      const raw =
        params.localeId ??
        (typeof document !== "undefined" ? document.documentElement?.lang : "") ??
        undefined;
      const trimmed = String(raw ?? "").trim();
      return trimmed || undefined;
    })();

    const workbookFileMetadata = typeof getWorkbookFileMetadata === "function" ? getWorkbookFileMetadata() : null;

    const namedRanges = await getNamedRanges();
    const tables = await getTables();
    const resolveNameToReference = (name: string): string | null => {
      if (!name) return null;
      const entry = namedRanges.get(name.trim().toUpperCase());
      if (!entry) return null;
      if (entry.sheetName) return `${entry.sheetName}!${entry.ref}`;
      return entry.ref;
    };

    const defaultSheetName = sheetNameResolver?.getSheetNameById(sheetId) ?? sheetId;
    const activeCellCoord = (() => {
      try {
        const parsed = fromA1(cellAddress);
        return { row: parsed.row0, col: parsed.col0 };
      } catch {
        return null;
      }
    })();
    const structuredTables = Array.from(tables.entries()).map(([key, table]) => ({
      name: key,
      columns: table.columns,
      sheetName: table.sheetName ?? undefined,
      startRow: table.startRow ?? undefined,
      startCol: table.startCol ?? undefined,
      endRow: table.endRow ?? undefined,
      endCol: table.endCol ?? undefined,
    }));

    const resolveThisRowStructuredRefRange = (refText: string): FormulaReferenceRange | null => {
      const trimmed = String(refText ?? "").trim();
      if (!trimmed) return null;
      // Structured refs are never sheet-qualified in formula text.
      if (!trimmed.includes("[") || trimmed.includes("!")) return null;
      if (!activeCellCoord) return null;

      // Only handle row-relative refs that require the edit row context.
      const lower = trimmed.toLowerCase();
      if (!(lower.includes("@") || lower.includes("#this row"))) return null;

      // Bracketed this-row shorthand forms like `[@[Col1]:[Col3]]` are equivalent to
      // `[[#This Row],[Col1]:[Col3]]`. Rewrite them so we can reuse the existing `#This Row`
      // multi-column resolver (which delegates column span parsing to the shared table resolver).
      const firstBracket = trimmed.indexOf("[");
      if (firstBracket >= 0) {
        const prefix = trimmed.slice(0, firstBracket);
        const suffix = trimmed.slice(firstBracket);
        if (suffix.startsWith("[@[") && suffix.endsWith("]]") && suffix.length > 5) {
          const columnExpr = suffix.slice(3, -2);
          const rewritten = `${prefix}[[#This Row],[${columnExpr}]]`;
          return resolveThisRowStructuredRefRange(rewritten);
        }
      }

      const unescapeStructuredRefItem = (value: string): string => value.replaceAll("]]", "]");
      const normalizeSelector = (value: string): string => value.trim().replace(/\s+/g, " ").toLowerCase();
      const normalizeSheetName = (value: string): string => normalizeSheetNameToken(value).toLowerCase();

      const currentSheetName = normalizeSheetName(defaultSheetName);

      const tableBounds = (table: TablePreviewInfo): { startRow: number; endRow: number; startCol: number; endCol: number } | null => {
        const startRow = table.startRow;
        const endRow = table.endRow;
        const startCol = table.startCol;
        const endCol = table.endCol;
        if (![startRow, endRow, startCol, endCol].every((v) => typeof v === "number" && Number.isFinite(v))) return null;
        return {
          startRow: Math.min(startRow!, endRow!),
          endRow: Math.max(startRow!, endRow!),
          startCol: Math.min(startCol!, endCol!),
          endCol: Math.max(startCol!, endCol!),
        };
      };

      const isActiveCellInTable = (table: TablePreviewInfo): boolean => {
        const bounds = tableBounds(table);
        if (!bounds) return false;
        // When the table has an explicit sheet name, ensure it matches the current sheet.
        const tableSheet = table.sheetName ? normalizeSheetName(table.sheetName) : currentSheetName;
        if (table.sheetName && tableSheet !== currentSheetName) return false;

        const dataStartRow = bounds.startRow + 1; // exclude header row
        if (activeCellCoord.row < dataStartRow || activeCellCoord.row > bounds.endRow) return false;
        if (activeCellCoord.col < bounds.startCol || activeCellCoord.col > bounds.endCol) return false;
        return true;
      };

      const findContainingTableName = (): string | null => {
        for (const [name, table] of tables.entries()) {
          if (!table) continue;
          if (!isActiveCellInTable(table)) continue;
          return name;
        }
        return null;
      };

      const resolveThisRowCell = (tableName: string, columnName: string): FormulaReferenceRange | null => {
        const key = String(tableName ?? "").trim().toUpperCase();
        if (!key) return null;
        const table = tables.get(key);
        if (!table) return null;
        if (!isActiveCellInTable(table)) return null;

        const bounds = tableBounds(table);
        if (!bounds) return null;

        const target = String(columnName ?? "").trim().toUpperCase();
        if (!target) return null;

        const colIndex = table.columns.findIndex((c) => String(c ?? "").trim().toUpperCase() === target);
        if (colIndex < 0) return null;
        const col = bounds.startCol + colIndex;
        if (col < bounds.startCol || col > bounds.endCol) return null;

        const sheet = typeof table.sheetName === "string" && table.sheetName.trim() ? table.sheetName.trim() : undefined;
        return {
          sheet,
          startRow: activeCellCoord.row,
          endRow: activeCellCoord.row,
          startCol: col,
          endCol: col,
        };
      };

      const resolveThisRowRow = (tableName: string): FormulaReferenceRange | null => {
        const key = String(tableName ?? "").trim().toUpperCase();
        if (!key) return null;
        const table = tables.get(key);
        if (!table) return null;
        if (!isActiveCellInTable(table)) return null;

        const bounds = tableBounds(table);
        if (!bounds) return null;

        const sheet = typeof table.sheetName === "string" && table.sheetName.trim() ? table.sheetName.trim() : undefined;
        return {
          sheet,
          startRow: activeCellCoord.row,
          endRow: activeCellCoord.row,
          startCol: bounds.startCol,
          endCol: bounds.endCol,
        };
      };

      const escapedItem = "((?:[^\\]]|\\]\\])+)"; // match non-] or escaped `]]`
      const qualifiedRe = new RegExp(
        `^([A-Za-z_][A-Za-z0-9_.]*)\\[\\[\\s*${escapedItem}\\s*\\]\\s*,\\s*\\[\\s*${escapedItem}\\s*\\]\\]$`,
        "i"
      );
      const qualifiedImplicitRe = new RegExp(`^\\[\\[\\s*${escapedItem}\\s*\\]\\s*,\\s*\\[\\s*${escapedItem}\\s*\\]\\]$`, "i");
      const atNestedRe = new RegExp(`^([A-Za-z_][A-Za-z0-9_.]*)\\[\\s*@\\s*\\[\\s*${escapedItem}\\s*\\]\\s*\\]$`, "i");
      const atNestedImplicitRe = new RegExp(`^\\[\\s*@\\s*\\[\\s*${escapedItem}\\s*\\]\\s*\\]$`, "i");
      const atRe = new RegExp(`^([A-Za-z_][A-Za-z0-9_.]*)\\[\\s*@\\s*${escapedItem}\\s*\\]$`, "i");
      const atImplicitRe = new RegExp(`^\\[\\s*@\\s*${escapedItem}\\s*\\]$`, "i");
      const atRowRe = new RegExp(`^([A-Za-z_][A-Za-z0-9_.]*)\\[\\s*@\\s*\\]$`, "i");
      const atRowImplicitRe = new RegExp(`^\\[\\s*@\\s*\\]$`, "i");

      const qualifiedMatch = qualifiedRe.exec(trimmed);
      if (qualifiedMatch) {
        const selector = normalizeSelector(unescapeStructuredRefItem(qualifiedMatch[2]!.trim()));
        if (selector !== "#this row") return null;
        const column = unescapeStructuredRefItem(qualifiedMatch[3]!.trim());
        return resolveThisRowCell(qualifiedMatch[1]!, column);
      }

      const qualifiedImplicitMatch = qualifiedImplicitRe.exec(trimmed);
      if (qualifiedImplicitMatch) {
        const selector = normalizeSelector(unescapeStructuredRefItem(qualifiedImplicitMatch[1]!.trim()));
        if (selector !== "#this row") return null;
        const column = unescapeStructuredRefItem(qualifiedImplicitMatch[2]!.trim());
        const tableName = findContainingTableName();
        if (!tableName) return null;
        return resolveThisRowCell(tableName, column);
      }

      const nestedMatch = atNestedRe.exec(trimmed);
      if (nestedMatch) {
        const column = unescapeStructuredRefItem(nestedMatch[2]!.trim());
        return resolveThisRowCell(nestedMatch[1]!, column);
      }

      const nestedImplicitMatch = atNestedImplicitRe.exec(trimmed);
      if (nestedImplicitMatch) {
        const column = unescapeStructuredRefItem(nestedImplicitMatch[1]!.trim());
        const tableName = findContainingTableName();
        if (!tableName) return null;
        return resolveThisRowCell(tableName, column);
      }

      const atMatch = atRe.exec(trimmed);
      if (atMatch) {
        const column = unescapeStructuredRefItem(atMatch[2]!.trim());
        // Shorthand `@Column` does not permit spaces/brackets; reject ambiguous matches like `@[Col]]...`.
        if (/[\s\[\],;\]]/.test(column)) return null;
        return resolveThisRowCell(atMatch[1]!, column);
      }

      const implicitAtMatch = atImplicitRe.exec(trimmed);
      if (implicitAtMatch) {
        const column = unescapeStructuredRefItem(implicitAtMatch[1]!.trim());
        if (/[\s\[\],;\]]/.test(column)) return null;
        const tableName = findContainingTableName();
        if (!tableName) return null;
        return resolveThisRowCell(tableName, column);
      }

      const atRowMatch = atRowRe.exec(trimmed);
      if (atRowMatch) {
        return resolveThisRowRow(atRowMatch[1]!);
      }

      const implicitAtRowMatch = atRowImplicitRe.exec(trimmed);
      if (implicitAtRowMatch) {
        const tableName = findContainingTableName();
        if (!tableName) return null;
        return resolveThisRowRow(tableName);
      }

      // Multi-column `#This Row` structured refs (`Table1[[#This Row],[Col1],[Col2]]`, etc).
      //
      // Delegate column-span parsing + contiguity checks to the shared structured-ref resolver
      // by rewriting the selector to `#All`, then clamp to the active edit row.
      const thisRowSelectorBracketRe = /\[\s*#\s*this\s+row\s*\]/i;
      if (thisRowSelectorBracketRe.test(trimmed)) {
        const firstBracket = trimmed.indexOf("[");
        if (firstBracket >= 0) {
          const suffix = trimmed.slice(firstBracket).trimStart();
          if (suffix.startsWith("[[")) {
            const explicitTableName = trimmed.slice(0, firstBracket).trim();
            const tableName = explicitTableName || findContainingTableName();
            if (!tableName) return null;

            const key = tableName.trim().toUpperCase();
            const table = tables.get(key);
            if (!table) return null;
            if (!isActiveCellInTable(table)) return null;

            const bounds = tableBounds(table);
            if (!bounds) return null;

            const fullRef = explicitTableName ? trimmed : `${tableName}${trimmed}`;
            const rewritten = fullRef.replace(/\[\s*#\s*this\s+row\s*\]/gi, "[#All]");
            const { references } = extractFormulaReferences(rewritten, undefined, undefined, { tables: structuredTables });
            const first = references[0];
            if (!first) return null;
            if (first.start !== 0 || first.end !== rewritten.length) return null;

            const r = first.range;
            if (r.startCol < bounds.startCol || r.endCol > bounds.endCol) return null;

            const sheet = typeof table.sheetName === "string" && table.sheetName.trim() ? table.sheetName.trim() : undefined;
            return {
              sheet,
              startRow: activeCellCoord.row,
              endRow: activeCellCoord.row,
              startCol: r.startCol,
              endCol: r.endCol,
            };
          }
        }
      }

      return null;
    };

    const resolveStructuredRefToReference = (refText: string): string | null => {
      const trimmed = String(refText ?? "").trim();
      // Structured refs are never sheet-qualified in formula text.
      if (!trimmed.includes("[") || trimmed.includes("!")) return null;

      // Implicit nested structured refs (no table prefix) are only valid inside a table context.
      // Infer the containing table based on the active edit cell and then delegate back to the
      // shared table resolver by rewriting the reference to a table-qualified form.
      if (trimmed.startsWith("[[")) {
        if (!activeCellCoord) return null;
        const normalizeSheetName = (value: string): string => normalizeSheetNameToken(value).toLowerCase();
        const currentSheetName = normalizeSheetName(defaultSheetName);

        const findContainingTableName = (): string | null => {
          for (const [name, table] of tables.entries()) {
            if (!table) continue;
            const startRow = table.startRow;
            const endRow = table.endRow;
            const startCol = table.startCol;
            const endCol = table.endCol;
            if (![startRow, endRow, startCol, endCol].every((v) => typeof v === "number" && Number.isFinite(v))) continue;

            const bounds = {
              startRow: Math.min(startRow!, endRow!),
              endRow: Math.max(startRow!, endRow!),
              startCol: Math.min(startCol!, endCol!),
              endCol: Math.max(startCol!, endCol!),
            };

            const tableSheet = table.sheetName ? normalizeSheetName(table.sheetName) : currentSheetName;
            if (table.sheetName && tableSheet !== currentSheetName) continue;

            const dataStartRow = bounds.startRow + 1; // exclude header row
            if (activeCellCoord.row < dataStartRow || activeCellCoord.row > bounds.endRow) continue;
            if (activeCellCoord.col < bounds.startCol || activeCellCoord.col > bounds.endCol) continue;
            return name;
          }
          return null;
        };

        const tableName = findContainingTableName();
        if (!tableName) return null;
        return resolveStructuredRefToReference(`${tableName}${trimmed}`);
      }

      const { references } = extractFormulaReferences(trimmed, undefined, undefined, {
        tables: structuredTables,
        resolveStructuredRef: resolveThisRowStructuredRefRange,
      });
      const first = references[0];
      if (!first) return null;
      if (first.start !== 0 || first.end !== trimmed.length) return null;

      const r = first.range;
      const sheet = typeof r.sheet === "string" && r.sheet.trim() ? r.sheet.trim() : defaultSheetName;
      const prefix = sheet ? formatSheetPrefix(sheet) : "";

      const start = toA1(r.startRow, r.startCol);
      const end = toA1(r.endRow, r.endCol);
      const a1 = start === end ? start : `${start}:${end}`;
      return `${prefix}${a1}`;
    };

    const knownSheets =
      typeof document.getSheetIds === "function"
        ? (document.getSheetIds() as string[]).filter((s) => typeof s === "string" && s.length > 0)
        : [];

    const resolveSheetId = (name: string): string | null => {
      const trimmed = normalizeSheetNameToken(name);
      if (!trimmed) return null;

      const resolved = sheetNameResolver?.getSheetIdByName(trimmed) ?? null;
      const candidate = resolved ?? trimmed;
      // Avoid creating phantom sheets during preview evaluation (DocumentController
      // lazily materializes sheets on read). If we don't have any known sheets
      // yet, only allow reads against the current sheet id.
      if (candidate.toLowerCase() === sheetId.toLowerCase()) return sheetId;
      if (knownSheets.length === 0) return null;

      return knownSheets.find((id) => id.toLowerCase() === candidate.toLowerCase()) ?? null;
    };

    let reads = 0;
    const memo = new Map<string, SpreadsheetValue>();
    const stack = new Set<string>();

    const getCellValue = (ref: string): SpreadsheetValue => {
      reads += 1;
      if (reads > MAX_CELL_READS) {
        throw new Error("preview too large");
      }

      const trimmed = ref.replaceAll("$", "").trim();
      let targetSheet = sheetId;
      let addr = trimmed;
      const bang = trimmed.indexOf("!");
      if (bang >= 0) {
        const sheetName = trimmed.slice(0, bang).trim();
        const resolved = resolveSheetId(sheetName);
        if (!resolved) return "#REF!";
        targetSheet = resolved;
        addr = trimmed.slice(bang + 1);
      }

      const normalized = addr.replaceAll("$", "").toUpperCase();
      const key = `${targetSheet}:${normalized}`;
      if (memo.has(key)) return memo.get(key) as SpreadsheetValue;
      if (stack.has(key)) return "#REF!";

      stack.add(key);
      const state =
        typeof (document as any).peekCell === "function"
          ? ((document as any).peekCell(targetSheet, normalized) as { value: unknown; formula: string | null })
          : (document.getCell(targetSheet, normalized) as { value: unknown; formula: string | null });
      let value: SpreadsheetValue;
      if (state?.formula) {
        value = evaluateFormula(state.formula, getCellValue, {
          cellAddress: `${targetSheet}!${normalized}`,
          resolveNameToReference,
          workbookFileMetadata,
          currentSheetName: resolveSheetDisplayNameById(targetSheet),
          localeId,
        });
      } else {
        const raw = state?.value ?? null;
        value = raw == null || typeof raw === "number" || typeof raw === "string" || typeof raw === "boolean" ? raw : null;
      }
      stack.delete(key);
      memo.set(key, value);
      return value;
    };

    try {
      const value = evaluateFormula(text, getCellValue, {
        cellAddress: `${sheetId}!${cellAddress}`,
        resolveNameToReference,
        resolveStructuredRefToReference,
        workbookFileMetadata,
        currentSheetName: resolveSheetDisplayNameById(sheetId),
        localeId,
      });
      // Errors from the lightweight evaluator usually mean unsupported syntax.
      if (typeof value === "string" && (value === "#NAME?" || value === "#VALUE!")) return "(preview unavailable)";
      return value;
    } catch {
      return "(preview unavailable)";
    }
  };
}

type TablePreviewInfo = {
  sheetName: string | null;
  startRow: number | null;
  startCol: number | null;
  endRow: number | null;
  endCol: number | null;
  columns: string[];
};

function parseIntLike(value: unknown): number | null {
  const n = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(n)) return null;
  return Math.trunc(n);
}

function formatSheetPrefix(sheetName: string): string {
  const trimmed = sheetName.trim();
  if (!trimmed) return "";
  const token = formatSheetNameForA1(trimmed);
  return token ? `${token}!` : "";
}

function toA1(row: number, col: number): string {
  let n = col + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return `${letters}${row + 1}`;
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
  let first: Suggestion | null = null;
  let bestBackend: Suggestion | null = null;

  for (const s of suggestions) {
    if (!s || typeof s.text !== "string") continue;
    if (s.text === draft) continue;
    if (!s.text.startsWith(prefix)) continue;
    if (suffix && !s.text.endsWith(suffix)) continue;

    const ghostLength = s.text.length - prefix.length - suffix.length;
    if (ghostLength <= 0) continue;

    // Ensure the suggested text actually represents an insertion at the caret.
    if (s.text.slice(cursor, s.text.length - suffix.length).length !== ghostLength) continue;

    if (!first) first = s;

    // Cursor backend suggestions are encoded as "formula" suggestions where
    // `displayText` equals the full suggested text. Prefer those over the
    // rule-based starter/function suggestions when available.
    const isBackendFormula = s.type === "formula" && s.displayText === s.text;
    if (isBackendFormula) {
      if (!bestBackend || (s.confidence ?? 0) > (bestBackend.confidence ?? 0)) {
        bestBackend = s;
      }
    }
  }

  return bestBackend ?? first;
}
