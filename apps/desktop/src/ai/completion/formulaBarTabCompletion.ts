import {
  CursorTabCompletionClient,
  TabCompletionEngine,
  type TabCompletionClient,
  type CompletionContext,
  type SchemaProvider,
  type Suggestion
} from "@formula/ai-completion";
import type { EngineClient } from "@formula/engine";

import type { DocumentController } from "../../document/documentController.js";
import type { FormulaBarView } from "../../formula-bar/FormulaBarView.js";
import { getLocale } from "../../i18n/index.js";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";
import { evaluateFormula, type SpreadsheetValue } from "../../spreadsheet/evaluateFormula.js";
import { createLocaleAwarePartialFormulaParser } from "./parsePartialFormula.js";

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
    this.#sheetNameResolver = opts.sheetNameResolver ?? null;
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
        const locale = getLocale();
        const combined = extra ? `${base}|${extra}` : base;
        return locale ? `${combined}|locale:${locale}` : combined;
      },
    };
    this.#schemaProvider = schemaProvider;

    const completionClient =
      opts.completionClient ?? createCursorTabCompletionClientFromEnv() ?? new CursorTabCompletionClient();
    this.#completion = new TabCompletionEngine({
      completionClient,
      schemaProvider,
      parsePartialFormula: createLocaleAwarePartialFormulaParser({
        getEngineClient: typeof opts.getEngineClient === "function" ? opts.getEngineClient : undefined,
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
            schemaProvider: this.#schemaProvider,
            sheetNameResolver: this.#sheetNameResolver,
          }),
        },
      )
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
  schemaProvider?: SchemaProvider | null;
  sheetNameResolver?: SheetNameResolver | null;
}): (args: { suggestion: Suggestion; context: CompletionContext }) => unknown | Promise<unknown> {
  const { document, sheetId, cellAddress } = params;
  const schemaProvider = params.schemaProvider ?? null;
  const sheetNameResolver = params.sheetNameResolver ?? null;

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
    const maybeRewrittenStructured = text.includes("[")
      ? rewriteStructuredReferences(text, tables, defaultSheetName)
      : null;
    const evalText = maybeRewrittenStructured ?? text;

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
      const value = evaluateFormula(evalText, getCellValue, {
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

function findColumnIndex(columns: string[], columnName: string): number | null {
  const target = columnName.trim().toUpperCase();
  if (!target) return null;
  for (let i = 0; i < columns.length; i += 1) {
    const col = String(columns[i] ?? "").trim();
    if (!col) continue;
    if (col.toUpperCase() === target) return i;
  }
  return null;
}

function resolveStructuredColumnRef(params: {
  tables: Map<string, TablePreviewInfo>;
  tableName: string;
  columnName: string;
  mode: "data" | "all";
  defaultSheetId: string;
}): string | null {
  const { tables, tableName, columnName, mode, defaultSheetId } = params;
  const table = tables.get(tableName.trim().toUpperCase());
  if (!table) return null;

  const colIdx = findColumnIndex(table.columns, columnName);
  if (colIdx == null) return null;

  if (
    table.startRow == null ||
    table.startCol == null ||
    table.endRow == null ||
    table.endCol == null ||
    table.startRow < 0 ||
    table.startCol < 0 ||
    table.endRow < 0 ||
    table.endCol < 0
  ) {
    return null;
  }

  // Table coordinates include the header row. For simple `Table[Column]` refs,
  // approximate Excel semantics by skipping the header row. (Totals rows are not
  // currently represented in `TableInfo`, so they remain included when present.)
  const startRow = mode === "all" ? table.startRow : table.startRow + 1;
  const endRow = table.endRow;
  if (startRow > endRow) return null;

  const col = table.startCol + colIdx;
  const start = toA1(startRow, col);
  const end = toA1(endRow, col);
  const range = start === end ? start : `${start}:${end}`;

  const sheet = table.sheetName ?? defaultSheetId;
  const prefix = sheet ? formatSheetPrefix(sheet) : "";
  return `${prefix}${range}`;
}

function rewriteStructuredReferences(
  formulaText: string,
  tables: Map<string, TablePreviewInfo>,
  defaultSheetId: string,
): string | null {
  let changed = false;
  let failed = false;

  // Match the same supported patterns as `parseStructuredReferenceText`:
  // - `TableName[ColumnName]`
  // - `TableName[[#All],[ColumnName]]`
  const allPattern = /([A-Za-z_][A-Za-z0-9_.]*)\[\[\s*#all\s*\]\s*,\s*\[\s*([^\]]+?)\s*\]\]/gi;
  const simplePattern = /([A-Za-z_][A-Za-z0-9_.]*)\[(?!\[)\s*([^\[\]]+?)\s*\]/gi;

  const rewriteSegment = (segment: string) => {
    let out = segment;

    out = out.replace(allPattern, (match, tableName, colName) => {
      const replacement = resolveStructuredColumnRef({
        tables,
        tableName,
        columnName: colName,
        mode: "all",
        defaultSheetId,
      });
      if (!replacement) {
        failed = true;
        return match;
      }
      changed = true;
      return replacement;
    });

    out = out.replace(simplePattern, (match, tableName, colName) => {
      const replacement = resolveStructuredColumnRef({
        tables,
        tableName,
        columnName: colName,
        mode: "data",
        defaultSheetId,
      });
      if (!replacement) {
        failed = true;
        return match;
      }
      changed = true;
      return replacement;
    });

    return out;
  };

  const rewritten = rewriteOutsideStrings(formulaText, rewriteSegment);
  if (!changed || failed) return null;
  return rewritten;
}

function rewriteOutsideStrings(text: string, transform: (segment: string) => string): string {
  let out = "";
  let segment = "";
  let inString = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    if (ch !== '"') {
      if (inString) out += ch;
      else segment += ch;
      continue;
    }

    if (!inString) {
      out += transform(segment);
      segment = "";
      inString = true;
      out += '"';
      continue;
    }

    // Escaped quote inside a string literal: "" -> "
    if (text[i + 1] === '"') {
      out += '""';
      i += 1;
      continue;
    }

    inString = false;
    out += '"';
  }

  out += transform(segment);
  return out;
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
