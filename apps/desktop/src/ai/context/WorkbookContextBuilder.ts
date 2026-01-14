import type { DocumentController } from "../../document/documentController.js";
import { rangeToA1 as rangeToA1Selection } from "../../selection/a1.ts";
import type { Range } from "../../selection/types.ts";
import type { SheetNameResolver } from "../../sheet/sheetNameResolver.js";
import { formatSheetNameForA1 } from "../../sheet/formatSheetNameForA1.js";

import type { SheetSchema } from "../../../../../packages/ai-context/src/schema.js";
import { extractSheetSchema } from "../../../../../packages/ai-context/src/schema.js";
import {
  createHeuristicTokenEstimator,
  packSectionsToTokenBudget,
  stableJsonStringify,
  type TokenEstimator
} from "../../../../../packages/ai-context/src/tokenBudget.js";
import { rectToA1 } from "../../../../../packages/ai-rag/src/workbook/rect.js";
import { ToolExecutor } from "../../../../../packages/ai-tools/src/executor/tool-executor.ts";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.ts";
import type { CellScalar } from "../../../../../packages/ai-tools/src/spreadsheet/types.ts";

import type { ContextBudgetMode } from "../contextBudget.ts";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../contextBudget.ts";
import { computeDlpCacheKey } from "../dlp/dlpCacheKey.js";

export type WorkbookContextBlockKind = "selection" | "sheet_sample" | "retrieved";

export interface WorkbookContextDataBlock {
  kind: WorkbookContextBlockKind;
  sheetId: string;
  /**
   * A1 range with sheet prefix (Excel quoting rules).
   * Example: `Sheet1!A1:B10` or `'My Sheet'!A1`.
   */
  range: string;
  rowHeaders: string[];
  colHeaders: string[];
  values: CellScalar[][];
  /**
   * Present when the block could not be retrieved (DLP denied, tool error, etc).
   * The `values` are still safe placeholders (never raw restricted data).
   */
  error?: { code: string; message: string };
}

export interface WorkbookContextSheetSummary {
  sheetId: string;
  /**
   * Range (0-based) used to build the schema and `sheet_sample` data block.
   * Helps debugging when schema is built from a capped window of the used range.
   */
  analyzedRange?: { startRow: number; endRow: number; startCol: number; endCol: number };
  schema: SheetSchema;
}

export interface WorkbookContextBudgetInfo {
  mode: ContextBudgetMode;
  model: string;
  contextWindowTokens: number;
  reserveForOutputTokens: number;
  /**
   * Token budget used to pack the serialized prompt context returned from `build()`.
   */
  maxPromptContextTokens: number;
  /**
   * Estimated token usage of the final packed prompt context string.
   */
  usedPromptContextTokens: number;
}

export interface WorkbookContextBuildStats {
  /**
   * Monotonic-ish start time (uses `performance.now()` when available).
   * Only meant for comparing with `durationMs`.
   */
  startedAtMs: number;
  /**
   * Total wall-clock duration of `build()` in ms (same clock source as `startedAtMs`).
   */
  durationMs: number;
  /**
   * True when `build()` returned successfully.
   *
   * When false, the builder threw an error (e.g. DLP violation from the underlying
   * RAG service). In that case, other fields may be partially populated (reflecting
   * work done before the error was thrown).
   */
  ok: boolean;
  /**
   * Shallow error info captured when `ok` is false.
   *
   * Note: This intentionally excludes stack traces to keep this payload safe for
   * logging/telemetry by default.
   */
  error?: { name: string; message: string };
  mode: ContextBudgetMode;
  model: string;
  /**
   * Number of sheet summaries included in the final payload (active sheet first, plus
   * additional sheets depending on mode + `maxSheets`).
   */
  sheetCountSummarized: number;
  /**
   * Total number of sampled data blocks included in the final payload.
   * This counts blocks regardless of whether they came from cache.
   */
  blockCount: number;
  /**
   * Breakdown of `blockCount` by block kind.
   */
  blockCountByKind: Record<WorkbookContextBlockKind, number>;
  /**
   * Total number of cells in `payload.blocks` (sum of `values[row].length` over all rows/blocks).
   * This counts cells included in the payload, not the number of cells *read* from the tool.
   */
  blockCellCount: number;
  /**
   * Breakdown of `blockCellCount` by block kind.
   */
  blockCellCountByKind: Record<WorkbookContextBlockKind, number>;
  /**
   * Character length of the final `promptContext` string.
   */
  promptContextChars: number;
  /**
   * Estimated token count for the final `promptContext` string using the builder's estimator.
   * (Heuristic; model providers may tokenize differently.)
   */
  promptContextTokens: number;
  /**
   * Token budget used for packing prompt context sections.
   *
   * This is the `maxPromptContextTokens` from the builder's computed budget (after
   * applying any overrides).
   */
  promptContextBudgetTokens: number;
  /**
   * Number of times the prompt context packer inserted the standard
   * `"(trimmed to fit token budget)"` marker.
   *
   * This is an approximate indicator that one or more sections were truncated.
   */
  promptContextTrimmedSectionCount: number;
  /**
   * Number of actual `read_range` tool calls executed during the build (cache misses only).
   */
  readBlockCount: number;
  /**
   * Total number of cells requested across all `read_range` calls (cache misses only),
   * after clamping ranges to `maxRows`/`maxCols`.
   */
  readBlockCellCount: number;
  /**
   * Breakdown of `readBlockCount` by block kind.
   */
  readBlockCountByKind: Record<WorkbookContextBlockKind, number>;
  /**
   * Breakdown of `readBlockCellCount` by block kind.
   */
  readBlockCellCountByKind: Record<WorkbookContextBlockKind, number>;
  cache: {
    /**
     * `schema` refers to the per-sheet schema summary cache in `WorkbookContextBuilder`
     * (not any caching performed by the underlying `schemaProvider`).
     */
    schema: { hits: number; misses: number; entries: number };
    /**
     * `block` refers to the per-range sampled block cache in `WorkbookContextBuilder`.
     */
    block: {
      hits: number;
      misses: number;
      /**
       * Current number of cached blocks (at end of build).
       */
      entries: number;
      /**
       * Breakdown of cached blocks by kind.
       */
      entriesByKind: Record<WorkbookContextBlockKind, number>;
    };
  };
  rag: {
    enabled: boolean;
    retrievedCount: number;
    retrievedBlockCount: number;
  };
  timingsMs: {
    /**
     * Total time spent in the underlying RAG service, if enabled.
     */
    ragMs: number;
    /**
     * Total time spent extracting schemas from sampled values (`extractSheetSchema`).
     */
    schemaMs: number;
    /**
     * Total time spent executing `read_range` tool calls (cache misses only).
     */
    readBlockMs: number;
    /**
     * Total time spent serializing + packing prompt context sections.
     */
    promptContextMs: number;
    /**
     * Breakdown of `readBlockMs` by block kind.
     */
    readBlockMsByKind: Record<WorkbookContextBlockKind, number>;
  };
}

export interface WorkbookContextPayload {
  version: 1;
  workbookId: string;
  /**
   * Version counter for workbook-level schema metadata (named ranges / explicit tables),
   * as reported by the active `WorkbookSchemaProvider` (or 0 when unavailable).
   *
   * This is intended as a cheap cache key component for higher-level context caching
   * (e.g. avoiding hashing the full namedRanges/tables arrays).
   */
  schemaVersion: number;
  activeSheetId: string;
  /**
   * Optional selection passed by the caller (inline edit, attachments, etc).
   */
  selection?: { sheetId: string; range: string };
  sheets: WorkbookContextSheetSummary[];
  /**
   * Named ranges, when available from the host application.
   *
   * Desktop surfaces can provide them via `WorkbookContextBuilderOptions.schemaProvider`.
   */
  namedRanges: Array<{ name: string; range: string }>;
  /**
   * Structured tables are currently inferred from contiguous data regions.
   * The `SheetSchema.tables` list is the source of truth; this is a denormalized index.
   */
  tables: Array<{ sheetId: string; name: string; range: string }>;
  blocks: WorkbookContextDataBlock[];
  retrieval?: {
    query: string;
    retrievedChunkIds: string[];
    retrievedRanges: string[];
  };
  budget: WorkbookContextBudgetInfo;
}

export interface BuildWorkbookContextInput {
  activeSheetId: string;
  selectedRange?: { sheetId: string; range: Range };
  /**
   * Optional abort signal used to cancel context building work. When aborted, `build()`
   * will throw an AbortError (`error.name === "AbortError"`).
   */
  signal?: AbortSignal;
  /**
   * Optional DLP settings to apply for this build. This is passed through to the
   * tool executor (`read_range`) and to any RAG service used for retrieval.
   *
   * IMPORTANT: Callers may provide a fresh object on every build (e.g. `maybeGetAiCloudDlpOptions`),
   * so caching must never rely on object identity.
   */
  dlp?: any;
  /**
   * Optional focus question string used to drive semantic retrieval.
   * Surfaces should pass the user's latest prompt/request here.
   */
  focusQuestion?: string;
  /**
   * Optional attachments (chat). Kept generic to avoid coupling this module to chat types.
   */
  attachments?: unknown[];
}

export interface BuildWorkbookContextResult {
  payload: WorkbookContextPayload;
  /**
   * The packed prompt context string suitable for appending to a system prompt.
   * (Does not include the `WORKBOOK_CONTEXT:` prefix; callers can wrap as needed.)
   */
  promptContext: string;
  /**
   * Raw retrieved chunks returned by the underlying RAG service, if enabled.
   * This is returned separately so chat/agent surfaces can preserve existing telemetry
   * (retrieved ids/ranges, index stats, etc).
   */
  retrieved: unknown[];
  indexStats?: unknown;
}

export interface WorkbookSchemaProvider {
  /**
   * Optional schema version counter for workbook-level schema metadata (named ranges / tables).
   *
   * When provided, callers can cache the normalized metadata and avoid re-reading + re-sorting
   * the full list on every context build. Providers should increment the version whenever any
   * named range or table definition changes.
   *
   * When omitted, the version is treated as 0.
   */
  getSchemaVersion?: () => number;
  getNamedRanges?: () => Array<{ name: string; sheetId: string; range: Range }>;
  getTables?: () => Array<{ name: string; sheetId: string; range: Range }>;
}

type NormalizedWorkbookSchemaMetadata = {
  schemaVersion: number;
  /**
   * Cache key for the current sheet id -> display name mapping.
   *
   * This ensures we rebuild sheet-qualified metadata (named ranges / tables) when sheets
   * are renamed, even if the schemaProvider's schemaVersion does not change.
   */
  sheetNameKey: string;
  schemaNamedRangesBySheet: Map<string, Array<{ name: string; range: string }>>;
  schemaTablesBySheet: Map<string, Array<{ name: string; range: string }>>;
  namedRanges: Array<{ name: string; range: string }>;
  explicitTables: Array<{ sheetId: string; name: string; range: string }>;
};

// Shared cache across WorkbookContextBuilder instances.
//
// The desktop chat orchestrator constructs a new WorkbookContextBuilder per message. Keeping the
// schema metadata cache here lets us avoid re-reading and re-sorting named ranges / tables on
// every message as long as the provider's schemaVersion is unchanged.
const GLOBAL_SCHEMA_METADATA_CACHE = new WeakMap<WorkbookSchemaProvider, NormalizedWorkbookSchemaMetadata>();

export interface WorkbookContextBuilderOptions {
  workbookId: string;
  documentController: DocumentController;
  spreadsheet: SpreadsheetApi;
  /**
   * When enabled, the underlying ToolExecutor will treat formula cells as having a computed value
   * (via `cell.value`) instead of always treating them as `null`.
   *
   * This is opt-in because many backends (including the in-memory workbook) do not evaluate formulas,
   * and because computed formula values can be a DLP inference channel if dependencies are not traced.
   * ToolExecutor enforces conservative DLP gating (formula values are only surfaced when the selected
   * range decision is pure `ALLOW`).
   */
  includeFormulaValues?: boolean;
  /**
   * Optional resolver that maps stable sheet ids to user-facing display names
   * (and reverse).
   *
   * When provided, workbook context:
   * - Formats sheet-qualified A1 references using display names (Excel quoting rules).
   * - Resolves retrieved sheet names (from RAG) back to stable sheet ids for DocumentController reads.
   */
  sheetNameResolver?: SheetNameResolver | null;
  /**
   * Optional schema metadata provider for workbook-level objects that are not
   * currently exposed via the `SpreadsheetApi` tool surface (named ranges,
   * explicit table definitions, etc).
   */
  schemaProvider?: WorkbookSchemaProvider | null;
  /**
   * Optional RAG-capable provider (DesktopRagService or ContextManager).
   * Must implement `buildWorkbookContextFromSpreadsheetApi`.
   */
  ragService?: {
    buildWorkbookContextFromSpreadsheetApi(params: {
      spreadsheet: any;
      workbookId: string;
      query: string;
      attachments?: any[];
      topK?: number;
      includePromptContext?: boolean;
      dlp?: any;
      signal?: AbortSignal;
    }): Promise<any>;
  } | null;
  dlp?: any;
  mode: ContextBudgetMode;
  model: string;
  /**
   * Optional override for the model context window used for budgeting.
   * When omitted, defaults are derived from `mode` + `model` via `contextBudget.ts`.
   */
  contextWindowTokens?: number;
  /**
   * Optional override for output token reservation.
   * When omitted, defaults are derived from `mode` + `contextWindowTokens` via `contextBudget.ts`.
   */
  reserveForOutputTokens?: number;
  tokenEstimator?: TokenEstimator;
  /**
   * Hard cap on the number of sheets to include schema for (active sheet is always included).
   * Prevents runaway workbooks (100s of sheets) from blowing up context building time.
   */
  maxSheets?: number;
  /**
   * Maximum size of the sheet window used for schema extraction.
   */
  maxSchemaRows?: number;
  maxSchemaCols?: number;
  /**
   * Maximum size of sampled data blocks included in prompt context.
   */
  maxBlockRows?: number;
  maxBlockCols?: number;
  /**
   * Maximum number of retrieved chunks to turn into sampled data blocks.
   */
  maxRetrievedBlocks?: number;
  /**
   * Optional override for the token budget used to pack the final prompt context.
   * When omitted, a mode-based heuristic derived from `contextBudget.ts` is used.
   */
  maxPromptContextTokens?: number;
  /**
   * Optional perf instrumentation hook for workbook context building.
   * Default behavior is unchanged when this is not provided.
   */
  onBuildStats?: (stats: WorkbookContextBuildStats) => void;
}

export class WorkbookContextBuilder {
  private readonly options: Required<
    Omit<
      WorkbookContextBuilderOptions,
      | "ragService"
      | "dlp"
      | "sheetNameResolver"
      | "tokenEstimator"
      | "maxPromptContextTokens"
      | "contextWindowTokens"
      | "reserveForOutputTokens"
      | "schemaProvider"
      | "onBuildStats"
    >
  > &
    Pick<
      WorkbookContextBuilderOptions,
      | "ragService"
      | "dlp"
      | "sheetNameResolver"
      | "tokenEstimator"
      | "maxPromptContextTokens"
      | "contextWindowTokens"
      | "reserveForOutputTokens"
      | "schemaProvider"
      | "onBuildStats"
    >;
  private readonly estimator: TokenEstimator;

  private readonly sheetSummaryCache = new Map<string, { contentVersion: number; summary: WorkbookContextSheetSummary }>();
  private readonly maxSheetSummaryCacheEntries = 50;

  private readonly blockCache = new Map<string, { contentVersion: number; block: WorkbookContextDataBlock }>();
  private readonly maxBlockCacheEntries = 75;

  private cachedSchemaMetadata:
    | {
        provider: WorkbookSchemaProvider | null;
        schemaVersion: number;
        sheetNameKey: string;
        schemaNamedRangesBySheet: Map<string, Array<{ name: string; range: string }>>;
        schemaTablesBySheet: Map<string, Array<{ name: string; range: string }>>;
        namedRanges: Array<{ name: string; range: string }>;
        explicitTables: Array<{ sheetId: string; name: string; range: string }>;
      }
    | null = null;

  constructor(options: WorkbookContextBuilderOptions) {
    const isInlineEdit = options.mode === "inline_edit";
    this.options = {
      ...options,
      sheetNameResolver: options.sheetNameResolver ?? null,
      includeFormulaValues: options.includeFormulaValues ?? false,
      maxSheets: Math.max(1, options.maxSheets ?? 10),
      maxSchemaRows: Math.max(1, options.maxSchemaRows ?? (isInlineEdit ? 100 : 200)),
      maxSchemaCols: Math.max(1, options.maxSchemaCols ?? (isInlineEdit ? 30 : 50)),
      maxBlockRows: Math.max(1, options.maxBlockRows ?? (isInlineEdit ? 10 : 20)),
      maxBlockCols: Math.max(1, options.maxBlockCols ?? (isInlineEdit ? 10 : 20)),
      maxRetrievedBlocks: Math.max(0, options.maxRetrievedBlocks ?? 3),
    };
    this.estimator = options.tokenEstimator ?? createHeuristicTokenEstimator();
  }

  private getSheetContentVersion(sheetId: string): number {
    const controllerAny = this.options.documentController as any;
    try {
      const perSheet = controllerAny?.getSheetContentVersion?.(sheetId);
      if (typeof perSheet === "number" && Number.isFinite(perSheet)) return Math.trunc(perSheet);
    } catch {
      // Ignore and fall back to coarser versions.
    }

    const contentVersion = controllerAny?.contentVersion;
    if (typeof contentVersion === "number" && Number.isFinite(contentVersion)) return Math.trunc(contentVersion);

    const updateVersion = controllerAny?.updateVersion;
    if (typeof updateVersion === "number" && Number.isFinite(updateVersion)) return Math.trunc(updateVersion);

    return 0;
  }

  private sheetDisplayName(sheetId: string): string {
    const id = String(sheetId ?? "").trim();
    if (!id) return "";
    return this.options.sheetNameResolver?.getSheetNameById(id) ?? id;
  }

  private resolveSheetId(nameOrId: string): string {
    const raw = String(nameOrId ?? "").trim();
    if (!raw) return "";
    return this.options.sheetNameResolver?.getSheetIdByName(raw) ?? raw;
  }

  private computeSheetNameKey(): string {
    const resolver = this.options.sheetNameResolver ?? null;
    if (!resolver) return "no_sheet_name_resolver";
    const ids = this.listSheetIdsSortedUnique();
    const entries: Array<{ id: string; name: string }> = [];
    for (const id of ids) {
      let name = id;
      try {
        name = resolver.getSheetNameById(id) ?? id;
      } catch {
        name = id;
      }
      entries.push({ id, name });
    }
    const json = safeStableJsonStringify(entries);
    return `${json.length}:${hashString(json)}`;
  }

  async build(input: BuildWorkbookContextInput): Promise<BuildWorkbookContextResult> {
    const signal = input.signal;
    throwIfAborted(signal);

    const onBuildStats = this.options.onBuildStats;
    const stats: WorkbookContextBuildStats | null = onBuildStats
      ? {
          startedAtMs: nowMs(),
          durationMs: 0,
          ok: false,
          mode: this.options.mode,
          model: this.options.model,
          sheetCountSummarized: 0,
          blockCount: 0,
          blockCountByKind: { selection: 0, sheet_sample: 0, retrieved: 0 },
          blockCellCount: 0,
           blockCellCountByKind: { selection: 0, sheet_sample: 0, retrieved: 0 },
           promptContextChars: 0,
           promptContextTokens: 0,
           promptContextBudgetTokens: 0,
           promptContextTrimmedSectionCount: 0,
           readBlockCount: 0,
           readBlockCellCount: 0,
           readBlockCountByKind: { selection: 0, sheet_sample: 0, retrieved: 0 },
           readBlockCellCountByKind: { selection: 0, sheet_sample: 0, retrieved: 0 },
          cache: {
            schema: { hits: 0, misses: 0, entries: 0 },
            block: { hits: 0, misses: 0, entries: 0, entriesByKind: { selection: 0, sheet_sample: 0, retrieved: 0 } },
          },
          rag: { enabled: false, retrievedCount: 0, retrievedBlockCount: 0 },
          timingsMs: {
            ragMs: 0,
            schemaMs: 0,
            readBlockMs: 0,
            promptContextMs: 0,
            readBlockMsByKind: { selection: 0, sheet_sample: 0, retrieved: 0 },
          },
        }
      : null;
    try {
      const selection = input.selectedRange;
      const schemaProvider = this.options.schemaProvider ?? null;

      const dlp = input.dlp ?? this.options.dlp ?? undefined;
      const dlpCacheKey = computeDlpCacheKey(dlp);

      const schemaMetadata = this.getSchemaMetadata(schemaProvider);
      const schemaNamedRangesBySheet = schemaMetadata.schemaNamedRangesBySheet;
      const schemaTablesBySheet = schemaMetadata.schemaTablesBySheet;
      const namedRanges = schemaMetadata.namedRanges;

      // Reuse a single ToolExecutor per build() invocation to avoid repeated schema
      // validation + option normalization on the hot path. All `read_range` calls
      // use explicit sheet-prefixed A1 ranges, so `default_sheet` can be stable.
      // Align `default_sheet` with any provided DLP sheet_id so DLP evaluation for
      // cross-sheet reads uses the actual range sheet name.
      const maxReadRangeCells = Math.max(
        1,
        this.options.maxSchemaRows * this.options.maxSchemaCols,
        this.options.maxBlockRows * this.options.maxBlockCols,
      );
      const defaultSheetForTools =
        typeof (dlp as any)?.sheet_id === "string" && String((dlp as any).sheet_id).trim()
          ? String((dlp as any).sheet_id).trim()
          : input.activeSheetId;
      const executor = new ToolExecutor(this.options.spreadsheet, {
        default_sheet: defaultSheetForTools,
        sheet_name_resolver: this.options.sheetNameResolver ?? null,
        include_formula_values: this.options.includeFormulaValues,
        dlp,
        max_read_range_cells: maxReadRangeCells,
        max_read_range_chars: 200_000,
      });

      // Retrieval (semantic search) is optional and query-driven.
      const ragEnabled = Boolean(input.focusQuestion && this.options.ragService);
      if (stats) stats.rag.enabled = ragEnabled;
      let ragResult: any = null;
      if (ragEnabled) {
        throwIfAborted(signal);
        const startedAt = stats ? nowMs() : 0;
        try {
          ragResult = await this.options.ragService!.buildWorkbookContextFromSpreadsheetApi({
            spreadsheet: this.options.spreadsheet,
            workbookId: this.options.workbookId,
            query: input.focusQuestion!,
            attachments: input.attachments,
            // WorkbookContextBuilder builds its own promptContext; avoid redundant string formatting
            // + token estimation work inside the underlying RAG service.
            includePromptContext: false,
            dlp,
            signal,
          });
        } finally {
          if (stats) stats.timingsMs.ragMs += nowMs() - startedAt;
        }
        throwIfAborted(signal);
      }

      const retrieved: unknown[] = Array.isArray(ragResult?.retrieved) ? ragResult.retrieved : [];
      const indexStats = ragResult?.indexStats;
      if (stats) stats.rag.retrievedCount = retrieved.length;

      const retrievalEnabled = Boolean(input.focusQuestion && this.options.ragService);
      const retrievedSheetIds = retrievalEnabled
        ? extractRetrievedSheetIds(retrieved, this.options.sheetNameResolver ?? null)
        : [];

      const selectionBlock = selection
        ? (throwIfAborted(signal),
          await this.readBlock(
            executor,
            {
              dlpCacheKey,
              kind: "selection",
              sheetId: selection.sheetId,
              range: selection.range,
              // Selection-first: inline edit should prefer showing selection content even if large.
              maxRows: this.options.maxBlockRows,
              maxCols: this.options.maxBlockCols,
            },
            stats,
            signal,
          ))
        : null;
      throwIfAborted(signal);

      const sheetsToSummarize =
        this.options.mode === "inline_edit"
          ? [input.activeSheetId]
          : this.resolveSheetsToSummarize({
              activeSheetId: input.activeSheetId,
              // Selection prioritization only applies when retrieval is enabled (chat/agent surfaces).
              selectionSheetId: retrievalEnabled ? selection?.sheetId : null,
              retrievedSheetIds,
              // When retrieval yields relevant sheets, summarize only those (active/selection/retrieved) for latency.
              // When retrieval is disabled or yields no sheets, fall back to the legacy behavior of filling up to maxSheets.
              includeFallbackSheets: !retrievalEnabled || retrievedSheetIds.length === 0,
            });

      const sheetSummaries: WorkbookContextSheetSummary[] = [];
      const blocks: WorkbookContextDataBlock[] = [];
      if (selectionBlock) blocks.push(selectionBlock);

      for (const sheetId of sheetsToSummarize) {
        throwIfAborted(signal);
        const summary = await this.buildSheetSummary(
          executor,
          sheetId,
          {
            dlpCacheKey,
            schemaVersion: schemaMetadata.schemaVersion,
            namedRanges: schemaNamedRangesBySheet.get(sheetId),
            tables: schemaTablesBySheet.get(sheetId),
          },
          stats,
          signal,
        );
        sheetSummaries.push(summary);

        // Only include an explicit sample block for the active sheet (or for inline edit, the only sheet).
        const shouldIncludeSheetSample = sheetId === input.activeSheetId;
        if (shouldIncludeSheetSample && summary.analyzedRange) {
          throwIfAborted(signal);
          const sampleRange: Range = {
            startRow: summary.analyzedRange.startRow,
            endRow: summary.analyzedRange.endRow,
            startCol: summary.analyzedRange.startCol,
            endCol: summary.analyzedRange.endCol,
          };
          blocks.push(
            await this.buildSheetSampleBlock(
              executor,
              { dlpCacheKey, sheetId, analyzedRange: sampleRange },
              stats,
              signal,
            ),
          );
        }
      }
    throwIfAborted(signal);

    // Add sampled blocks for retrieved chunks (query-aware).
    if (retrieved.length && this.options.maxRetrievedBlocks > 0) {
      throwIfAborted(signal);
      const retrievedBlocks = await this.blocksFromRetrievedChunks(
        executor,
        retrieved,
        this.options.maxRetrievedBlocks,
        { dlpCacheKey },
        stats,
        signal,
      );
      blocks.push(...retrievedBlocks);
    }
    throwIfAborted(signal);

    // Stable ordering for snapshotting + deterministic prompts:
    // - Put the active sheet first (it tends to be the most relevant context).
    // - Order blocks by relevance (selection -> retrieved -> sheet sample) and then lexicographically.
    sheetSummaries.sort((a, b) => {
      if (a.sheetId === input.activeSheetId && b.sheetId !== input.activeSheetId) return -1;
      if (b.sheetId === input.activeSheetId && a.sheetId !== input.activeSheetId) return 1;
      return a.sheetId.localeCompare(b.sheetId);
    });
    const kindRank: Record<WorkbookContextBlockKind, number> = { selection: 0, retrieved: 1, sheet_sample: 2 };
    blocks.sort((a, b) => {
      const rank = (kindRank[a.kind] ?? 99) - (kindRank[b.kind] ?? 99);
      if (rank !== 0) return rank;
      const sheet = a.sheetId.localeCompare(b.sheetId);
      if (sheet !== 0) return sheet;
      return a.range.localeCompare(b.range);
    });

    const tablesByRange = new Map<string, { sheetId: string; name: string; range: string }>();
    for (const entry of sheetSummaries.flatMap((s) => s.schema.tables.map((t) => ({ sheetId: s.sheetId, name: t.name, range: t.range })))) {
      tablesByRange.set(`${entry.sheetId}\u0000${entry.range}`, entry);
    }
    for (const t of schemaMetadata.explicitTables) {
      tablesByRange.set(`${t.sheetId}\u0000${t.range}`, { sheetId: t.sheetId, name: t.name, range: t.range });
    }
    const tables = Array.from(tablesByRange.values()).sort((a, b) => {
      const ak = `${a.sheetId}\u0000${a.name}\u0000${a.range}`;
      const bk = `${b.sheetId}\u0000${b.name}\u0000${b.range}`;
      return ak.localeCompare(bk);
    });

      // Avoid spending time serializing large payloads when a caller already aborted.
      throwIfAborted(signal);
      const budget = this.computeBudget();

    const payload: WorkbookContextPayload = {
      version: 1,
      workbookId: this.options.workbookId,
      schemaVersion: schemaMetadata.schemaVersion,
      activeSheetId: input.activeSheetId,
      ...(selection
         ? { selection: { sheetId: selection.sheetId, range: this.rangeRef(selection.sheetId, selection.range) } }
         : {}),
      sheets: sheetSummaries,
      namedRanges,
      tables,
      blocks,
      ...(input.focusQuestion
        ? {
            retrieval: {
              query: input.focusQuestion,
              retrievedChunkIds: extractedRetrievedChunkIds(retrieved),
              retrievedRanges: extractRetrievedRanges(retrieved, this.options.sheetNameResolver ?? null),
            },
          }
        : {}),
      budget: {
        ...budget,
        usedPromptContextTokens: 0,
      },
    };

    const promptStartedAt = stats ? nowMs() : 0;
    throwIfAborted(signal);
    const promptContext = this.buildPromptContext({
      payload,
      ragResult,
      maxTokens: budget.maxPromptContextTokens,
    }, signal);
    if (stats) stats.timingsMs.promptContextMs += nowMs() - promptStartedAt;

    const usedPromptContextTokens = this.estimator.estimateTextTokens(promptContext);
    payload.budget.usedPromptContextTokens = usedPromptContextTokens;

    if (stats) {
      stats.durationMs = nowMs() - stats.startedAtMs;
      stats.ok = true;
      stats.sheetCountSummarized = sheetSummaries.length;
      stats.blockCount = blocks.length;
      const blockCountByKind: Record<WorkbookContextBlockKind, number> = { selection: 0, sheet_sample: 0, retrieved: 0 };
      const blockCellCountByKind: Record<WorkbookContextBlockKind, number> = { selection: 0, sheet_sample: 0, retrieved: 0 };
      let blockCellCount = 0;
      for (const block of blocks) {
        blockCountByKind[block.kind] += 1;
        const cells = countCellMatrix(block.values);
        blockCellCount += cells;
        blockCellCountByKind[block.kind] += cells;
      }
      stats.blockCountByKind = blockCountByKind;
      stats.blockCellCount = blockCellCount;
      stats.blockCellCountByKind = blockCellCountByKind;
      stats.cache.schema.entries = this.sheetSummaryCache.size;
      stats.cache.block.entries = this.blockCache.size;
      const cacheEntriesByKind: Record<WorkbookContextBlockKind, number> = { selection: 0, sheet_sample: 0, retrieved: 0 };
      for (const entry of this.blockCache.values()) {
        cacheEntriesByKind[entry.block.kind] += 1;
      }
      stats.cache.block.entriesByKind = cacheEntriesByKind;
      stats.promptContextChars = promptContext.length;
      stats.promptContextTokens = usedPromptContextTokens;
      stats.promptContextBudgetTokens = budget.maxPromptContextTokens;
      stats.promptContextTrimmedSectionCount = countSubstringOccurrences(promptContext, "trimmed to fit token budget");
      stats.rag.retrievedBlockCount = blocks.filter((b) => b.kind === "retrieved").length;
      try {
        onBuildStats?.(stats);
      } catch {
        // Ignore instrumentation failures; context building should be robust.
      }
    }

      return { payload, promptContext, retrieved, indexStats };
    } catch (error) {
      if (stats) {
        stats.durationMs = nowMs() - stats.startedAtMs;
        stats.ok = false;
        if (error instanceof Error) {
          stats.error = { name: error.name || "Error", message: error.message };
        } else {
          stats.error = { name: "Error", message: String(error) };
        }
        // Best-effort: expose current cache sizes even on failure.
        stats.cache.schema.entries = this.sheetSummaryCache.size;
        stats.cache.block.entries = this.blockCache.size;
        const cacheEntriesByKind: Record<WorkbookContextBlockKind, number> = { selection: 0, sheet_sample: 0, retrieved: 0 };
        for (const entry of this.blockCache.values()) {
          cacheEntriesByKind[entry.block.kind] += 1;
        }
        stats.cache.block.entriesByKind = cacheEntriesByKind;
        // Prompt budget is deterministic; report it even if the build fails before packing.
        try {
          stats.promptContextBudgetTokens = this.computeBudget().maxPromptContextTokens;
        } catch {
          // ignore
        }
        try {
          onBuildStats?.(stats);
        } catch {
          // Ignore instrumentation failures; build errors should still propagate.
        }
      }
      throw error;
    }
  }

  static serializePayload(payload: WorkbookContextPayload): string {
    // stableJsonStringify already sorts object keys deterministically.
    // Pretty-print for debugging + snapshot readability.
    const stable = stableJsonStringify(payload);
    try {
      return JSON.stringify(JSON.parse(stable), null, 2);
    } catch {
      return stable;
    }
  }

  private resolveSheetsToSummarize(params: {
    activeSheetId: string;
    selectionSheetId?: string | null;
    retrievedSheetIds: string[];
    includeFallbackSheets: boolean;
  }): string[] {
    const cap = Math.max(1, this.options.maxSheets);

    const candidates: string[] = [];
    const seen = new Set<string>();
    const pushUnique = (id: string | null | undefined) => {
      const normalized = typeof id === "string" ? id.trim() : "";
      if (!normalized) return;
      if (seen.has(normalized)) return;
      seen.add(normalized);
      candidates.push(normalized);
    };

    // Stable priority order:
    // 1) active sheet
    // 2) selection sheet (if any)
    // 3) sheets referenced by retrieved chunks
    // 4) optionally fill with remaining sheets (legacy behavior)
    pushUnique(params.activeSheetId);
    pushUnique(params.selectionSheetId);
    for (const id of params.retrievedSheetIds) pushUnique(id);

    if (params.includeFallbackSheets) {
      for (const id of this.listSheetIdsSortedUnique()) pushUnique(id);
    }

    // Cap, but always keep the active sheet included.
    const out = candidates.slice(0, cap);
    if (!out.includes(params.activeSheetId) && params.activeSheetId) {
      out.pop();
      out.unshift(params.activeSheetId);
    }
    return out;
  }

  private rememberSheetSummary(
    key: string,
    entry: { contentVersion: number; summary: WorkbookContextSheetSummary },
  ): void {
    // Bump recency when overwriting.
    if (this.sheetSummaryCache.has(key)) this.sheetSummaryCache.delete(key);
    this.sheetSummaryCache.set(key, entry);
    while (this.sheetSummaryCache.size > this.maxSheetSummaryCacheEntries) {
      const oldest = this.sheetSummaryCache.keys().next().value as string | undefined;
      if (!oldest) break;
      this.sheetSummaryCache.delete(oldest);
    }
  }

  private getReadBlockCacheKey(params: {
    dlpCacheKey: string;
    kind: WorkbookContextBlockKind;
    sheetId: string;
    range: Range;
    maxRows: number;
    maxCols: number;
  }): { key: string; clamped: Range; rangeRef: string } {
    const clamped = clampRange(params.range, { maxRows: params.maxRows, maxCols: params.maxCols });
    const rangeRef = this.rangeRef(params.sheetId, clamped);
    const key = `${params.dlpCacheKey}\u0000${params.kind}\u0000${params.sheetId}\u0000${rangeRef}`;
    return { key, clamped, rangeRef };
  }

  private getCachedReadBlock(params: {
    dlpCacheKey: string;
    kind: WorkbookContextBlockKind;
    sheetId: string;
    range: Range;
    maxRows: number;
    maxCols: number;
  }): WorkbookContextDataBlock | null {
    const { key } = this.getReadBlockCacheKey(params);
    const cached = this.blockCache.get(key);
    if (!cached) return null;

    const contentVersion = this.getSheetContentVersion(params.sheetId);
    if (cached.contentVersion !== contentVersion) {
      // Drop stale entries eagerly to keep the cache small.
      this.blockCache.delete(key);
      return null;
    }

    // Bump recency for eviction.
    this.blockCache.delete(key);
    this.blockCache.set(key, cached);
    return cached.block;
  }

  private rememberReadBlock(key: string, entry: { contentVersion: number; block: WorkbookContextDataBlock }): void {
    // Bump recency when overwriting.
    if (this.blockCache.has(key)) this.blockCache.delete(key);
    this.blockCache.set(key, entry);
    while (this.blockCache.size > this.maxBlockCacheEntries) {
      const oldest = this.blockCache.keys().next().value as string | undefined;
      if (!oldest) break;
      this.blockCache.delete(oldest);
    }
  }

  private listSheetIdsSortedUnique(): string[] {
    const ids = safeList(() => this.options.documentController.getSheetIds?.() ?? []);
    const unique = Array.from(new Set(ids.map((s) => String(s).trim()))).filter(Boolean);
    unique.sort((a, b) => a.localeCompare(b));
    return unique;
  }

  private async buildSheetSummary(
    executor: ToolExecutor,
    sheetId: string,
    extras: {
      dlpCacheKey: string;
      schemaVersion?: number;
      namedRanges?: Array<{ name: string; range: string }>;
      tables?: Array<{ name: string; range: string }>;
    },
    stats?: WorkbookContextBuildStats | null,
    signal?: AbortSignal,
  ): Promise<WorkbookContextSheetSummary> {
    throwIfAborted(signal);
    if (stats) stats.sheetCountSummarized += 1;

    const contentVersion = this.getSheetContentVersion(sheetId);
    const schemaVersion =
      typeof extras?.schemaVersion === "number" && Number.isFinite(extras.schemaVersion) ? Math.trunc(extras.schemaVersion) : 0;

    // Cache key should depend on schemaVersion when available (cheap invalidation signal).
    // Only fall back to hashing the full extras payload when schemaProvider cannot provide
    // a version counter (to avoid stale schemas when named ranges/tables change).
    const shouldHashExtras = Boolean(this.options.schemaProvider && !this.options.schemaProvider.getSchemaVersion);
    const cacheKeyParts = [
      sheetId,
      extras.dlpCacheKey,
      `schema:${this.options.maxSchemaRows}x${this.options.maxSchemaCols}`,
      `schemaVersion:${schemaVersion}`,
    ];
    if (shouldHashExtras) {
      const extrasJson = safeStableJsonStringify({
        namedRanges: extras?.namedRanges ?? [],
        tables: extras?.tables ?? [],
      });
      cacheKeyParts.push(`extras:${extrasJson.length}:${hashString(extrasJson)}`);
    }
    const cacheKey = cacheKeyParts.join("\u0000");

    const cached = this.sheetSummaryCache.get(cacheKey);
    if (cached && cached.contentVersion === contentVersion) {
      // Bump recency for eviction.
      this.sheetSummaryCache.delete(cacheKey);
      this.sheetSummaryCache.set(cacheKey, cached);
      if (stats) stats.cache.schema.hits += 1;
      return cached.summary;
    }
    if (stats) stats.cache.schema.misses += 1;

    const used = this.options.documentController.getUsedRange(sheetId);
    if (!used) {
      throwIfAborted(signal);
      const schema = extractSheetSchema({
        name: this.sheetDisplayName(sheetId) || sheetId,
        values: [],
        ...(extras?.namedRanges?.length ? { namedRanges: extras.namedRanges } : {}),
        ...(extras?.tables?.length ? { tables: extras.tables } : {}),
      } as any, { signal } as any);
      const summary = { sheetId, schema };
      throwIfAborted(signal);
      this.rememberSheetSummary(cacheKey, { contentVersion, summary });
      return summary;
    }

    const analyzedRange = clampRange(
      { startRow: used.startRow, endRow: used.endRow, startCol: used.startCol, endCol: used.endCol },
      { maxRows: this.options.maxSchemaRows, maxCols: this.options.maxSchemaCols },
    );

    throwIfAborted(signal);
    const block = await this.readBlock(
      executor,
      {
        dlpCacheKey: extras.dlpCacheKey,
        kind: "sheet_sample",
        sheetId,
        range: analyzedRange,
        maxRows: this.options.maxSchemaRows,
        maxCols: this.options.maxSchemaCols,
      },
      stats,
      signal,
    );
    throwIfAborted(signal);

    // If we couldn't read the sheet sample (DLP denied, runtime error, etc), do not attempt
    // schema extraction from placeholder values (it would create misleading fake tables).
    if (block.error) {
      throwIfAborted(signal);
      const schema = extractSheetSchema({
        name: this.sheetDisplayName(sheetId) || sheetId,
        values: [],
        ...(extras?.namedRanges?.length ? { namedRanges: extras.namedRanges } : {}),
        ...(extras?.tables?.length ? { tables: extras.tables } : {}),
      } as any, { signal } as any);
      const summary = { sheetId, analyzedRange, schema };
      throwIfAborted(signal);
      this.rememberSheetSummary(cacheKey, { contentVersion, summary });
      return summary;
    }

    const schemaValues: unknown[][] = block.values;
    throwIfAborted(signal);
    const schemaStart = stats ? nowMs() : 0;
    const schema = extractSheetSchema({
      name: this.sheetDisplayName(sheetId) || sheetId,
      values: schemaValues,
      // When schema is built from a capped window, preserve the original coordinates.
      origin: { row: analyzedRange.startRow, col: analyzedRange.startCol },
      ...(extras?.namedRanges?.length ? { namedRanges: extras.namedRanges } : {}),
      ...(extras?.tables?.length ? { tables: extras.tables } : {}),
    } as any, { signal } as any);
    if (stats) stats.timingsMs.schemaMs += nowMs() - schemaStart;

    const summary = { sheetId, analyzedRange, schema };
    throwIfAborted(signal);
    this.rememberSheetSummary(cacheKey, { contentVersion, summary });
    return summary;
  }

  private async buildSheetSampleBlock(
    executor: ToolExecutor,
    params: { dlpCacheKey: string; sheetId: string; analyzedRange: Range },
    stats?: WorkbookContextBuildStats | null,
    signal?: AbortSignal,
  ): Promise<WorkbookContextDataBlock> {
    throwIfAborted(signal);
    const sampleRange = clampRange(params.analyzedRange, { maxRows: this.options.maxBlockRows, maxCols: this.options.maxBlockCols });
    const sampleRangeRef = this.rangeRef(params.sheetId, sampleRange);

    // First try to reuse the (larger) schema extraction sample block, if it's cached.
    // This avoids an extra `read_range` call for the active sheet.
    const cachedSchemaSample = this.getCachedReadBlock({
      dlpCacheKey: params.dlpCacheKey,
      kind: "sheet_sample",
      sheetId: params.sheetId,
      range: params.analyzedRange,
      maxRows: this.options.maxSchemaRows,
      maxCols: this.options.maxSchemaCols,
    });

    if (cachedSchemaSample) {
      throwIfAborted(signal);
      if (stats) stats.cache.block.hits += 1;
      if (cachedSchemaSample.range === sampleRangeRef) return cachedSchemaSample;

      const derived = sliceBlock({
        block: cachedSchemaSample,
        kind: "sheet_sample",
        sheetId: params.sheetId,
        sourceRange: params.analyzedRange,
        targetRange: sampleRange,
        rangeRef: sampleRangeRef,
      });

      // Cache the derived prompt block under the smaller rangeRef key so subsequent builds can
      // reuse it even if the larger schema sample block is evicted.
      const { key } = this.getReadBlockCacheKey({
        dlpCacheKey: params.dlpCacheKey,
        kind: "sheet_sample",
        sheetId: params.sheetId,
        range: params.analyzedRange,
        maxRows: this.options.maxBlockRows,
        maxCols: this.options.maxBlockCols,
      });
      const contentVersion = this.getSheetContentVersion(params.sheetId);
      throwIfAborted(signal);
      this.rememberReadBlock(key, { contentVersion, block: derived });

      return derived;
    }

    // If the schema sample isn't cached (e.g. sheet summary served from cache but block cache evicted),
    // fall back to reading just the smaller prompt sample.
    return this.readBlock(
      executor,
      {
        dlpCacheKey: params.dlpCacheKey,
        kind: "sheet_sample",
        sheetId: params.sheetId,
        range: params.analyzedRange,
        maxRows: this.options.maxBlockRows,
        maxCols: this.options.maxBlockCols,
      },
      stats,
      signal,
    );
  }

  private async blocksFromRetrievedChunks(
    executor: ToolExecutor,
    retrieved: unknown[],
    maxBlocks: number,
    params: { dlpCacheKey: string },
    stats?: WorkbookContextBuildStats | null,
    signal?: AbortSignal,
  ): Promise<WorkbookContextDataBlock[]> {
    const out: WorkbookContextDataBlock[] = [];
    const seen = new Set<string>();

    for (const chunk of retrieved) {
      throwIfAborted(signal);
      if (out.length >= maxBlocks) break;
      const meta = (chunk as any)?.metadata;
      if (!meta) continue;
      const sheetName = typeof meta.sheetName === "string" ? meta.sheetName.trim() : "";
      const rect = meta.rect;
      if (!sheetName || !rect) continue;
      const sheetId = this.resolveSheetId(sheetName);
      if (!sheetId) continue;

      const range = rectToRange(rect);
      if (!range) continue;

      // Deduplicate retrieved reads by the actual (clamped) range we will read.
      const clamped = clampRange(range, { maxRows: this.options.maxBlockRows, maxCols: this.options.maxBlockCols });
      const dedupeKey = `${sheetId}\u0000${clamped.startRow},${clamped.startCol},${clamped.endRow},${clamped.endCol}`;
      if (seen.has(dedupeKey)) continue;
      seen.add(dedupeKey);

      out.push(
        await this.readBlock(
          executor,
          {
            dlpCacheKey: params.dlpCacheKey,
            kind: "retrieved",
            sheetId,
            range,
            maxRows: this.options.maxBlockRows,
            maxCols: this.options.maxBlockCols,
          },
          stats,
          signal,
        ),
      );
    }

    if (stats) stats.rag.retrievedBlockCount = out.length;
    return out;
  }

  private computeBudget(): Omit<WorkbookContextBudgetInfo, "usedPromptContextTokens"> {
    const contextWindowTokens =
      this.options.contextWindowTokens ?? getModeContextWindowTokens(this.options.mode, this.options.model);
    const reserveForOutputTokens =
      this.options.reserveForOutputTokens ?? getDefaultReserveForOutputTokens(this.options.mode, contextWindowTokens);

    const allowedPromptTokens = Math.max(0, contextWindowTokens - reserveForOutputTokens);
    const maxPromptContextTokens =
      this.options.maxPromptContextTokens ??
      clamp(
        Math.floor(
          allowedPromptTokens *
            (this.options.mode === "inline_edit" ? 0.25 : this.options.mode === "agent" ? 0.2 : 0.15),
        ),
        this.options.mode === "inline_edit" ? 256 : this.options.mode === "agent" ? 1024 : 512,
        this.options.mode === "inline_edit" ? 1024 : this.options.mode === "agent" ? 8192 : 4096,
      );

    return {
      mode: this.options.mode,
      model: this.options.model,
      contextWindowTokens,
      reserveForOutputTokens,
      maxPromptContextTokens,
    };
  }

  private buildPromptContext(
    params: { payload: WorkbookContextPayload; ragResult: any; maxTokens: number },
    signal?: AbortSignal,
  ): string {
    throwIfAborted(signal);
    const priorities =
      this.options.mode === "inline_edit"
        ? // Selection-first: prioritize data blocks over schemas for inline edit.
          { workbook_summary: 4, sheet_schemas: 3, data_blocks: 5, retrieved: 2, attachments: 1 }
        : this.options.mode === "agent"
          ? // Agents can benefit from deeper schemas + tables overview.
            { workbook_summary: 4, sheet_schemas: 5, data_blocks: 3, retrieved: 2, attachments: 1 }
          : // Chat default.
            { workbook_summary: 5, sheet_schemas: 4, data_blocks: 3, retrieved: 2, attachments: 1 };

    const sheets = params.payload.sheets.map((s) => ({ sheetId: s.sheetId, schema: s.schema }));
    const blocks = params.payload.blocks;

    const workbookSummary = {
      id: params.payload.workbookId,
      activeSheetId: params.payload.activeSheetId,
      sheets: params.payload.sheets.map((s) => s.sheetId).sort(),
      tables: params.payload.tables,
      namedRanges: params.payload.namedRanges,
      ...(params.payload.selection ? { selection: params.payload.selection } : {}),
    };

    const retrievedText =
      Array.isArray(params.ragResult?.retrieved) && params.ragResult.retrieved.length
        ? params.ragResult.retrieved
            .map((c: any) => (typeof c?.text === "string" ? c.text : ""))
            .filter(Boolean)
            .join("\n\n")
        : "";

    // `stableJsonStringify` already sorts object keys deterministically.
    // Keep the prompt context JSON compact (minified) for:
    // - Smaller prompts (more room for relevant workbook data).
    // - Easier machine parsing (systems can reliably scan for `"kind":"selection"`, etc).
    const stableCompactJson = (value: unknown): string => stableJsonStringify(value);

    const sections = [
      {
        key: "workbook_summary",
        priority: priorities.workbook_summary,
        text: `Workbook summary:\n${stableCompactJson(workbookSummary)}`,
      },
      {
        key: "sheet_schemas",
        priority: priorities.sheet_schemas,
        text: `Sheet schemas (schema-first):\n${stableCompactJson(sheets)}`,
      },
      {
        key: "data_blocks",
        priority: priorities.data_blocks,
        text: blocks.length ? `Sampled data blocks:\n${stableCompactJson(blocks)}` : "",
      },
      {
        key: "retrieved",
        priority: priorities.retrieved,
        text: retrievedText ? `Retrieved workbook context:\n${retrievedText}` : "",
      },
      {
        key: "attachments",
        priority: priorities.attachments,
        text:
          Array.isArray(params.ragResult?.attachments) && params.ragResult.attachments.length
            ? `Attachments:\n${stableCompactJson(params.ragResult.attachments)}`
            : "",
      },
    ].filter((s) => s.text);

    // Packing + token budgeting can be expensive for large workbooks; respect aborts.
    throwIfAborted(signal);
    const packed = packSectionsToTokenBudget(sections as any, Math.max(0, params.maxTokens), this.estimator, { signal });
    return packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n");
  }

  private async readBlock(
    executor: ToolExecutor,
    params: {
      dlpCacheKey: string;
      kind: WorkbookContextBlockKind;
      sheetId: string;
      range: Range;
      maxRows: number;
      maxCols: number;
    },
    stats?: WorkbookContextBuildStats | null,
    signal?: AbortSignal,
  ): Promise<WorkbookContextDataBlock> {
    throwIfAborted(signal);
    const contentVersion = this.getSheetContentVersion(params.sheetId);
    const { key: cacheKey, clamped, rangeRef } = this.getReadBlockCacheKey(params);
    const cached = this.blockCache.get(cacheKey);
    if (cached && cached.contentVersion === contentVersion) {
      // Bump recency for eviction.
      this.blockCache.delete(cacheKey);
      this.blockCache.set(cacheKey, cached);
      if (stats) stats.cache.block.hits += 1;
      return cached.block;
    }
    if (cached && cached.contentVersion !== contentVersion) {
      // Drop stale entries eagerly to keep the cache small.
      this.blockCache.delete(cacheKey);
    }

    if (stats) {
      stats.cache.block.misses += 1;
      stats.readBlockCount += 1;
      stats.readBlockCountByKind[params.kind] += 1;
      const cells = Math.max(0, (clamped.endRow - clamped.startRow + 1) * (clamped.endCol - clamped.startCol + 1));
      stats.readBlockCellCount += cells;
      stats.readBlockCellCountByKind[params.kind] += cells;
    }
    throwIfAborted(signal);
    const startedAt = stats ? nowMs() : 0;
    const toolResult = await executor.execute({
      name: "read_range",
      parameters: { range: rangeRef, include_formulas: false },
    } as any);
    throwIfAborted(signal);

    if (stats) {
      const elapsed = nowMs() - startedAt;
      stats.timingsMs.readBlockMs += elapsed;
      stats.timingsMs.readBlockMsByKind[params.kind] += elapsed;
    }

    if (toolResult.ok && toolResult.data && (toolResult.data as any).values) {
      const values = ((toolResult.data as any).values ?? []) as CellScalar[][];
      const block: WorkbookContextDataBlock = {
        kind: params.kind,
        sheetId: params.sheetId,
        range: rangeRef,
        rowHeaders: rowHeaders(clamped),
        colHeaders: colHeaders(clamped),
        values,
      };
      throwIfAborted(signal);
      this.rememberReadBlock(cacheKey, { contentVersion, block });
      return block;
    }

    // Never leak restricted content: fall back to a deterministic placeholder matrix.
    const error = toolResult.error ?? { code: "runtime_error", message: "Unknown error reading range." };
    const placeholder = String(error.code) === "permission_denied" ? "[POLICY_DENIED]" : "[UNAVAILABLE]";
    const block: WorkbookContextDataBlock = {
      kind: params.kind,
      sheetId: params.sheetId,
      range: rangeRef,
      rowHeaders: rowHeaders(clamped),
      colHeaders: colHeaders(clamped),
      values: [[placeholder]],
      error: { code: String(error.code ?? "runtime_error"), message: String(error.message ?? "Unknown error") },
    };
    throwIfAborted(signal);
    this.rememberReadBlock(cacheKey, { contentVersion, block });
    return block;
  }

  private rangeRef(sheetId: string, range: Range): string {
    const sheetName = this.sheetDisplayName(sheetId) || sheetId;
    return `${formatSheetNameForA1(sheetName)}!${rangeToA1Selection(range)}`;
  }

  private getSchemaProviderVersion(schemaProvider: WorkbookSchemaProvider | null): number {
    if (!schemaProvider?.getSchemaVersion) return 0;
    try {
      const v = schemaProvider.getSchemaVersion();
      return typeof v === "number" && Number.isFinite(v) ? Math.trunc(v) : 0;
    } catch {
      return 0;
    }
  }

  private getSchemaMetadata(schemaProvider: WorkbookSchemaProvider | null): NonNullable<typeof this.cachedSchemaMetadata> {
    const schemaVersion = this.getSchemaProviderVersion(schemaProvider);
    const sheetNameKey = this.computeSheetNameKey();
    const cached = this.cachedSchemaMetadata;
    if (
      cached &&
      cached.provider === schemaProvider &&
      cached.schemaVersion === schemaVersion &&
      cached.sheetNameKey === sheetNameKey
    ) {
      return cached;
    }
    if (
      cached &&
      (cached.provider !== schemaProvider || cached.schemaVersion !== schemaVersion || cached.sheetNameKey !== sheetNameKey)
    ) {
      // Sheet schemas incorporate named ranges / explicit tables. If workbook-level schema
      // metadata changes, invalidate the sheet summary cache even if workbook content
      // didn't change.
      this.sheetSummaryCache.clear();
    }

    // If this provider supports schemaVersion, consult a shared cross-builder cache first.
    // This keeps chat (which creates a new builder each message) from re-enumerating the
    // workbook's named ranges/tables on every message.
    if (schemaProvider?.getSchemaVersion) {
      const shared = GLOBAL_SCHEMA_METADATA_CACHE.get(schemaProvider);
      if (shared && shared.schemaVersion === schemaVersion && shared.sheetNameKey === sheetNameKey) {
        const out = { provider: schemaProvider, ...shared };
        this.cachedSchemaMetadata = out;
        return out;
      }
    }

    const namedRangeDefs = safeList(() => schemaProvider?.getNamedRanges?.() ?? []);
    const explicitTableDefs = safeList(() => schemaProvider?.getTables?.() ?? []);

    const schemaNamedRangesBySheet = new Map<string, Array<{ name: string; range: string }>>();
    for (const nr of namedRangeDefs) {
      const name = typeof nr?.name === "string" ? nr.name.trim() : "";
      const sheetId = typeof nr?.sheetId === "string" ? nr.sheetId : "";
      if (!name || !sheetId) continue;
      let range: string;
      try {
        range = this.schemaRangeRef(sheetId, nr.range);
      } catch {
        continue;
      }
      const list = schemaNamedRangesBySheet.get(sheetId) ?? [];
      list.push({ name, range });
      schemaNamedRangesBySheet.set(sheetId, list);
    }

    const schemaTablesBySheet = new Map<string, Array<{ name: string; range: string }>>();
    /** @type {Array<{ sheetId: string; name: string; range: string }>} */
    const explicitTables: Array<{ sheetId: string; name: string; range: string }> = [];
    for (const t of explicitTableDefs) {
      const name = typeof t?.name === "string" ? t.name.trim() : "";
      const sheetId = typeof t?.sheetId === "string" ? t.sheetId : "";
      if (!name || !sheetId) continue;
      let range: string;
      try {
        range = this.schemaRangeRef(sheetId, t.range);
      } catch {
        continue;
      }
      explicitTables.push({ sheetId, name, range });
      const list = schemaTablesBySheet.get(sheetId) ?? [];
      list.push({ name, range });
      schemaTablesBySheet.set(sheetId, list);
    }

    for (const [_sheetId, list] of schemaNamedRangesBySheet) {
      list.sort((a, b) => {
        const ak = `${a.name}\u0000${a.range}`;
        const bk = `${b.name}\u0000${b.range}`;
        return ak.localeCompare(bk);
      });
    }
    for (const [_sheetId, list] of schemaTablesBySheet) {
      list.sort((a, b) => {
        const ak = `${a.name}\u0000${a.range}`;
        const bk = `${b.name}\u0000${b.range}`;
        return ak.localeCompare(bk);
      });
    }

    const namedRanges = Array.from(schemaNamedRangesBySheet.entries())
      .flatMap(([sheetId, list]) => list.map((r) => ({ name: r.name, range: r.range })))
      .sort((a, b) => {
        const ak = `${a.name}\u0000${a.range}`;
        const bk = `${b.name}\u0000${b.range}`;
        return ak.localeCompare(bk);
      });

    const normalized: NormalizedWorkbookSchemaMetadata = {
      schemaVersion,
      sheetNameKey,
      schemaNamedRangesBySheet,
      schemaTablesBySheet,
      namedRanges,
      explicitTables,
    };

    // Only persist in the shared cache when the provider exposes schemaVersion. Otherwise, we
    // don't have a reliable invalidation mechanism and should avoid surprising cross-builder
    // staleness (e.g. tests swapping out provider results).
    if (schemaProvider?.getSchemaVersion) {
      GLOBAL_SCHEMA_METADATA_CACHE.set(schemaProvider, normalized);
    }

    const out = { provider: schemaProvider, ...normalized };
    this.cachedSchemaMetadata = out;
    return out;
  }

  private schemaRangeRef(sheetId: string, range: Range): string {
    const sheetName = this.sheetDisplayName(sheetId) || sheetId;
    return `${formatSheetNameForA1(sheetName)}!${rangeToA1Selection(range)}`;
  }
}

function createAbortError(message = "Aborted"): Error {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal?: AbortSignal): void {
  if (signal?.aborted) throw createAbortError();
}

function nowMs(): number {
  if (typeof performance !== "undefined" && typeof performance.now === "function") return performance.now();
  return Date.now();
}

function safeList<T>(fn: () => T[] | null | undefined): T[] {
  try {
    const out = fn();
    return Array.isArray(out) ? out : [];
  } catch {
    return [];
  }
}

function safeStableJsonStringify(value: unknown): string {
  try {
    return stableJsonStringify(value);
  } catch {
    try {
      return JSON.stringify(value) ?? "";
    } catch {
      return "";
    }
  }
}

function fnv1a32(value: string): number {
  // 32-bit FNV-1a hash. (Stable across runs.)
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return hash >>> 0;
}

function hashString(value: string): string {
  return fnv1a32(value).toString(16);
}

function sliceBlock(params: {
  block: WorkbookContextDataBlock;
  kind: WorkbookContextBlockKind;
  sheetId: string;
  sourceRange: Range;
  targetRange: Range;
  rangeRef: string;
}): WorkbookContextDataBlock {
  const { block, kind, sheetId, sourceRange, targetRange, rangeRef } = params;

  if (block.error) {
    return {
      kind,
      sheetId,
      range: rangeRef,
      rowHeaders: rowHeaders(targetRange),
      colHeaders: colHeaders(targetRange),
      values: block.values,
      error: block.error,
    };
  }

  const rowOffset = Math.max(0, targetRange.startRow - sourceRange.startRow);
  const colOffset = Math.max(0, targetRange.startCol - sourceRange.startCol);
  const rows = Math.max(1, targetRange.endRow - targetRange.startRow + 1);
  const cols = Math.max(1, targetRange.endCol - targetRange.startCol + 1);

  const values = Array.isArray(block.values)
    ? block.values
        .slice(rowOffset, rowOffset + rows)
        .map((row) => (Array.isArray(row) ? row.slice(colOffset, colOffset + cols) : []))
    : block.values;

  const fallback = Array.isArray(values) && values.length > 0 ? values : block.values;

  return {
    kind,
    sheetId,
    range: rangeRef,
    rowHeaders: rowHeaders(targetRange),
    colHeaders: colHeaders(targetRange),
    values: fallback,
  };
}

function clampRange(range: Range, limits: { maxRows: number; maxCols: number }): Range {
  const rows = Math.max(1, range.endRow - range.startRow + 1);
  const cols = Math.max(1, range.endCol - range.startCol + 1);
  const endRow = range.startRow + Math.min(rows, limits.maxRows) - 1;
  const endCol = range.startCol + Math.min(cols, limits.maxCols) - 1;
  return { startRow: range.startRow, endRow, startCol: range.startCol, endCol };
}

function clamp(value: number, min: number, max: number): number {
  return Math.max(min, Math.min(max, value));
}

function rowHeaders(range: Range): string[] {
  const rows = Math.max(0, range.endRow - range.startRow + 1);
  return Array.from({ length: rows }, (_v, idx) => String(range.startRow + idx + 1));
}

function colHeaders(range: Range): string[] {
  const cols = Math.max(0, range.endCol - range.startCol + 1);
  return Array.from({ length: cols }, (_v, idx) => columnIndexToA1(range.startCol + idx));
}

function countCellMatrix(values: CellScalar[][]): number {
  if (!Array.isArray(values)) return 0;
  let out = 0;
  for (const row of values) {
    if (Array.isArray(row)) out += row.length;
  }
  return out;
}

function countSubstringOccurrences(haystack: string, needle: string): number {
  if (!haystack || !needle) return 0;
  let count = 0;
  let idx = haystack.indexOf(needle);
  while (idx !== -1) {
    count += 1;
    idx = haystack.indexOf(needle, idx + needle.length);
  }
  return count;
}

function columnIndexToA1(columnIndex: number): string {
  // 0 -> A
  let n = columnIndex + 1;
  let letters = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function rectToRange(rect: any): Range | null {
  if (!rect || typeof rect !== "object") return null;
  const { r0, c0, r1, c1 } = rect as any;
  if (![r0, c0, r1, c1].every((n) => Number.isInteger(n) && n >= 0)) return null;
  return { startRow: r0, startCol: c0, endRow: r1, endCol: c1 };
}

function extractedRetrievedChunkIds(retrieved: unknown[]): string[] {
  return retrieved
    .map((c: any) => (typeof c?.id === "string" ? c.id : null))
    .filter(Boolean)
    .sort() as string[];
}

function extractRetrievedRanges(retrieved: unknown[], sheetNameResolver?: SheetNameResolver | null): string[] {
  const resolver = sheetNameResolver ?? null;
  const out: string[] = [];
  for (const chunk of retrieved) {
    const meta = (chunk as any)?.metadata;
    if (!meta) continue;
    const rawSheet = typeof meta.sheetName === "string" ? meta.sheetName.trim() : "";
    const rect = meta.rect;
    if (!rawSheet || !rect) continue;
    try {
      const range = rectToA1(rect);
      const sheetName = resolver?.getSheetNameById(rawSheet) ?? rawSheet;
      out.push(`${formatSheetNameForA1(sheetName)}!${range}`);
    } catch {
      // ignore malformed rect metadata
    }
  }
  return out.sort();
}

function extractRetrievedSheetIds(retrieved: unknown[], sheetNameResolver?: SheetNameResolver | null): string[] {
  const resolver = sheetNameResolver ?? null;
  const sheetIds: string[] = [];
  for (const chunk of retrieved) {
    const meta = (chunk as any)?.metadata;
    if (!meta) continue;
    const sheetName = typeof meta.sheetName === "string" ? meta.sheetName.trim() : "";
    if (!sheetName) continue;
    const resolved = resolver?.getSheetIdByName(sheetName) ?? sheetName;
    sheetIds.push(resolved);
  }
  const unique = Array.from(new Set(sheetIds));
  unique.sort((a, b) => a.localeCompare(b));
  return unique;
}
