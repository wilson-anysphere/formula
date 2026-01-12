import type { DocumentController } from "../../document/documentController.js";
import { rangeToA1 as rangeToA1Selection } from "../../selection/a1";
import type { Range } from "../../selection/types";

import type { SheetSchema } from "../../../../../packages/ai-context/src/schema.js";
import { extractSheetSchema } from "../../../../../packages/ai-context/src/schema.js";
import {
  createHeuristicTokenEstimator,
  packSectionsToTokenBudget,
  stableJsonStringify,
  type TokenEstimator
} from "../../../../../packages/ai-context/src/tokenBudget.js";
import { rectToA1 } from "../../../../../packages/ai-rag/src/workbook/rect.js";
import { ToolExecutor } from "../../../../../packages/ai-tools/src/executor/tool-executor.js";
import type { SpreadsheetApi } from "../../../../../packages/ai-tools/src/spreadsheet/api.js";
import type { CellScalar } from "../../../../../packages/ai-tools/src/spreadsheet/types.js";

import type { ContextBudgetMode } from "../contextBudget.js";
import { getDefaultReserveForOutputTokens, getModeContextWindowTokens } from "../contextBudget.js";

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

export interface WorkbookContextPayload {
  version: 1;
  workbookId: string;
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
  getNamedRanges?: () => Array<{ name: string; sheetId: string; range: Range }>;
  getTables?: () => Array<{ name: string; sheetId: string; range: Range }>;
}

export interface WorkbookContextBuilderOptions {
  workbookId: string;
  documentController: DocumentController;
  spreadsheet: SpreadsheetApi;
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
      dlp?: any;
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
}

export class WorkbookContextBuilder {
  private readonly options: Required<
    Omit<
      WorkbookContextBuilderOptions,
      | "ragService"
      | "dlp"
      | "tokenEstimator"
      | "maxPromptContextTokens"
      | "contextWindowTokens"
      | "reserveForOutputTokens"
      | "schemaProvider"
    >
  > &
    Pick<
      WorkbookContextBuilderOptions,
      "ragService" | "dlp" | "tokenEstimator" | "maxPromptContextTokens" | "contextWindowTokens" | "reserveForOutputTokens" | "schemaProvider"
    >;
  private readonly estimator: TokenEstimator;

  constructor(options: WorkbookContextBuilderOptions) {
    const isInlineEdit = options.mode === "inline_edit";
    this.options = {
      ...options,
      maxSheets: Math.max(1, options.maxSheets ?? 10),
      maxSchemaRows: Math.max(1, options.maxSchemaRows ?? (isInlineEdit ? 100 : 200)),
      maxSchemaCols: Math.max(1, options.maxSchemaCols ?? (isInlineEdit ? 30 : 50)),
      maxBlockRows: Math.max(1, options.maxBlockRows ?? (isInlineEdit ? 10 : 20)),
      maxBlockCols: Math.max(1, options.maxBlockCols ?? (isInlineEdit ? 10 : 20)),
      maxRetrievedBlocks: Math.max(0, options.maxRetrievedBlocks ?? 3),
    };
    this.estimator = options.tokenEstimator ?? createHeuristicTokenEstimator();
  }

  async build(input: BuildWorkbookContextInput): Promise<BuildWorkbookContextResult> {
    const sheetIds =
      this.options.mode === "inline_edit" ? [input.activeSheetId] : this.resolveSheetIds({ activeSheetId: input.activeSheetId });

    const selection = input.selectedRange;
    const schemaProvider = this.options.schemaProvider ?? null;

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

    // Retrieval (semantic search) is optional and query-driven.
    const ragResult =
      input.focusQuestion && this.options.ragService
        ? await this.options.ragService.buildWorkbookContextFromSpreadsheetApi({
            spreadsheet: this.options.spreadsheet,
            workbookId: this.options.workbookId,
            query: input.focusQuestion,
            attachments: input.attachments,
            dlp: this.options.dlp,
          })
        : null;

    const retrieved: unknown[] = Array.isArray(ragResult?.retrieved) ? ragResult.retrieved : [];
    const indexStats = ragResult?.indexStats;

    const selectionBlock = selection
      ? await this.readBlock({
          kind: "selection",
          sheetId: selection.sheetId,
          range: selection.range,
          // Selection-first: inline edit should prefer showing selection content even if large.
          maxRows: this.options.maxBlockRows,
          maxCols: this.options.maxBlockCols,
        })
      : null;

    const sheetsToSummarize =
      this.options.mode === "inline_edit"
        ? [input.activeSheetId]
        : [input.activeSheetId, ...sheetIds.filter((id) => id !== input.activeSheetId)];

    const sheetSummaries: WorkbookContextSheetSummary[] = [];
    const blocks: WorkbookContextDataBlock[] = [];
    if (selectionBlock) blocks.push(selectionBlock);

    for (const sheetId of sheetsToSummarize) {
      const summary = await this.buildSheetSummary(sheetId, {
        namedRanges: schemaNamedRangesBySheet.get(sheetId),
        tables: schemaTablesBySheet.get(sheetId),
      });
      sheetSummaries.push(summary);

      // Only include an explicit sample block for the active sheet (or for inline edit, the only sheet).
      const shouldIncludeSheetSample = sheetId === input.activeSheetId;
      if (shouldIncludeSheetSample && summary.analyzedRange) {
        const sampleRange: Range = {
          startRow: summary.analyzedRange.startRow,
          endRow: summary.analyzedRange.endRow,
          startCol: summary.analyzedRange.startCol,
          endCol: summary.analyzedRange.endCol,
        };
        blocks.push(
          await this.readBlock({
            kind: "sheet_sample",
            sheetId,
            range: sampleRange,
            maxRows: this.options.maxBlockRows,
            maxCols: this.options.maxBlockCols,
          }),
        );
      }
    }

    // Add sampled blocks for retrieved chunks (query-aware).
    if (retrieved.length && this.options.maxRetrievedBlocks > 0) {
      const retrievedBlocks = await this.blocksFromRetrievedChunks(retrieved, this.options.maxRetrievedBlocks);
      blocks.push(...retrievedBlocks);
    }

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
    for (const t of explicitTableDefs) {
      const name = typeof t?.name === "string" ? String(t.name).trim() : "";
      const sheetId = typeof t?.sheetId === "string" ? String(t.sheetId) : "";
      if (!name || !sheetId) continue;
      let range: string;
      try {
        range = this.schemaRangeRef(sheetId, t.range);
      } catch {
        continue;
      }
      tablesByRange.set(`${sheetId}\u0000${range}`, { sheetId, name, range });
    }
    const tables = Array.from(tablesByRange.values()).sort((a, b) => {
      const ak = `${a.sheetId}\u0000${a.name}\u0000${a.range}`;
      const bk = `${b.sheetId}\u0000${b.name}\u0000${b.range}`;
      return ak.localeCompare(bk);
    });

    const budget = this.computeBudget();

    const payload: WorkbookContextPayload = {
      version: 1,
      workbookId: this.options.workbookId,
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
              retrievedRanges: extractRetrievedRanges(retrieved),
            },
          }
        : {}),
      budget: {
        ...budget,
        usedPromptContextTokens: 0,
      },
    };

    const promptContext = this.buildPromptContext({
      payload,
      ragResult,
      maxTokens: budget.maxPromptContextTokens,
    });

    const usedPromptContextTokens = this.estimator.estimateTextTokens(promptContext);
    payload.budget.usedPromptContextTokens = usedPromptContextTokens;

    return { payload, promptContext, retrieved, indexStats };
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

  private resolveSheetIds(params: { activeSheetId: string }): string[] {
    const ids = Array.isArray(this.options.spreadsheet.listSheets?.()) ? this.options.spreadsheet.listSheets() : [];
    const unique = Array.from(new Set(ids.map((s) => String(s)))).filter(Boolean);
    unique.sort((a, b) => a.localeCompare(b));

    // Always include the active sheet, even if adapters return a stale list.
    if (!unique.includes(params.activeSheetId)) unique.unshift(params.activeSheetId);

    const cap = Math.max(1, this.options.maxSheets);
    const out = unique.slice(0, cap);
    if (!out.includes(params.activeSheetId)) {
      // Ensure active sheet inclusion when it was pushed out by the cap.
      out.pop();
      out.unshift(params.activeSheetId);
    }
    return out;
  }

  private async buildSheetSummary(
    sheetId: string,
    extras?: { namedRanges?: Array<{ name: string; range: string }>; tables?: Array<{ name: string; range: string }> },
  ): Promise<WorkbookContextSheetSummary> {
    const used = this.options.documentController.getUsedRange(sheetId);
    if (!used) {
      return { sheetId, schema: { name: sheetId, tables: [], namedRanges: [], dataRegions: [] } };
    }

    const analyzedRange = clampRange(
      { startRow: used.startRow, endRow: used.endRow, startCol: used.startCol, endCol: used.endCol },
      { maxRows: this.options.maxSchemaRows, maxCols: this.options.maxSchemaCols },
    );

    const block = await this.readBlock({
      kind: "sheet_sample",
      sheetId,
      range: analyzedRange,
      maxRows: this.options.maxSchemaRows,
      maxCols: this.options.maxSchemaCols,
    });

    // If we couldn't read the sheet sample (DLP denied, runtime error, etc), do not attempt
    // schema extraction from placeholder values (it would create misleading fake tables).
    if (block.error) {
      return { sheetId, analyzedRange, schema: { name: sheetId, tables: [], namedRanges: [], dataRegions: [] } };
    }

    const schemaValues: unknown[][] = block.values;
    const schema = extractSheetSchema({
      name: sheetId,
      values: schemaValues,
      // When schema is built from a capped window, preserve the original coordinates.
      origin: { row: analyzedRange.startRow, col: analyzedRange.startCol },
      ...(extras?.namedRanges?.length ? { namedRanges: extras.namedRanges } : {}),
      ...(extras?.tables?.length ? { tables: extras.tables } : {}),
    } as any);

    return { sheetId, analyzedRange, schema };
  }

  private async blocksFromRetrievedChunks(retrieved: unknown[], maxBlocks: number): Promise<WorkbookContextDataBlock[]> {
    const out: WorkbookContextDataBlock[] = [];

    for (const chunk of retrieved) {
      if (out.length >= maxBlocks) break;
      const meta = (chunk as any)?.metadata;
      if (!meta) continue;
      const sheetId = typeof meta.sheetName === "string" ? meta.sheetName : null;
      const rect = meta.rect;
      if (!sheetId || !rect) continue;

      const range = rectToRange(rect);
      if (!range) continue;

      out.push(
        await this.readBlock({
          kind: "retrieved",
          sheetId,
          range,
          maxRows: this.options.maxBlockRows,
          maxCols: this.options.maxBlockCols,
        }),
      );
    }

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

  private buildPromptContext(params: { payload: WorkbookContextPayload; ragResult: any; maxTokens: number }): string {
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

    const sections = [
      {
        key: "workbook_summary",
        priority: priorities.workbook_summary,
        text: `Workbook summary:\n${JSON.stringify(workbookSummary, null, 2)}`,
      },
      {
        key: "sheet_schemas",
        priority: priorities.sheet_schemas,
        text: `Sheet schemas (schema-first):\n${JSON.stringify(sheets, null, 2)}`,
      },
      {
        key: "data_blocks",
        priority: priorities.data_blocks,
        text: blocks.length ? `Sampled data blocks:\n${JSON.stringify(blocks, null, 2)}` : "",
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
            ? `Attachments:\n${JSON.stringify(params.ragResult.attachments, null, 2)}`
            : "",
      },
    ].filter((s) => s.text);

    const packed = packSectionsToTokenBudget(sections as any, Math.max(0, params.maxTokens));
    return packed.map((s) => `## ${s.key}\n${s.text}`).join("\n\n");
  }

  private async readBlock(params: {
    kind: WorkbookContextBlockKind;
    sheetId: string;
    range: Range;
    maxRows: number;
    maxCols: number;
  }): Promise<WorkbookContextDataBlock> {
    const clamped = clampRange(params.range, { maxRows: params.maxRows, maxCols: params.maxCols });
    const rangeRef = this.rangeRef(params.sheetId, clamped);

    const executor = new ToolExecutor(this.options.spreadsheet, {
      default_sheet: params.sheetId,
      dlp: this.options.dlp,
      // Keep samples small even if callers accidentally request huge ranges.
      max_read_range_cells: Math.max(1, params.maxRows * params.maxCols),
      // Allow somewhat larger payloads for schema extraction, but still guard against
      // pathological "giant text pasted into a single cell" scenarios.
      max_read_range_chars: 200_000,
    });

    const toolResult = await executor.execute({
      name: "read_range",
      parameters: { range: rangeRef, include_formulas: false },
    } as any);

    if (toolResult.ok && toolResult.data && (toolResult.data as any).values) {
      const values = ((toolResult.data as any).values ?? []) as CellScalar[][];
      return {
        kind: params.kind,
        sheetId: params.sheetId,
        range: rangeRef,
        rowHeaders: rowHeaders(clamped),
        colHeaders: colHeaders(clamped),
        values,
      };
    }

    // Never leak restricted content: fall back to a deterministic placeholder matrix.
    const error = toolResult.error ?? { code: "runtime_error", message: "Unknown error reading range." };
    const placeholder = String(error.code) === "permission_denied" ? "[POLICY_DENIED]" : "[UNAVAILABLE]";
    return {
      kind: params.kind,
      sheetId: params.sheetId,
      range: rangeRef,
      rowHeaders: rowHeaders(clamped),
      colHeaders: colHeaders(clamped),
      values: [[placeholder]],
      error: { code: String(error.code ?? "runtime_error"), message: String(error.message ?? "Unknown error") },
    };
  }

  private rangeRef(sheetId: string, range: Range): string {
    return `${formatSheetNameForA1(sheetId)}!${rangeToA1Selection(range)}`;
  }

  private schemaRangeRef(sheetId: string, range: Range): string {
    // `packages/ai-context`'s A1 parser does not implement Excel-style quoted sheet names.
    // Keep sheet names unquoted here so schema extraction can safely parse its own ranges.
    return `${sheetId}!${rangeToA1Selection(range)}`;
  }
}

function safeList<T>(fn: () => T[] | null | undefined): T[] {
  try {
    const out = fn();
    return Array.isArray(out) ? out : [];
  } catch {
    return [];
  }
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

function formatSheetNameForA1(sheetName: string): string {
  if (/^[A-Za-z0-9_]+$/.test(sheetName)) return sheetName;
  return `'${sheetName.replace(/'/g, "''")}'`;
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

function extractRetrievedRanges(retrieved: unknown[]): string[] {
  const out: string[] = [];
  for (const chunk of retrieved) {
    const meta = (chunk as any)?.metadata;
    if (!meta) continue;
    const sheetName = typeof meta.sheetName === "string" ? meta.sheetName : null;
    const rect = meta.rect;
    if (!sheetName || !rect) continue;
    try {
      const range = rectToA1(rect);
      out.push(`${formatSheetNameForA1(sheetName)}!${range}`);
    } catch {
      // ignore malformed rect metadata
    }
  }
  return out.sort();
}
