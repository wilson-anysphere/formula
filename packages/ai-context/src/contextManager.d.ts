import type { DlpFinding, DlpLevel } from "./dlp.js";
import type { NamedRangeSchema } from "./schema.js";
import type { RagIndex } from "./rag.js";
import type { SheetSchema } from "./schema.js";
import type { TokenEstimator } from "./tokenBudget.js";

export type Attachment = { type: "range" | "formula" | "table" | "chart"; reference: string; data?: unknown };

export interface ContextSheet {
  name: string;
  values: unknown[][];
  /**
   * Optional coordinate origin (0-based) for the provided `values` matrix.
   *
   * When `values` is a cropped window of a larger sheet (e.g. a capped used-range
   * sample), `origin` lets schema extraction and context formatting produce
   * correct absolute A1 ranges.
   */
  origin?: { row: number; col: number };
  namedRanges?: NamedRangeSchema[];
  /**
   * Optional explicit table definitions (used by schema extraction).
   */
  tables?: Array<{ name: string; range: string; [key: string]: unknown }>;
  /**
   * Allow host-specific sheet fields without tripping TS excess-property checks
   * (e.g. internal ids, metadata, etc).
   */
  [key: string]: unknown;
}

export interface WorkbookRagVectorStore {
  /**
   * Optional embedding dimension. When provided, indexing will validate that all
   * embeddings match this size.
   */
  readonly dimension?: number;

  /**
   * Query for the nearest neighbors to a vector. Must return `{ id, score, metadata }`
   * objects (metadata is prompt-unsafe and will be redacted before ContextManager
   * returns it).
   */
  query(
    vector: ArrayLike<number>,
    topK: number,
    opts?: { filter?: (metadata: unknown, id: string) => boolean; workbookId?: string; signal?: AbortSignal },
  ): Promise<Array<{ id: string; score: number; metadata: unknown }>>;

  /**
   * Optional fast-path listing used by `packages/ai-rag` for incremental indexing.
   */
  listContentHashes?: (opts?: {
    workbookId?: string;
    signal?: AbortSignal;
  }) => Promise<Array<{ id: string; contentHash: string | null; metadataHash: string | null }>>;

  // Indexing/persistence methods (required when `skipIndexing` is false).
  list?: (opts?: {
    filter?: (metadata: unknown, id: string) => boolean;
    workbookId?: string;
    includeVector?: boolean;
    signal?: AbortSignal;
  }) => Promise<Array<{ id: string; vector?: Float32Array; metadata: unknown }>>;
  upsert?: (records: Array<{ id: string; vector: ArrayLike<number>; metadata: unknown }>) => Promise<void>;
  updateMetadata?: (records: Array<{ id: string; metadata: unknown }>) => Promise<void>;
  delete?: (ids: string[]) => Promise<void>;
  batch?: <T>(fn: () => Promise<T> | T) => Promise<T>;
  close?: () => Promise<void>;
}

export interface WorkbookRagRect {
  r0: number;
  c0: number;
  r1: number;
  c1: number;
}

export interface WorkbookRagTable {
  name: string;
  sheetName: string;
  rect: WorkbookRagRect;
  [key: string]: unknown;
}

export interface WorkbookRagNamedRange {
  name: string;
  sheetName: string;
  rect: WorkbookRagRect;
  [key: string]: unknown;
}

export interface WorkbookRagSheet {
  name: string;
  /**
   * Either a 2D matrix `[row][col]`, a sparse Map keyed by coordinates, or any other
   * host-specific cell representation understood by `packages/ai-rag`.
   */
  cells?: unknown;
  /**
   * Optional alternative to `cells` (treated as `[row][col]`).
   */
  values?: unknown[][];
  /**
   * Optional random-access cell reader. When provided, workbook schema extraction can
   * avoid materializing a dense matrix for very large sheets.
   */
  getCell?: (row: number, col: number) => unknown;
  [key: string]: unknown;
}

export interface WorkbookRagWorkbook {
  id: string;
  sheets: WorkbookRagSheet[];
  tables?: WorkbookRagTable[];
  namedRanges?: WorkbookRagNamedRange[];
  [key: string]: unknown;
}

export interface SpreadsheetApiLike {
  listSheets(): string[];
  listNonEmptyCells?: (
    sheet?: string,
  ) => Array<{ address: { sheet: string; row: number; col: number }; cell: { value?: unknown; formula?: string } }>;
  /**
   * Optional sheet name resolver available on some SpreadsheetApi hosts (desktop).
   * ContextManager forwards this through to DLP enforcement when callers do not
   * provide an explicit resolver.
   */
  sheetNameResolver?: SheetNameResolverLike | null;
  sheet_name_resolver?: SheetNameResolverLike | null;
  [key: string]: unknown;
}

export interface SpreadsheetApiWithNonEmptyCells extends SpreadsheetApiLike {
  listNonEmptyCells: (
    sheet?: string,
  ) => Array<{ address: { sheet: string; row: number; col: number }; cell: { value?: unknown; formula?: string } }>;
}

export interface SheetNameResolverLike {
  getSheetIdByName(name: string): string | null | undefined;
  getSheetNameById?: (id: string) => string | null | undefined;
  [key: string]: unknown;
}

export interface RetrievedSheetChunk {
  range: string;
  score: number;
  preview: string;
}

export interface WorkbookChunkDlpInfo {
  level: DlpLevel;
  findings: DlpFinding[];
}

export interface WorkbookChunkMetadata {
  workbookId?: string;
  sheetName?: string;
  kind?: string;
  title?: string;
  rect?: WorkbookRagRect;
  [key: string]: unknown;
}

export interface RetrievedWorkbookChunk {
  id: string;
  score: number;
  metadata: WorkbookChunkMetadata;
  text: string;
  dlp: WorkbookChunkDlpInfo;
}

export interface BuildContextResult {
  schema: SheetSchema;
  retrieved: RetrievedSheetChunk[];
  sampledRows: unknown[][];
  promptContext: string;
}

export interface BuildWorkbookContextResult {
  indexStats: WorkbookIndexStats | null;
  retrieved: RetrievedWorkbookChunk[];
  promptContext: string;
}

export interface WorkbookIndexStats {
  totalChunks: number;
  upserted: number;
  skipped: number;
  deleted: number;
}

export interface DlpClassificationRecord {
  selector: unknown;
  classification: unknown;
  [key: string]: unknown;
}

/**
 * DLP configuration and inputs for ContextManager.
 *
 * Note: ContextManager accepts both camelCase and snake_case field names for
 * compatibility with a variety of hosts.
 */
export interface DlpOptions {
  documentId?: string;
  document_id?: string;
  sheetId?: string;
  sheet_id?: string;
  /**
   * Policy object passed through to `packages/security/dlp`.
   *
   * Kept intentionally generic to avoid cross-package type coupling.
   */
  policy: unknown;
  classificationRecords?: DlpClassificationRecord[];
  classification_records?: DlpClassificationRecord[];
  classificationStore?: { list(documentId: string): DlpClassificationRecord[]; [key: string]: unknown } | null;
  classification_store?: { list(documentId: string): DlpClassificationRecord[]; [key: string]: unknown } | null;
  includeRestrictedContent?: boolean;
  include_restricted_content?: boolean;
  /**
   * Optional audit sink. Return type is intentionally flexible to allow both sync and async loggers.
   */
  auditLogger?: { log(event: unknown): unknown; [key: string]: unknown } | null;
  /**
   * Optional sheet name <-> id resolver used for structured DLP enforcement.
   */
  sheetNameResolver?: SheetNameResolverLike | null;
  sheet_name_resolver?: SheetNameResolverLike | null;
}

/**
 * Input type for `ContextManager` DLP settings.
 *
 * ContextManager requires a document id when DLP is enabled. We support both camelCase
 * (`documentId`) and snake_case (`document_id`) to match a variety of hosts.
 */
export type DlpOptionsInput = DlpOptions & ({ documentId: string } | { document_id: string });

export type WorkbookRagOptions = {
  vectorStore: WorkbookRagVectorStore;
  /**
   * Workbook RAG embedder.
   *
   * Note: In Formula, embeddings are not user-configurable; the desktop app uses
   * deterministic hash embeddings by default. A future Cursor-managed embedding
   * service can replace this.
   */
  embedder: { embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<ArrayLike<number>[]> };
  topK?: number;
  sampleRows?: number;
};
export class ContextManager {
  tokenBudgetTokens: number;
  ragIndex: RagIndex;
  workbookRag?: WorkbookRagOptions;

  constructor(options?: {
    tokenBudgetTokens?: number;
    ragIndex?: RagIndex;
    /**
     * Cache single-sheet RAG indexing by content signature.
     *
     * When enabled (default), repeated `buildContext()` calls for an unchanged sheet
     * will reuse the previously indexed chunks instead of re-embedding.
     */
    cacheSheetIndex?: boolean;
    /**
     * Maximum number of sheet index entries cached per ContextManager instance.
     *
     * Defaults to 32. When the cache evicts an active sheet entry, its in-memory
     * RAG chunks are also removed from the underlying store to keep memory bounded.
     */
    sheetIndexCacheLimit?: number;
    workbookRag?: WorkbookRagOptions;
    /**
     * Safety cap for the number of rows included from `sheet.values` when building
     * single-sheet context. Defaults to 1000.
     */
    maxContextRows?: number;
    /**
     * Safety cap for the number of columns included from `sheet.values` when building
     * single-sheet context. Defaults to 500.
     *
     * This protects against pathological "short but extremely wide" selections (e.g.
     * full-row attachments) that would otherwise produce enormous schema/sample/chunk
     * strings even when the row count is small.
     */
    maxContextCols?: number;
    /**
     * Safety cap for the total number of cells included from `sheet.values` when building
     * single-sheet context. Defaults to 200_000.
     */
    maxContextCells?: number;
    /**
     * Max rows included in each sheet-level RAG chunk preview (TSV lines). Defaults to 30.
     */
    maxChunkRows?: number;
    /**
     * Split tall sheet regions into multiple row windows for better retrieval quality.
     *
     * Defaults to `false` (opt-in).
     */
    splitRegions?: boolean;
    /**
     * Row overlap between region windows (only when splitting). Defaults to 3.
     */
    chunkRowOverlap?: number;
    /**
     * Maximum number of chunks per region (only when splitting). Defaults to 50.
     */
    maxChunksPerRegion?: number;
    /**
     * Top-K retrieved regions for sheet-level (non-workbook) RAG. Defaults to 5.
     */
    sheetRagTopK?: number;
    redactor?: (text: string) => string;
    tokenEstimator?: TokenEstimator;
  });

  buildContext(params: {
    sheet: ContextSheet;
    query: string;
    attachments?: Attachment[];
    sampleRows?: number;
    samplingStrategy?: "random" | "stratified" | "head" | "tail" | "systematic";
    stratifyByColumn?: number;
    limits?: {
      maxContextRows?: number;
      maxContextCols?: number;
      maxContextCells?: number;
      maxChunkRows?: number;
      /**
       * Split tall sheet regions into multiple row windows for better retrieval quality.
       *
       * Defaults to `false` for backwards compatibility.
       */
      splitRegions?: boolean;
      /**
       * Row overlap between region windows (only when splitting).
       */
      chunkRowOverlap?: number;
      /**
       * Maximum number of chunks per region (only when splitting).
       */
      maxChunksPerRegion?: number;
    };
    signal?: AbortSignal;
    dlp?: DlpOptionsInput;
  }): Promise<BuildContextResult>;

  clearSheetIndexCache(options?: { clearStore?: boolean; signal?: AbortSignal }): Promise<void>;

  buildWorkbookContext(params: {
    workbook: WorkbookRagWorkbook;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    skipIndexing?: boolean;
    skipIndexingWithDlp?: boolean;
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp?: DlpOptionsInput;
  }): Promise<BuildWorkbookContextResult>;

  buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: SpreadsheetApiLike;
    workbookId: string;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    /**
     * Cheap path: caller manages indexing externally and does not provide DLP.
     * `spreadsheet.listNonEmptyCells` is not required in this mode.
     */
    skipIndexing: true;
    skipIndexingWithDlp?: boolean;
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp?: undefined;
  }): Promise<BuildWorkbookContextResult>;

  buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: SpreadsheetApiLike;
    workbookId: string;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    /**
     * Cheap path with DLP: caller provides an up-to-date DLP-safe index (already
     * redacted before embedding), so ContextManager can skip workbook scans.
     * `spreadsheet.listNonEmptyCells` is not required in this mode.
     */
    skipIndexing: true;
    skipIndexingWithDlp: true;
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp: DlpOptionsInput;
  }): Promise<BuildWorkbookContextResult>;

  buildWorkbookContextFromSpreadsheetApi(params: {
    /**
     * When ContextManager may need to scan workbook cells for indexing (default),
     * callers must provide `listNonEmptyCells`.
     */
    spreadsheet: SpreadsheetApiWithNonEmptyCells;
    workbookId: string;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    skipIndexing?: boolean;
    skipIndexingWithDlp?: boolean;
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp?: DlpOptionsInput;
  }): Promise<BuildWorkbookContextResult>;
}
