import type { SheetSchema } from "./schema.js";
import type { TokenEstimator } from "./tokenBudget.js";

export type Attachment = { type: "range" | "formula" | "table" | "chart"; reference: string; data?: unknown };

export interface RetrievedSheetChunk {
  range: string;
  score: number;
  preview: string;
}

export interface WorkbookChunkDlpInfo {
  level: string;
  findings: string[];
}

export interface RetrievedWorkbookChunk {
  id: string;
  score: number;
  metadata: Record<string, unknown>;
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
  indexStats: unknown | null;
  retrieved: RetrievedWorkbookChunk[];
  promptContext: string;
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
  policy?: unknown;
  classificationRecords?: Array<{ selector: unknown; classification: unknown }>;
  classification_records?: Array<{ selector: unknown; classification: unknown }>;
  classificationStore?: { list(documentId: string): Array<{ selector: unknown; classification: unknown }> };
  classification_store?: { list(documentId: string): Array<{ selector: unknown; classification: unknown }> };
  includeRestrictedContent?: boolean;
  include_restricted_content?: boolean;
  auditLogger?: { log(event: any): void };
  /**
   * Optional sheet name <-> id resolver used for structured DLP enforcement.
   */
  sheetNameResolver?: { getSheetIdByName(name: string): string | null | undefined };
  sheet_name_resolver?: { getSheetIdByName(name: string): string | null | undefined };
}

export type WorkbookRagOptions = {
  vectorStore: any;
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
  constructor(options?: {
    tokenBudgetTokens?: number;
    ragIndex?: any;
    /**
     * Cache single-sheet RAG indexing by content signature.
     *
     * When enabled (default), repeated `buildContext()` calls for an unchanged sheet
     * will reuse the previously indexed chunks instead of re-embedding.
     */
    cacheSheetIndex?: boolean;
    workbookRag?: WorkbookRagOptions;
    /**
     * Safety cap for the number of rows included from `sheet.values` when building
     * single-sheet context. Defaults to 1000.
     */
    maxContextRows?: number;
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
     * Top-K retrieved regions for sheet-level (non-workbook) RAG. Defaults to 5.
     */
    sheetRagTopK?: number;
    redactor?: (text: string) => string;
    tokenEstimator?: TokenEstimator;
  });

  buildContext(params: {
    sheet: { name: string; values: unknown[][]; origin?: { row: number; col: number }; namedRanges?: any[] };
    query: string;
    attachments?: Attachment[];
    sampleRows?: number;
    samplingStrategy?: "random" | "stratified" | "head" | "tail" | "systematic";
    stratifyByColumn?: number;
    limits?: { maxContextRows?: number; maxContextCells?: number; maxChunkRows?: number };
    signal?: AbortSignal;
    dlp?: DlpOptions;
  }): Promise<BuildContextResult>;

  buildWorkbookContext(params: {
    workbook: any;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    skipIndexing?: boolean;
    skipIndexingWithDlp?: boolean;
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp?: DlpOptions;
  }): Promise<BuildWorkbookContextResult>;

  buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: any;
    workbookId: string;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    skipIndexing?: boolean;
    skipIndexingWithDlp?: boolean;
    includePromptContext?: boolean;
    signal?: AbortSignal;
    dlp?: DlpOptions;
  }): Promise<BuildWorkbookContextResult>;
}
