export type Attachment = { type: "range" | "formula" | "table" | "chart"; reference: string; data?: any };

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

/**
 * DLP options accepted by ContextManager methods.
 *
 * Both camelCase and snake_case field names are supported so callers can pass options
 * deserialized from JSON or from non-TS hosts.
 */
export type DlpOptions = {
  // Required identifiers (at least one form should be provided).
  documentId?: string;
  document_id?: string;

  // Single-sheet contexts may provide a stable sheet id.
  sheetId?: string;
  sheet_id?: string;

  policy?: any;

  classificationRecords?: Array<{ selector: any; classification: any }>;
  classification_records?: Array<{ selector: any; classification: any }>;
  classificationStore?: { list(documentId: string): Array<{ selector: any; classification: any }> };
  classification_store?: { list(documentId: string): Array<{ selector: any; classification: any }> };

  includeRestrictedContent?: boolean;
  include_restricted_content?: boolean;

  auditLogger?: { log(event: any): void };

  sheetNameResolver?: any;
  sheet_name_resolver?: any;
};

import type { TokenEstimator } from "./tokenBudget.js";

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
  }): Promise<{ schema: any; retrieved: any[]; sampledRows: any[]; promptContext: string }>;

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
  }): Promise<{ indexStats: any; retrieved: any[]; promptContext: string }>;

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
  }): Promise<{ indexStats: any; retrieved: any[]; promptContext: string }>;
}
