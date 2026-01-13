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

import type { TokenEstimator } from "./tokenBudget.js";

export class ContextManager {
  constructor(options?: {
    tokenBudgetTokens?: number;
    ragIndex?: any;
    workbookRag?: WorkbookRagOptions;
    redactor?: (text: string) => string;
    tokenEstimator?: TokenEstimator;
  });

  buildContext(params: {
    sheet: { name: string; values: unknown[][]; namedRanges?: any[] };
    query: string;
    attachments?: Attachment[];
    sampleRows?: number;
    samplingStrategy?: "random" | "stratified" | "head" | "systematic";
    stratifyByColumn?: number;
    signal?: AbortSignal;
    dlp?: any;
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
    dlp?: any;
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
    dlp?: any;
  }): Promise<{ indexStats: any; retrieved: any[]; promptContext: string }>;
}
