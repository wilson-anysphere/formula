export type Attachment = { type: "range" | "formula" | "table" | "chart"; reference: string; data?: any };

export type WorkbookRagOptions = {
  vectorStore: any;
  embedder: { embedTexts(texts: string[]): Promise<ArrayLike<number>[]> };
  topK?: number;
  sampleRows?: number;
};

export class ContextManager {
  constructor(options?: {
    tokenBudgetTokens?: number;
    ragIndex?: any;
    workbookRag?: WorkbookRagOptions;
    redactor?: (text: string) => string;
  });

  buildContext(params: {
    sheet: { name: string; values: unknown[][]; namedRanges?: any[] };
    query: string;
    attachments?: Attachment[];
    sampleRows?: number;
    samplingStrategy?: "random" | "stratified";
    stratifyByColumn?: number;
    dlp?: any;
  }): Promise<{ schema: any; retrieved: any[]; sampledRows: any[]; promptContext: string }>;

  buildWorkbookContext(params: {
    workbook: any;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    skipIndexing?: boolean;
    dlp?: any;
  }): Promise<{ indexStats: any; retrieved: any[]; promptContext: string }>;

  buildWorkbookContextFromSpreadsheetApi(params: {
    spreadsheet: any;
    workbookId: string;
    query: string;
    attachments?: Attachment[];
    topK?: number;
    skipIndexing?: boolean;
    dlp?: any;
  }): Promise<{ indexStats: any; retrieved: any[]; promptContext: string }>;
}

