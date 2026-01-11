export function approximateTokenCount(text: string): number;

export function indexWorkbook(params: {
  workbook: any;
  vectorStore: any;
  embedder: { embedTexts(texts: string[]): Promise<ArrayLike<number>[]> };
  sampleRows?: number;
  transform?: (
    record: { id: string; text: string; metadata: any }
  ) => { text?: string; metadata?: any } | null | Promise<{ text?: string; metadata?: any } | null>;
}): Promise<{ totalChunks: number; upserted: number; skipped: number; deleted: number }>;

