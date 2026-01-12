export function approximateTokenCount(text: string): number;

export function indexWorkbook(params: {
  workbook: any;
  vectorStore: any;
  /**
   * Embedder used to convert chunk text into vectors.
   *
   * Note: Formula's desktop workbook RAG uses deterministic, offline hash
   * embeddings (`HashEmbedder`) by default. Embeddings are not user-configurable
   * (no API keys / no local model setup). A future Cursor-managed
   * embedding service can replace this to improve retrieval quality.
   */
  embedder: { embedTexts(texts: string[]): Promise<ArrayLike<number>[]> };
  sampleRows?: number;
  transform?: (
    record: { id: string; text: string; metadata: any }
  ) => { text?: string; metadata?: any } | null | Promise<{ text?: string; metadata?: any } | null>;
  signal?: AbortSignal;
}): Promise<{ totalChunks: number; upserted: number; skipped: number; deleted: number }>;
