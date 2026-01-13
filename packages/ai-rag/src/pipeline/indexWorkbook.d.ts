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
  embedder: {
    /**
     * Optional identity string (used for cache keys / persisted metadata).
     * When absent, `indexWorkbook` falls back to `"unknown-embedder"`.
     */
    name?: string;
    embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<ArrayLike<number>[]>;
  };
  sampleRows?: number;
  maxColumnsForSchema?: number;
  maxColumnsForRows?: number;
  /**
   * Custom token estimator used to populate `metadata.tokenCount` for each chunk.
   *
   * Defaults to {@link approximateTokenCount}.
   */
  tokenCount?: (text: string) => number;
  embedBatchSize?: number;
  onProgress?: (info: { phase: "chunk" | "hash" | "embed" | "upsert" | "delete"; processed: number; total?: number }) => void;
  transform?: (
    record: { id: string; text: string; metadata: any }
  ) =>
    | { text?: string | null; metadata?: any }
    | null
    | Promise<{ text?: string | null; metadata?: any } | null>;
  signal?: AbortSignal;
}): Promise<{ totalChunks: number; upserted: number; skipped: number; deleted: number }>;
