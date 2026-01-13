/**
 * Deterministic, offline hash-based embeddings used by Formula for workbook RAG.
 *
 * This avoids requiring user API keys or local model setup. Retrieval quality is
 * lower than true ML embeddings, but works for basic semantic-ish retrieval.
 *
 * Note: Embeddings are not user-configurable in Formula; a future Cursor-managed
 * embedding service can replace this.
 */
export class HashEmbedder {
  constructor(opts?: { dimension?: number; cacheSize?: number });
  readonly dimension: number;
  readonly name: string;
  embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<Float32Array[]>;
}
