import type { VectorSearchResult } from "../store/inMemoryVectorStore.js";

export function searchWorkbookRag(params: {
  queryText: string;
  workbookId?: string;
  topK?: number;
  vectorStore: {
    query(
      vector: ArrayLike<number>,
      topK: number,
      opts?: { workbookId?: string; signal?: AbortSignal }
    ): Promise<VectorSearchResult[]>;
  };
  embedder: {
    embedTexts(texts: string[], options?: { signal?: AbortSignal }): Promise<ArrayLike<number>[]>;
  };
  rerank?: boolean;
  dedupe?: boolean;
  signal?: AbortSignal;
}): Promise<VectorSearchResult[]>;
