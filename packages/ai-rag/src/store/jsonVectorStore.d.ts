import type { InMemoryVectorStore, VectorRecord, VectorSearchResult } from "./inMemoryVectorStore.js";

export class JsonVectorStore extends InMemoryVectorStore {
  constructor(opts: { storage?: any; dimension: number; autoSave?: boolean });
  load(): Promise<void>;

  upsert(records: VectorRecord[]): Promise<void>;
  delete(ids: string[]): Promise<void>;
  get(id: string): Promise<{ id: string; vector: Float32Array; metadata: any } | null>;
  list(opts?: {
    filter?: (metadata: any, id: string) => boolean;
    workbookId?: string;
    includeVector?: boolean;
    signal?: AbortSignal;
  }): Promise<Array<{ id: string; vector?: Float32Array; metadata: any }>>;
  query(
    vector: ArrayLike<number>,
    topK: number,
    opts?: { filter?: (metadata: any, id: string) => boolean; workbookId?: string; signal?: AbortSignal }
  ): Promise<VectorSearchResult[]>;
  close(): Promise<void>;
}
