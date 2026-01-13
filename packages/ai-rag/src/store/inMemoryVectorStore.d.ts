export type VectorRecord = {
  id: string;
  vector: ArrayLike<number>;
  metadata: any;
};

export type VectorSearchResult = {
  id: string;
  score: number;
  metadata: any;
};

export class InMemoryVectorStore {
  constructor(opts: { dimension: number });
  readonly dimension: number;

  upsert(records: VectorRecord[]): Promise<void>;
  delete(ids: string[]): Promise<void>;
  deleteWorkbook(workbookId: string): Promise<number>;
  clear(): Promise<void>;
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
