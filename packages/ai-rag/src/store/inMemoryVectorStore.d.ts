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
  get(id: string): Promise<{ id: string; vector: Float32Array; metadata: any } | null>;
  list(opts?: {
    filter?: (metadata: any, id: string) => boolean;
    workbookId?: string;
    includeVector?: boolean;
  }): Promise<Array<{ id: string; vector?: Float32Array; metadata: any }>>;
  query(
    vector: ArrayLike<number>,
    topK: number,
    opts?: { filter?: (metadata: any, id: string) => boolean; workbookId?: string }
  ): Promise<VectorSearchResult[]>;
  close(): Promise<void>;
}

