export class SqliteVectorStore {
  static create(opts: {
    storage?: any;
    dimension: number;
    autoSave?: boolean;
    locateFile?: (file: string, prefix?: string) => string;
  }): Promise<SqliteVectorStore>;

  readonly dimension: number;

  upsert(records: Array<{ id: string; vector: ArrayLike<number>; metadata: any }>): Promise<void>;
  delete(ids: string[]): Promise<void>;
  compact(): Promise<void>;
  vacuum(): Promise<void>;
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
  ): Promise<Array<{ id: string; score: number; metadata: any }>>;
  close(): Promise<void>;
}
