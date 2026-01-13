import type { BinaryStorage } from "./binaryStorage.js";

export class SqliteVectorStore {
  static create(opts: {
    storage?: BinaryStorage;
    dimension: number;
    autoSave?: boolean;
    /**
     * When true (default), failures to load/initialize an existing persisted DB
     * (e.g. corrupted bytes) will cause the store to clear the persisted payload
     * (when possible) and create a fresh empty DB so callers can re-index.
     *
     * Set to false to preserve the historical behaviour (throw on corruption).
     */
    resetOnCorrupt?: boolean;
    /**
     * When true (default), if a persisted DB exists with a different embedding
     * dimension than requested, the store will wipe the persisted bytes and
     * create a fresh empty DB so callers can re-index.
     *
     * Set to false to preserve the historical behaviour (throw on mismatch).
     */
    resetOnDimensionMismatch?: boolean;
    locateFile?: (file: string, prefix?: string) => string;
  }): Promise<SqliteVectorStore>;

  readonly dimension: number;

  batch<T>(fn: () => Promise<T> | T): Promise<T>;

  upsert(records: Array<{ id: string; vector: ArrayLike<number>; metadata: any }>): Promise<void>;
  updateMetadata(records: Array<{ id: string; metadata: any }>): Promise<void>;
  delete(ids: string[]): Promise<void>;
  deleteWorkbook(workbookId: string): Promise<number>;
  clear(): Promise<void>;
  /**
   * Run SQLite `VACUUM` to reclaim space after large deletions and persist the
   * compacted database snapshot.
   *
   * This persists regardless of the `autoSave` setting (compaction is an explicit,
   * manual operation intended to reclaim storage).
   */
  compact(): Promise<void>;
  /**
   * Alias for {@link compact}.
   */
  vacuum(): Promise<void>;
  get(id: string): Promise<{ id: string; vector: Float32Array; metadata: any } | null>;
  listContentHashes(opts?: {
    workbookId?: string;
    signal?: AbortSignal;
  }): Promise<Array<{ id: string; contentHash: string | null; metadataHash: string | null }>>;
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

export type SqliteVectorStoreDimensionMismatchError = Error & {
  name: "SqliteVectorStoreDimensionMismatchError";
  dbDimension: number;
  requestedDimension: number;
};
