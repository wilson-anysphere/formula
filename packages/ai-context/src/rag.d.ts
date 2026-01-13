export interface RagSheet {
  name: string;
  values: unknown[][];
}

export interface RagChunk<TMetadata = any> {
  id: string;
  range: string;
  text: string;
  metadata: TMetadata;
}

export interface Embedder {
  embed(text: string, options?: { signal?: AbortSignal }): Promise<number[]>;
}

export interface VectorStoreItem<TMetadata = any> {
  id: string;
  embedding: number[];
  metadata: TMetadata;
  text: string;
}

export interface VectorStoreSearchResult<TItem = any> {
  item: TItem;
  score: number;
}

export interface VectorStore<TItem = VectorStoreItem> {
  add(items: TItem[], options?: { signal?: AbortSignal }): Promise<void>;
  search(
    queryEmbedding: number[],
    topK: number,
    options?: { signal?: AbortSignal },
  ): Promise<Array<VectorStoreSearchResult<TItem>>>;
  deleteByPrefix?: (prefix: string, options?: { signal?: AbortSignal }) => Promise<void>;
  readonly size?: number;
}

export class HashEmbedder implements Embedder {
  dimension: number;

  constructor(options?: { dimension?: number });

  /**
   * Deterministic, offline hash-based embeddings.
   */
  embed(text: string, options?: { signal?: AbortSignal }): Promise<number[]>;
}

export class InMemoryVectorStore<TMetadata = any>
  implements VectorStore<VectorStoreItem<TMetadata>>
{
  items: Map<string, VectorStoreItem<TMetadata>>;

  constructor();

  add(items: Array<VectorStoreItem<TMetadata>>, options?: { signal?: AbortSignal }): Promise<void>;

  search(
    queryEmbedding: number[],
    topK: number,
    options?: { signal?: AbortSignal },
  ): Promise<Array<VectorStoreSearchResult<VectorStoreItem<TMetadata>>>>;

  deleteByPrefix(prefix: string, options?: { signal?: AbortSignal }): Promise<void>;

  get size(): number;
}

export function chunkSheetByRegions(
  sheet: RagSheet,
  options?: { maxChunkRows?: number; signal?: AbortSignal },
): Array<RagChunk<{ type: "region"; sheetName: string }>>;

export class RagIndex {
  embedder: Embedder;
  store: VectorStore;

  constructor(options?: { embedder?: Embedder; store?: VectorStore });

  indexSheet(sheet: RagSheet, options?: { signal?: AbortSignal }): Promise<void>;

  search(
    query: string,
    topK?: number,
    options?: { signal?: AbortSignal },
  ): Promise<Array<{ range: string; score: number; preview: string }>>;
}

export function rangeToChunk(
  sheet: RagSheet,
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
  options?: { maxRows?: number },
): RagChunk<{ type: "range"; sheetName: string }>;
