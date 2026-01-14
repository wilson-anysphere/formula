import type { SheetSchema } from "./schema.js";

export interface RagSheet {
  name: string;
  values: unknown[][];
  /**
   * Optional coordinate origin (0-based) for the provided `values` matrix.
   *
   * When `values` is a cropped window of a larger sheet (e.g. a capped used-range
   * sample), `origin` lets schema extraction and RAG chunking produce correct
   * absolute A1 ranges while slicing from the local matrix.
   */
  origin?: { row: number; col: number };
}

export type RagChunkMetadata =
  | { type: "region"; sheetName: string; regionRange: string }
  | { type: "range"; sheetName: string };

export interface RagChunk<TMetadata = RagChunkMetadata> {
  id: string;
  /**
   * Sheet-qualified A1 range string (e.g. `"Sheet1!A1:B3"`).
   */
  range: string;
  /**
   * Prompt-friendly text representation (TSV-ish preview) of the chunk.
   */
  text: string;
  metadata: TMetadata;
}

export interface ChunkSheetByRegionsOptions {
  maxChunkRows?: number;
  /**
   * Alias for `splitByRowWindows`.
   */
  splitRegions?: boolean;
  /**
   * Split tall regions into multiple row-window chunks to improve retrieval quality.
   * Defaults to `false` for backwards compatibility.
   */
  splitByRowWindows?: boolean;
  /**
   * Alias for `rowOverlap`.
   */
  chunkRowOverlap?: number;
  /** Row overlap between windows (only when splitting). Defaults to 3. */
  rowOverlap?: number;
  /** Maximum number of chunks per region (only when splitting). Defaults to 50. */
  maxChunksPerRegion?: number;
  signal?: AbortSignal;
}

export interface Embedder {
  embed(text: string, options?: { signal?: AbortSignal }): Promise<number[]>;
}

export interface VectorStoreItem<TMetadata = unknown> {
  id: string;
  embedding: number[];
  metadata: TMetadata;
  text: string;
}

export interface VectorStoreSearchResult<TItem = VectorStoreItem> {
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
  /**
   * Optional method for per-sheet re-indexing when the number of chunks can shrink.
   */
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

export class InMemoryVectorStore<TMetadata = unknown> implements VectorStore<VectorStoreItem<TMetadata>> {
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

/**
 * Chunk a sheet by detected regions for a simple RAG pipeline.
 */
export function chunkSheetByRegions(
  sheet: RagSheet,
  options?: ChunkSheetByRegionsOptions,
): Array<RagChunk<{ type: "region"; sheetName: string; regionRange: string }>>;

/**
 * Like `chunkSheetByRegions`, but also returns the extracted schema so callers can
 * reuse it without a second pass.
 */
export function chunkSheetByRegionsWithSchema(
  sheet: RagSheet,
  options?: ChunkSheetByRegionsOptions,
): { schema: SheetSchema; chunks: Array<RagChunk<{ type: "region"; sheetName: string; regionRange: string }>> };

export class RagIndex {
  embedder: Embedder;
  store: VectorStore;
  constructor(options?: { embedder?: Embedder; store?: VectorStore });
  /**
   * Index a sheet, replacing any previous region chunks for the same sheet.
   */
  indexSheet(
    sheet: RagSheet,
    options?: ChunkSheetByRegionsOptions,
  ): Promise<{ schema: SheetSchema; chunkCount: number }>;
  search(
    query: string,
    topK?: number,
    options?: { signal?: AbortSignal },
  ): Promise<Array<{ range: string; score: number; preview: string }>>;
}

/**
 * Convenience for building a single chunk from a numeric range in a sheet.
 */
export function rangeToChunk(
  sheet: RagSheet,
  range: { startRow: number; startCol: number; endRow: number; endCol: number },
  options?: { maxRows?: number },
): RagChunk<{ type: "range"; sheetName: string }>;
