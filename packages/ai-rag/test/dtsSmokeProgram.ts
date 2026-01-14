import {
  ChunkedLocalStorageBinaryStorage,
  type BinaryStorage,
  HashEmbedder,
  IndexedDBBinaryStorage,
  InMemoryBinaryStorage,
  InMemoryVectorStore,
  JsonVectorStore,
  LocalStorageBinaryStorage,
  SqliteVectorStore,
  type SqliteVectorStoreDimensionMismatchError,
  type SqliteVectorStoreInvalidMetadataError,
  approximateTokenCount,
  cellToA1,
  chunkToText,
  chunkWorkbook,
  dedupeOverlappingResults,
  indexWorkbook,
  rectToA1,
  rerankWorkbookResults,
  searchWorkbookRag,
  workbookFromSpreadsheetApi,
} from "../src/index.js";
import { fromBase64, toBase64 } from "../src/store/binaryStorage.js";
import type { VectorRecord, VectorSearchResult } from "../src/store/inMemoryVectorStore.js";
import type { WorkbookSearchResult } from "../src/retrieval/rankResults.js";
import {
  dedupeOverlappingResults as dedupeOverlappingResultsLegacy,
  rerankWorkbookResults as rerankWorkbookResultsLegacy,
} from "../src/retrieval/ranking.js";
import type { Rect } from "../src/workbook/rect.js";

// This file is intentionally not executed. It's a compilation target for the
// d.ts smoke test to ensure our hand-written declaration files match the
// runtime JS API surface (exports, method names, option shapes, etc).

async function smoke() {
  const abortController = new AbortController();

  const embedder = new HashEmbedder({ dimension: 8, cacheSize: 1000 });
  await embedder.embedTexts(["hello"], { signal: abortController.signal });

  const store = new InMemoryVectorStore({ dimension: embedder.dimension });
  const storeAsBinary: BinaryStorage = new InMemoryBinaryStorage();
  await storeAsBinary.save(new Uint8Array([0]));
  await storeAsBinary.load();
  await storeAsBinary.remove?.();
  void storeAsBinary;

  const record: VectorRecord = { id: "typed", vector: new Float32Array(embedder.dimension), metadata: { workbookId: "wb" } };
  const typedResults: VectorSearchResult[] = [{ id: "typed", score: 0.123, metadata: { workbookId: "wb" } }];
  void record;
  void typedResults;
  await store.upsert([
    { id: "a", vector: new Float32Array(embedder.dimension), metadata: { workbookId: "wb" } },
  ]);

  await store.list({
    workbookId: "wb",
    includeVector: false,
    signal: abortController.signal,
    filter: (_metadata, id) => id === "a",
  });

  await store.query(new Float32Array(embedder.dimension), 3, {
    workbookId: "wb",
    signal: abortController.signal,
    filter: (_metadata, id) => id.startsWith("a"),
  });

  await store.listContentHashes({ workbookId: "wb", signal: abortController.signal });

  await store.updateMetadata([{ id: "a", metadata: { workbookId: "wb", updated: true } }]);
  const deletedFromStore: number = await store.deleteWorkbook("wb");
  void deletedFromStore;
  await store.clear();
  await store.batch(async () => {
    await store.upsert([{ id: "b", vector: new Float32Array(embedder.dimension), metadata: { workbookId: "wb" } }]);
    await store.delete(["b"]);
  });

  const jsonStore = new JsonVectorStore({
    dimension: embedder.dimension,
    storage: new InMemoryBinaryStorage(),
    autoSave: false,
    resetOnCorrupt: true,
  });
  await jsonStore.load();
  await jsonStore.listContentHashes({ workbookId: "wb", signal: abortController.signal });
  await jsonStore.batch(async () => {
    await jsonStore.upsert([{ id: "a", vector: new Float32Array(embedder.dimension), metadata: { workbookId: "wb" } }]);
    await jsonStore.updateMetadata([{ id: "a", metadata: { workbookId: "wb", tag: "json" } }]);
  });
  await jsonStore.close();

  const sqliteStore = await SqliteVectorStore.create({
    dimension: embedder.dimension,
    storage: new InMemoryBinaryStorage(),
    autoSave: false,
    resetOnCorrupt: true,
    resetOnDimensionMismatch: true,
    locateFile: (file, prefix) => `${prefix ?? ""}${file}`,
  });

  // Error types are intentionally exported as type-only helpers so callers can
  // programmatically inspect errors thrown when reset options are disabled.
  function handleSqliteCreateError(err: unknown) {
    if ((err as any)?.name === "SqliteVectorStoreDimensionMismatchError") {
      const mismatch = err as SqliteVectorStoreDimensionMismatchError;
      const _dbDim: number = mismatch.dbDimension;
      const _requested: number = mismatch.requestedDimension;
      void _dbDim;
      void _requested;
    }
    if ((err as any)?.name === "SqliteVectorStoreInvalidMetadataError") {
      const invalid = err as SqliteVectorStoreInvalidMetadataError;
      const _raw: unknown = invalid.rawDimension;
      void _raw;
    }
  }
  handleSqliteCreateError(null);

  await sqliteStore.batch(async () => {
    await sqliteStore.upsert([
      { id: "a", vector: new Float32Array(embedder.dimension), metadata: { workbookId: "wb", tag: "sqlite" } },
    ]);
    await sqliteStore.delete(["a"]);
  });

  await sqliteStore.list({ includeVector: true, signal: abortController.signal, workbookId: "wb" });
  await sqliteStore.listContentHashes({ workbookId: "wb", signal: abortController.signal });
  await sqliteStore.get("a");
  await sqliteStore.query(new Float32Array(embedder.dimension), 3, { signal: abortController.signal });
  await sqliteStore.query(new Float32Array(embedder.dimension), 3, {
    workbookId: "wb",
    signal: abortController.signal,
    filter: (_metadata, _id) => true,
  });
  await sqliteStore.updateMetadata([{ id: "a", metadata: { workbookId: "wb", tag: "sqlite" } }]);
  const deletedFromSqlite: number = await sqliteStore.deleteWorkbook("wb");
  void deletedFromSqlite;
  await sqliteStore.clear();
  await sqliteStore.compact();
  await sqliteStore.vacuum();
  await sqliteStore.close();

  const localStorage = new LocalStorageBinaryStorage({ workbookId: "wb", namespace: "ns" });
  const key: string = localStorage.key;
  void key;
  await localStorage.save(new Uint8Array([1, 2, 3]));
  await localStorage.load();
  await localStorage.remove();

  const chunkedLocalStorage: BinaryStorage = new ChunkedLocalStorageBinaryStorage({ workbookId: "wb", namespace: "ns" });
  await chunkedLocalStorage.save(new Uint8Array([1, 2, 3]));
  await chunkedLocalStorage.load();
  await chunkedLocalStorage.remove?.();

  const indexedDbStorage: BinaryStorage = new IndexedDBBinaryStorage({ workbookId: "wb", namespace: "ns", dbName: "db" });
  await indexedDbStorage.save(new Uint8Array([1, 2, 3]));
  await indexedDbStorage.load();
  await indexedDbStorage.remove?.();

  const encoded: string = toBase64(new Uint8Array([1, 2, 3]));
  const decoded: Uint8Array = fromBase64(encoded);
  void decoded;

  const workbook = { id: "wb", sheets: [{ name: "Sheet1", cells: [[{ v: "hello" }]] }] };
  const chunks = chunkWorkbook(workbook, {
    signal: abortController.signal,
    extractMaxRows: 10,
    extractMaxCols: 10,
    detectRegionsCellLimit: 1000,
    maxRegionsPerSheet: 5,
    maxDataRegionsPerSheet: 5,
    maxFormulaRegionsPerSheet: 5,
  });
  const text = chunkToText(chunks[0], { sampleRows: 1, maxColumnsForSchema: 8, maxColumnsForRows: 8 });
  const tokenCount: number = approximateTokenCount(text);
  void tokenCount;

  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    sampleRows: 2,
    maxColumnsForSchema: 8,
    maxColumnsForRows: 8,
    tokenCount: approximateTokenCount,
    embedBatchSize: 16,
    onProgress: (_info) => {},
    signal: abortController.signal,
    transform: async (record) => {
      if (record.id === "skip") return null;
      // The implementation accepts `text: null` and coerces to "".
      return { text: null, metadata: { ...record.metadata, transformed: true } };
    },
  });

  // Object-literal embedder support: ensure `embedder.name` is allowed in the public types
  // (runtime reads it when present).
  const objectEmbedder = {
    name: "object-embedder",
    async embedTexts(texts: string[], _options?: { signal?: AbortSignal }) {
      return texts.map(() => new Float32Array(embedder.dimension));
    },
  };
  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder: objectEmbedder,
  });

  const rect: Rect = { r0: 0, c0: 0, r1: 0, c1: 0 };
  rectToA1(rect);
  cellToA1(0, 0);

  const rerankedExample = rerankWorkbookResults("hello", [
    { id: "a", score: 0.5, metadata: { workbookId: "wb", sheetName: "Sheet1", kind: "table", title: "Hello" } },
    { id: "b", score: 0.5, metadata: { workbookId: "wb", sheetName: "Sheet1", kind: "dataRegion", title: "Other" } },
  ]);
  const dedupedExample = dedupeOverlappingResults(rerankedExample, { overlapRatioThreshold: 0.8 });
  void dedupedExample;

  await searchWorkbookRag({
    queryText: "hello",
    topK: 3,
    vectorStore: store,
    embedder,
    rerank: true,
    dedupe: true,
    signal: abortController.signal,
  });

  workbookFromSpreadsheetApi({
    workbookId: "wb",
    spreadsheet: {
      listSheets: () => ["Sheet1"],
      listNonEmptyCells: () => [],
    },
    coordinateBase: "auto",
    signal: abortController.signal,
  });

  const results = await searchWorkbookRag({
    queryText: "hello",
    workbookId: "wb",
    topK: 5,
    vectorStore: store,
    embedder,
    rerank: true,
    dedupe: true,
    signal: abortController.signal,
  });
  const workbookResults: WorkbookSearchResult[] = results;
  void workbookResults;
  const reranked = rerankWorkbookResults("hello", results);
  const deduped = dedupeOverlappingResults(reranked);
  void deduped;

  // Backwards-compatible wrappers (object-shaped params).
  const rerankedLegacy = rerankWorkbookResultsLegacy({ queryText: "hello", results });
  const dedupedLegacy = dedupeOverlappingResultsLegacy({ results: rerankedLegacy, overlapRatio: 0.8 });
  void dedupedLegacy;
}

void smoke;
