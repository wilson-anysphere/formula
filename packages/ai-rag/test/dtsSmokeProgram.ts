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

// This file is intentionally not executed. It's a compilation target for the
// d.ts smoke test to ensure our hand-written declaration files match the
// runtime JS API surface (exports, method names, option shapes, etc).

async function smoke() {
  const abortController = new AbortController();

  const embedder = new HashEmbedder({ dimension: 8 });
  await embedder.embedTexts(["hello"], { signal: abortController.signal });

  const store = new InMemoryVectorStore({ dimension: embedder.dimension });
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
  await sqliteStore.list({ includeVector: true, signal: abortController.signal, workbookId: "wb" });
  await sqliteStore.query(new Float32Array(embedder.dimension), 3, { signal: abortController.signal });
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

  rectToA1({ r0: 0, c0: 0, r1: 0, c1: 0 });
  cellToA1(0, 0);

  const rerankedExample = rerankWorkbookResults("hello", [
    { id: "a", score: 0.5, metadata: { workbookId: "wb", sheetName: "Sheet1", kind: "table", title: "Hello" } },
    { id: "b", score: 0.5, metadata: { workbookId: "wb", sheetName: "Sheet1", kind: "dataRegion", title: "Other" } },
  ]);
  const dedupedExample = dedupeOverlappingResults(rerankedExample, { overlapRatioThreshold: 0.8 });
  void dedupedExample;

  await searchWorkbookRag({
    queryText: "hello",
    workbookId: "wb",
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
  const reranked = rerankWorkbookResults("hello", results);
  const deduped = dedupeOverlappingResults(reranked);
  void deduped;
}

void smoke;
