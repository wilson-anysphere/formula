import {
  HashEmbedder,
  InMemoryBinaryStorage,
  InMemoryVectorStore,
  JsonVectorStore,
  LocalStorageBinaryStorage,
  SqliteVectorStore,
  approximateTokenCount,
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

  const jsonStore = new JsonVectorStore({
    dimension: embedder.dimension,
    storage: new InMemoryBinaryStorage(),
    autoSave: false,
  });
  await jsonStore.load();
  await jsonStore.close();

  const sqliteStore = await SqliteVectorStore.create({
    dimension: embedder.dimension,
    storage: new InMemoryBinaryStorage(),
    autoSave: false,
    locateFile: (file, prefix) => `${prefix ?? ""}${file}`,
  });
  await sqliteStore.list({ includeVector: true, signal: abortController.signal });
  await sqliteStore.query(new Float32Array(embedder.dimension), 3, { signal: abortController.signal });
  await sqliteStore.close();

  const localStorage = new LocalStorageBinaryStorage({ workbookId: "wb", namespace: "ns" });
  const key: string = localStorage.key;
  void key;

  const encoded: string = toBase64(new Uint8Array([1, 2, 3]));
  const decoded: Uint8Array = fromBase64(encoded);
  void decoded;

  const workbook = { id: "wb", sheets: [{ name: "Sheet1", cells: [[{ v: "hello" }]] }] };
  const chunks = chunkWorkbook(workbook, { signal: abortController.signal });
  const text = chunkToText(chunks[0], { sampleRows: 1 });
  const tokenCount: number = approximateTokenCount(text);
  void tokenCount;

  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    sampleRows: 2,
    signal: abortController.signal,
    transform: async (record) => {
      if (record.id === "skip") return null;
      // The implementation accepts `text: null` and coerces to "".
      return { text: null, metadata: { ...record.metadata, transformed: true } };
    },
  });

  rectToA1({ r0: 0, c0: 0, r1: 0, c1: 0 });

  const reranked = rerankWorkbookResults("hello", [
    { id: "a", score: 0.5, metadata: { workbookId: "wb", sheetName: "Sheet1", kind: "table", title: "Hello" } },
    { id: "b", score: 0.5, metadata: { workbookId: "wb", sheetName: "Sheet1", kind: "dataRegion", title: "Other" } },
  ]);
  const deduped = dedupeOverlappingResults(reranked, { overlapRatioThreshold: 0.8 });
  void deduped;

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
}

void smoke;
