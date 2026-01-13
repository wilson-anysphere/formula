import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder, InMemoryVectorStore, indexWorkbook, searchWorkbookRag } from "../src/index.js";

function makeRevenueWorkbook(id) {
  return {
    id,
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ v: "Region" }, { v: "Revenue" }, { v: "Units" }],
          [{ v: "North" }, { v: 1200 }, { v: 10 }],
          [{ v: "South" }, { v: 800 }, { v: 9 }],
        ],
      },
    ],
    tables: [{ name: "RevenueByRegion", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 2 } }],
  };
}

function makeSalariesWorkbook(id) {
  return {
    id,
    sheets: [
      {
        name: "SheetA",
        cells: [
          [{ v: "Employee" }, { v: "Salary" }],
          [{ v: "Alice" }, { v: 100000 }],
          [{ v: "Bob" }, { v: 120000 }],
        ],
      },
    ],
    tables: [{ name: "Salaries", sheetName: "SheetA", rect: { r0: 0, c0: 0, r1: 2, c1: 1 } }],
  };
}

test("searchWorkbookRag returns relevant chunks for a workbook", async () => {
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const wb1 = makeRevenueWorkbook("wb-search-1");
  const wb2 = makeSalariesWorkbook("wb-search-2");

  await indexWorkbook({ workbook: wb1, vectorStore: store, embedder });
  await indexWorkbook({ workbook: wb2, vectorStore: store, embedder });

  const results1 = await searchWorkbookRag({
    queryText: "revenue by region",
    workbookId: wb1.id,
    topK: 3,
    vectorStore: store,
    embedder,
  });
  assert.ok(results1.length > 0);
  assert.equal(results1[0].metadata.title, "RevenueByRegion");
  assert.ok(results1.every((r) => r.metadata.workbookId === wb1.id));

  const results2 = await searchWorkbookRag({
    queryText: "salary",
    workbookId: wb2.id,
    topK: 1,
    vectorStore: store,
    embedder,
  });
  assert.equal(results2.length, 1);
  assert.equal(results2[0].metadata.title, "Salaries");
  assert.ok(results2.every((r) => r.metadata.workbookId === wb2.id));
});

test("searchWorkbookRag honors workbookId filtering", async () => {
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const wb1 = makeRevenueWorkbook("wb-filter-1");
  const wb2 = makeSalariesWorkbook("wb-filter-2");

  await indexWorkbook({ workbook: wb1, vectorStore: store, embedder });
  await indexWorkbook({ workbook: wb2, vectorStore: store, embedder });

  const results = await searchWorkbookRag({
    queryText: "salary",
    workbookId: wb1.id,
    topK: 5,
    vectorStore: store,
    embedder,
  });

  assert.ok(results.length > 0);
  assert.ok(results.every((r) => r.metadata.workbookId === wb1.id));
  assert.ok(results.every((r) => r.metadata.title !== "Salaries"));
});

test("searchWorkbookRag reranks + dedupes by default (kind priority + overlap)", async () => {
  const calls = { queryTopK: 0 };
  const embedder = {
    async embedTexts() {
      return [[1, 0, 0]];
    },
  };

  const rect = { r0: 0, c0: 0, r1: 2, c1: 2 };

  const vectorStore = {
    /**
     * @param {ArrayLike<number>} vector
     * @param {number} topK
     * @param {{ workbookId?: string }} [opts]
     */
    async query(vector, topK, opts = {}) {
      calls.queryTopK = topK;
      assert.deepEqual(Array.from(vector), [1, 0, 0]);
      assert.equal(opts.workbookId, "wb");
      return [
        {
          id: "data",
          score: 0.51,
          metadata: {
            workbookId: "wb",
            sheetName: "Sheet1",
            kind: "dataRegion",
            title: "Data region A1:C3",
            rect,
            tokenCount: 50,
          },
        },
        {
          id: "table",
          score: 0.5,
          metadata: {
            workbookId: "wb",
            sheetName: "Sheet1",
            kind: "table",
            title: "RevenueByRegion",
            rect,
            tokenCount: 50,
          },
        },
      ];
    },
  };

  const results = await searchWorkbookRag({
    queryText: "revenue by region",
    workbookId: "wb",
    topK: 2,
    vectorStore,
    embedder,
  });

  assert.ok(calls.queryTopK > 2, "expected oversampled queryK when rerank/dedupe enabled");
  // After reranking, the table should bubble up; after dedupe, the overlapping dataRegion should be removed.
  assert.deepEqual(results.map((r) => r.id), ["table"]);
});

test("searchWorkbookRag respects AbortSignal", async () => {
  const embedder = new HashEmbedder({ dimension: 8 });
  const store = new InMemoryVectorStore({ dimension: 8 });

  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(
    searchWorkbookRag({
      queryText: "revenue",
      workbookId: "wb",
      topK: 1,
      vectorStore: store,
      embedder,
      signal: abortController.signal,
    }),
    { name: "AbortError" }
  );
});

test("searchWorkbookRag aborts while awaiting query embedding", async () => {
  const abortController = new AbortController();
  let embedCalled = false;
  let queryCalled = false;

  const embedder = {
    async embedTexts() {
      embedCalled = true;
      // Never resolve unless aborted.
      return new Promise(() => {});
    },
  };

  const vectorStore = {
    async query() {
      queryCalled = true;
      return [];
    },
  };

  const promise = searchWorkbookRag({
    queryText: "hello",
    workbookId: "wb",
    topK: 1,
    vectorStore,
    embedder,
    signal: abortController.signal,
  });

  assert.equal(embedCalled, true);
  abortController.abort();

  await assert.rejects(promise, { name: "AbortError" });
  assert.equal(queryCalled, false);
});

test("searchWorkbookRag aborts while awaiting vectorStore.query", async () => {
  const abortController = new AbortController();
  let queryCalled = false;
  let resolveQueryStarted = () => {};
  const queryStarted = new Promise((resolve) => {
    resolveQueryStarted = resolve;
  });

  const embedder = {
    async embedTexts() {
      return [[1, 0, 0]];
    },
  };

  const vectorStore = {
    async query() {
      queryCalled = true;
      resolveQueryStarted();
      // Never resolve unless aborted.
      return new Promise(() => {});
    },
  };

  const promise = searchWorkbookRag({
    queryText: "hello",
    workbookId: "wb",
    topK: 1,
    vectorStore,
    embedder,
    signal: abortController.signal,
  });

  await queryStarted;
  assert.equal(queryCalled, true);
  abortController.abort();

  await assert.rejects(promise, { name: "AbortError" });
});

test("searchWorkbookRag forwards workbookId + signal and avoids oversampling when rerank/dedupe are disabled", async () => {
  const abortController = new AbortController();
  const { signal } = abortController;

  /** @type {{ embed: number, query: number, queryTopK: number | null }} */
  const calls = { embed: 0, query: 0, queryTopK: null };

  const embedder = {
    /**
     * @param {string[]} texts
     * @param {{ signal?: AbortSignal }} [options]
     */
    async embedTexts(texts, options = {}) {
      calls.embed += 1;
      assert.deepEqual(texts, ["hello"]);
      assert.equal(options.signal, signal);
      return [[1, 0, 0]];
    },
  };

  const vectorStore = {
    /**
     * @param {ArrayLike<number>} vector
     * @param {number} topK
     * @param {{ workbookId?: string, signal?: AbortSignal }} [opts]
     */
    async query(vector, topK, opts = {}) {
      calls.query += 1;
      calls.queryTopK = topK;
      assert.deepEqual(Array.from(vector), [1, 0, 0]);
      assert.equal(opts.workbookId, "wb");
      assert.equal(opts.signal, signal);
      return [
        {
          id: "a",
          score: 1,
          metadata: { workbookId: "wb" },
        },
      ];
    },
  };

  const results = await searchWorkbookRag({
    queryText: "hello",
    workbookId: "wb",
    topK: 1,
    vectorStore,
    embedder,
    rerank: false,
    dedupe: false,
    signal,
  });

  assert.equal(calls.embed, 1);
  assert.equal(calls.query, 1);
  assert.equal(calls.queryTopK, 1);
  assert.equal(results.length, 1);
  assert.equal(results[0].id, "a");
});

test("searchWorkbookRag returns [] for topK<=0 without embedding or querying", async () => {
  let embedCalled = false;
  let queryCalled = false;

  const embedder = {
    async embedTexts() {
      embedCalled = true;
      return [[1, 0, 0]];
    },
  };

  const vectorStore = {
    async query() {
      queryCalled = true;
      return [];
    },
  };

  const results = await searchWorkbookRag({
    queryText: "hello",
    workbookId: "wb",
    topK: 0,
    vectorStore,
    embedder,
  });

  assert.deepEqual(results, []);
  assert.equal(embedCalled, false);
  assert.equal(queryCalled, false);
});

test("searchWorkbookRag returns [] for empty queryText without embedding or querying", async () => {
  let embedCalled = false;
  let queryCalled = false;

  const embedder = {
    async embedTexts() {
      embedCalled = true;
      return [[1, 0, 0]];
    },
  };

  const vectorStore = {
    async query() {
      queryCalled = true;
      return [];
    },
  };

  const results = await searchWorkbookRag({
    queryText: "   ",
    workbookId: "wb",
    topK: 3,
    vectorStore,
    embedder,
  });

  assert.deepEqual(results, []);
  assert.equal(embedCalled, false);
  assert.equal(queryCalled, false);
});

test("searchWorkbookRag requires workbookId when queryText is non-empty", async () => {
  const embedder = new HashEmbedder({ dimension: 8 });
  const vectorStore = new InMemoryVectorStore({ dimension: 8 });

  await assert.rejects(
    searchWorkbookRag({
      queryText: "hello",
      // workbookId omitted intentionally
      topK: 1,
      vectorStore,
      embedder,
    }),
    /workbookId/
  );
});
