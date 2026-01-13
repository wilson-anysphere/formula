import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";

function defer() {
  /** @type {(value?: any) => void} */
  let resolve;
  /** @type {(reason?: any) => void} */
  let reject;
  const promise = new Promise((res, rej) => {
    resolve = res;
    reject = rej;
  });
  return { promise, resolve, reject };
}

function makeWorkbook() {
  return {
    id: "wb-incremental",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ v: "Region" }, { v: "Revenue" }],
          [{ v: "North" }, { v: 100 }],
          [{ v: "South" }, { v: 200 }],
        ],
      },
    ],
    tables: [{ name: "RevenueByRegion", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 1 } }],
  };
}

function makeWorkbookTwoTables() {
  return {
    id: "wb-incremental-two",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ v: "A" }, { v: "B" }, null, { v: "C" }, { v: "D" }],
          [{ v: 1 }, { v: 2 }, null, { v: 3 }, { v: 4 }],
        ],
      },
    ],
    tables: [
      { name: "T1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } },
      { name: "T2", sheetName: "Sheet1", rect: { r0: 0, c0: 3, r1: 1, c1: 4 } },
    ],
  };
}

test("indexWorkbook is incremental (unchanged chunks are skipped)", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const first = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.ok(first.totalChunks > 0);
  assert.equal(first.upserted, first.totalChunks);

  const second = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.equal(second.upserted, 0);
  assert.equal(second.skipped, second.totalChunks);
  assert.equal(second.deleted, 0);
});

test("indexWorkbook re-embeds all chunks when embedder identity changes", async () => {
  const workbook = makeWorkbook();
  const store = new InMemoryVectorStore({ dimension: 128 });

  const embedderV2 = new HashEmbedder({ dimension: 128 });
  assert.equal(embedderV2.name, "hash:v2:128");
  const first = await indexWorkbook({ workbook, vectorStore: store, embedder: embedderV2 });
  assert.ok(first.totalChunks > 0);
  assert.equal(first.upserted, first.totalChunks);

  class HashEmbedderV3 extends HashEmbedder {
    get name() {
      return "hash:v3:128";
    }
  }
  const embedderV3 = new HashEmbedderV3({ dimension: 128 });
  const second = await indexWorkbook({ workbook, vectorStore: store, embedder: embedderV3 });
  assert.equal(second.upserted, second.totalChunks);
  assert.equal(second.skipped, 0);
  assert.equal(second.deleted, 0);
});

test("indexWorkbook only upserts changed chunks", async () => {
  const workbook = makeWorkbookTwoTables();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const first = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.equal(first.totalChunks, 2);
  assert.equal(first.upserted, 2);

  const id1 = `${workbook.id}::Sheet1::table::T1`;
  const id2 = `${workbook.id}::Sheet1::table::T2`;
  const rec1a = await store.get(id1);
  const rec2a = await store.get(id2);
  assert.ok(rec1a);
  assert.ok(rec2a);
  const hash1a = rec1a.metadata.contentHash;
  const hash2a = rec2a.metadata.contentHash;

  // Change only table T2.
  const workbook2 = JSON.parse(JSON.stringify(workbook));
  workbook2.sheets[0].cells[1][3] = { v: 999 };

  const second = await indexWorkbook({ workbook: workbook2, vectorStore: store, embedder });
  assert.equal(second.totalChunks, 2);
  assert.equal(second.upserted, 1);
  assert.equal(second.skipped, 1);
  assert.equal(second.deleted, 0);

  const rec1b = await store.get(id1);
  const rec2b = await store.get(id2);
  assert.ok(rec1b);
  assert.ok(rec2b);
  assert.equal(rec1b.metadata.contentHash, hash1a);
  assert.notEqual(rec2b.metadata.contentHash, hash2a);
  assert.match(rec2b.metadata.text, /\b999\b/);
});

test("indexWorkbook deletes stale chunks", async () => {
  const workbook = makeWorkbookTwoTables();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const first = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.equal(first.totalChunks, 2);
  assert.equal(first.upserted, 2);

  const id2 = `${workbook.id}::Sheet1::table::T2`;
  assert.ok(await store.get(id2));

  // Remove table T2 and its cells so no replacement chunk is generated.
  const workbook2 = JSON.parse(JSON.stringify(workbook));
  workbook2.tables = workbook2.tables.filter((t) => t.name !== "T2");
  workbook2.sheets[0].cells[0][3] = null;
  workbook2.sheets[0].cells[0][4] = null;
  workbook2.sheets[0].cells[1][3] = null;
  workbook2.sheets[0].cells[1][4] = null;

  const second = await indexWorkbook({ workbook: workbook2, vectorStore: store, embedder });
  assert.equal(second.totalChunks, 1);
  assert.equal(second.upserted, 0);
  assert.equal(second.deleted, 1);
  assert.equal(await store.get(id2), null);
});

test("indexWorkbook respects AbortSignal and avoids partial writes", async () => {
  const workbook = makeWorkbook();
  const store = new InMemoryVectorStore({ dimension: 128 });
  const abortController = new AbortController();

  let embedCalls = 0;
  const embedder = {
    async embedTexts(texts) {
      embedCalls += 1;
      abortController.abort();
      return texts.map(() => new Float32Array(128));
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore: store, embedder, signal: abortController.signal }),
    { name: "AbortError" }
  );

  assert.equal(embedCalls, 1);
  assert.deepEqual(await store.list({ workbookId: workbook.id, includeVector: false }), []);
});

test("indexWorkbook does not early-abort while awaiting vectorStore persistence", async () => {
  const workbook = makeWorkbook();
  const abortController = new AbortController();

  const upsertCalled = defer();
  const upsertDone = defer();
  let upsertWasAwaited = false;

  /** @type {any[]} */
  let upsertRecords = [];

  const vectorStore = {
    async list() {
      return [];
    },
    upsert(records) {
      upsertRecords = records;
      upsertCalled.resolve();

      // Return a thenable so the test can detect whether the write was actually awaited.
      return {
        then(onFulfilled, onRejected) {
          upsertWasAwaited = true;
          return upsertDone.promise.then(onFulfilled, onRejected);
        },
      };
    },
    async delete() {
      throw new Error("Unexpected delete() during test");
    },
  };

  const embedder = {
    async embedTexts(texts) {
      return texts.map(() => new Float32Array(8));
    },
  };

  const indexPromise = indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    signal: abortController.signal,
  });

  let settled = false;
  indexPromise.then(
    () => {
      settled = true;
    },
    () => {
      settled = true;
    }
  );

  // Wait until persistence has started (upsert called) but keep it blocked.
  await upsertCalled.promise;
  assert.ok(upsertRecords.length > 0);

  // Give the await machinery a microtask turn to assimilate the thenable.
  await Promise.resolve();
  assert.equal(upsertWasAwaited, true);

  // Abort while the upsert is still in flight. `indexWorkbook` should keep waiting for the
  // store write to finish, then reject on abort afterwards.
  abortController.abort();
  await Promise.resolve();
  assert.equal(settled, false);

  upsertDone.resolve();
  await assert.rejects(indexPromise, { name: "AbortError" });
});

test("indexWorkbook does not early-abort while awaiting vectorStore delete persistence", async () => {
  const workbook = makeWorkbook();
  const abortController = new AbortController();

  const deleteCalled = defer();
  const deleteDone = defer();
  let deleteWasAwaited = false;

  /** @type {string[] | undefined} */
  let deletedIds;

  const vectorStore = {
    async list() {
      // Include a stale record so the delete path runs.
      return [{ id: "stale-id", metadata: { workbookId: workbook.id, contentHash: "stale" } }];
    },
    async upsert() {
      // Upserts aren't the focus of this test; resolve immediately so we can block on delete.
    },
    delete(ids) {
      deletedIds = ids;
      deleteCalled.resolve();

      // Return a thenable so the test can detect whether the delete was actually awaited.
      return {
        then(onFulfilled, onRejected) {
          deleteWasAwaited = true;
          return deleteDone.promise.then(onFulfilled, onRejected);
        },
      };
    },
  };

  const embedder = {
    async embedTexts(texts) {
      return texts.map(() => new Float32Array(8));
    },
  };

  const indexPromise = indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    signal: abortController.signal,
  });

  let settled = false;
  indexPromise.then(
    () => {
      settled = true;
    },
    () => {
      settled = true;
    }
  );

  // Wait until deletion persistence has started but keep it blocked.
  await deleteCalled.promise;
  assert.deepEqual(deletedIds, ["stale-id"]);

  // Give the await machinery a microtask turn to assimilate the thenable.
  await Promise.resolve();
  assert.equal(deleteWasAwaited, true);

  // Abort while delete is still in flight. `indexWorkbook` should keep waiting for the store write
  // to finish, then reject on abort afterwards.
  abortController.abort();
  await Promise.resolve();
  assert.equal(settled, false);

  deleteDone.resolve();
  await assert.rejects(indexPromise, { name: "AbortError" });
});

test("indexWorkbook batches embedding requests (embedBatchSize=1)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  /** @type {string[][]} */
  const embedBatches = [];
  const embedder = {
    async embedTexts(texts) {
      embedBatches.push(texts);
      return texts.map(() => new Float32Array(128));
    },
  };

  const result = await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    embedBatchSize: 1,
  });

  assert.equal(result.totalChunks, 2);
  assert.equal(result.upserted, 2);
  assert.equal(embedBatches.length, 2);
  assert.deepEqual(
    embedBatches.map((b) => b.length),
    [1, 1]
  );

  const stored = await store.list({ workbookId: workbook.id, includeVector: false });
  assert.equal(stored.length, 2);
});

test("indexWorkbook rejects when embedder returns the wrong number of vectors (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  let embedCalls = 0;
  const embedder = {
    async embedTexts(texts) {
      embedCalls += 1;
      // Return fewer vectors than texts to simulate a misbehaving embedder.
      return texts.slice(0, -1).map(() => new Float32Array(128));
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore: store, embedder }),
    /returned 1 vector\(s\); expected 2/
  );

  assert.equal(embedCalls, 1);
  assert.deepEqual(await store.list({ workbookId: workbook.id, includeVector: false }), []);
});

test("indexWorkbook rejects when embedder vectors have the wrong dimension (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  const embedder = {
    async embedTexts(texts) {
      // First vector matches dimension, second does not.
      return texts.map((_, i) => new Float32Array(i === 0 ? 128 : 127));
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore: store, embedder }),
    /Vector dimension mismatch/
  );

  assert.deepEqual(await store.list({ workbookId: workbook.id, includeVector: false }), []);
});

test("indexWorkbook can be aborted from onProgress before persistence (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });
  const abortController = new AbortController();

  let sawUpsertStart = false;
  const onProgress = (info) => {
    if (info.phase === "upsert" && info.processed === 0) {
      sawUpsertStart = true;
      abortController.abort();
    }
  };

  const embedder = new HashEmbedder({ dimension: 128 });

  await assert.rejects(
    indexWorkbook({
      workbook,
      vectorStore: store,
      embedder,
      onProgress,
      signal: abortController.signal,
    }),
    { name: "AbortError" }
  );

  assert.equal(sawUpsertStart, true);
  assert.deepEqual(await store.list({ workbookId: workbook.id, includeVector: false }), []);
});
