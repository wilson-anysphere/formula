import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { InMemoryBinaryStorage } from "../src/store/binaryStorage.js";
import { JsonVectorStore } from "../src/store/jsonVectorStore.js";
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

let sqlJsAvailable = true;
try {
  // Keep this as a computed dynamic import (no literal bare specifier) so
  // `scripts/run-node-tests.mjs` can still execute this file when `node_modules/`
  // is missing.
  const sqlJsModuleName = "sql" + ".js";
  await import(sqlJsModuleName);
} catch {
  sqlJsAvailable = false;
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

test("indexWorkbook persists metadata-only changes without re-embedding", async () => {
  const workbook = makeWorkbook();
  const baseEmbedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  let embedCalls = 0;
  const embedder = {
    name: baseEmbedder.name,
    async embedTexts(texts, options) {
      embedCalls += 1;
      return baseEmbedder.embedTexts(texts, options);
    },
  };

  const first = await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    transform: (record) => ({ metadata: { ...record.metadata, tag: "v1" } }),
  });
  assert.ok(first.totalChunks > 0);
  assert.equal(first.upserted, first.totalChunks);

  const second = await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    transform: (record) => ({ metadata: { ...record.metadata, tag: "v2" } }),
  });
  assert.equal(second.upserted, 0);
  assert.equal(second.deleted, 0);
  assert.equal(embedCalls, 1);

  const records = await store.list({ workbookId: workbook.id, includeVector: false });
  assert.ok(records.length > 0);
  for (const r of records) {
    assert.equal(r.metadata.tag, "v2");
  }
});

test("indexWorkbook persists metadata-only changes for JsonVectorStore", async () => {
  const workbook = makeWorkbook();
  const baseEmbedder = new HashEmbedder({ dimension: 128 });
  const storage = new InMemoryBinaryStorage();
  const store = new JsonVectorStore({ storage, dimension: 128, autoSave: true });

  let embedCalls = 0;
  const embedder = {
    name: baseEmbedder.name,
    async embedTexts(texts, options) {
      embedCalls += 1;
      return baseEmbedder.embedTexts(texts, options);
    },
  };

  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    transform: (record) => ({ metadata: { ...record.metadata, tag: "v1" } }),
  });

  const second = await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    transform: (record) => ({ metadata: { ...record.metadata, tag: "v2" } }),
  });
  assert.equal(second.upserted, 0);
  assert.equal(second.deleted, 0);
  assert.equal(embedCalls, 1);

  await store.close();

  const store2 = new JsonVectorStore({ storage, dimension: 128, autoSave: false });
  const records = await store2.list({ workbookId: workbook.id, includeVector: false });
  assert.ok(records.length > 0);
  for (const r of records) {
    assert.equal(r.metadata.tag, "v2");
  }
  await store2.close();
});

test(
  "indexWorkbook persists metadata-only changes for SqliteVectorStore",
  { skip: !sqlJsAvailable },
  async () => {
    const workbook = makeWorkbook();
    const baseEmbedder = new HashEmbedder({ dimension: 128 });
    const storage = new InMemoryBinaryStorage();
    const modulePath = "../src/store/" + "sqliteVectorStore.js";
    const { SqliteVectorStore } = await import(modulePath);
    const store = await SqliteVectorStore.create({ storage, dimension: 128, autoSave: true });

    let embedCalls = 0;
    const embedder = {
      name: baseEmbedder.name,
      async embedTexts(texts, options) {
        embedCalls += 1;
        return baseEmbedder.embedTexts(texts, options);
      },
    };

    await indexWorkbook({
      workbook,
      vectorStore: store,
      embedder,
      transform: (record) => ({ metadata: { ...record.metadata, tag: "v1" } }),
    });

    const second = await indexWorkbook({
      workbook,
      vectorStore: store,
      embedder,
      transform: (record) => ({ metadata: { ...record.metadata, tag: "v2" } }),
    });
    assert.equal(second.upserted, 0);
    assert.equal(second.deleted, 0);
    assert.equal(embedCalls, 1);

    await store.close();

    const store2 = await SqliteVectorStore.create({ storage, dimension: 128, autoSave: false });
    const records = await store2.list({ workbookId: workbook.id, includeVector: false });
    assert.ok(records.length > 0);
    for (const r of records) {
      assert.equal(r.metadata.tag, "v2");
      assert.equal(r.metadata.embedder, embedder.name);
    }

    // Ensure metadata_json only contains *extra* keys (and does not include standard fields)
    // after a metadata-only update path.
    const sampleId = records[0].id;
    const stmt = store2._db.prepare(
      `
        SELECT
          workbook_id,
          sheet_name,
          kind,
          title,
          r0,
          c0,
          r1,
          c1,
          content_hash,
          metadata_hash,
          token_count,
          text,
          metadata_json
        FROM vectors
        WHERE id = ?
        LIMIT 1;
      `
    );
    stmt.bind([sampleId]);
    assert.ok(stmt.step());
    const row = stmt.get();
    stmt.free();

    assert.equal(row[0], workbook.id);
    // Structured fields should be populated (not stored redundantly in metadata_json).
    assert.ok(typeof row[8] === "string" && row[8].length > 0);
    assert.ok(typeof row[9] === "string" && row[9].length > 0);
    assert.ok(Number.isFinite(row[10]) && row[10] > 0);
    assert.ok(typeof row[11] === "string" && row[11].length > 0);

    const extra = JSON.parse(row[12]);
    assert.equal(extra.tag, "v2");
    assert.equal(extra.embedder, embedder.name);
    for (const key of [
      "workbookId",
      "sheetName",
      "kind",
      "title",
      "rect",
      "text",
      "contentHash",
      "metadataHash",
      "tokenCount",
    ]) {
      assert.equal(Object.prototype.hasOwnProperty.call(extra, key), false, `metadata_json should not include ${key}`);
    }
    await store2.close();
  }
);

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

test("indexWorkbook uses vectorStore.listContentHashes when available", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });

  let listCalls = 0;
  let listHashCalls = 0;
  let upserts = 0;

  const store = {
    dimension: 128,
    async list() {
      listCalls += 1;
      throw new Error("indexWorkbook should prefer listContentHashes when available");
    },
    async listContentHashes() {
      listHashCalls += 1;
      return [];
    },
    async upsert(records) {
      upserts += records.length;
    },
    async delete() {
      throw new Error("Unexpected delete");
    },
  };

  const res = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.ok(res.totalChunks > 0);
  assert.equal(listCalls, 0);
  assert.equal(listHashCalls, 1);
  assert.equal(upserts, res.upserted);
});

test("indexWorkbook batching respects AbortSignal and avoids partial writes", async () => {
  const workbook = makeWorkbookTwoTables();
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
    indexWorkbook({
      workbook,
      vectorStore: store,
      embedder,
      embedBatchSize: 1,
      signal: abortController.signal,
    }),
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

test("indexWorkbook does not early-abort while awaiting vectorStore updateMetadata persistence", async () => {
  const workbook = makeWorkbook();
  const abortController = new AbortController();

  /** @type {Map<string, { id: string, metadata: any }>} */
  const records = new Map();

  const updateCalled = defer();
  const updateDone = defer();
  let updateWasAwaited = false;
  /** @type {any[]} */
  let updateRecords = [];

  const vectorStore = {
    async list(opts) {
      const workbookId = opts?.workbookId;
      return Array.from(records.values()).filter((r) => (workbookId ? r.metadata?.workbookId === workbookId : true));
    },
    async upsert(items) {
      for (const item of items) records.set(item.id, { id: item.id, metadata: item.metadata });
    },
    updateMetadata(items) {
      updateRecords = items;
      updateCalled.resolve();
      return {
        then(onFulfilled, onRejected) {
          updateWasAwaited = true;
          return updateDone.promise
            .then(() => {
              for (const item of updateRecords) {
                const existing = records.get(item.id);
                if (!existing) continue;
                existing.metadata = item.metadata;
              }
            })
            .then(onFulfilled, onRejected);
        },
      };
    },
    async delete() {
      throw new Error("Unexpected delete() during test");
    },
  };

  const embedder = {
    name: "stub-embedder",
    async embedTexts(texts) {
      return texts.map(() => new Float32Array(8));
    },
  };

  await indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({ metadata: { ...record.metadata, tag: "v1" } }),
  });

  const indexPromise = indexWorkbook({
    workbook,
    vectorStore,
    embedder,
    transform: (record) => ({ metadata: { ...record.metadata, tag: "v2" } }),
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

  // Wait until metadata persistence has started but keep it blocked.
  await updateCalled.promise;
  assert.ok(updateRecords.length > 0);

  // Give the await machinery a microtask turn to assimilate the thenable.
  await Promise.resolve();
  assert.equal(updateWasAwaited, true);

  // Abort while updateMetadata is still in flight. `indexWorkbook` should keep waiting for the store write
  // to finish, then reject on abort afterwards.
  abortController.abort();
  await Promise.resolve();
  assert.equal(settled, false);

  updateDone.resolve();
  await assert.rejects(indexPromise, { name: "AbortError" });

  // The metadata update should still have been applied.
  for (const rec of records.values()) {
    assert.equal(rec.metadata.tag, "v2");
  }
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

test("indexWorkbook treats non-finite embedBatchSize as Infinity (single embed call)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  let embedCalls = 0;
  const embedder = {
    async embedTexts(texts) {
      embedCalls += 1;
      return texts.map(() => new Float32Array(128));
    },
  };

  const result = await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    embedBatchSize: Number.NaN,
  });

  assert.equal(result.totalChunks, 2);
  assert.equal(result.upserted, 2);
  assert.equal(embedCalls, 1);
});

test("indexWorkbook reports upsert progress for metadata-only updates", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    transform: (record) => ({ metadata: { ...(record.metadata ?? {}), tag: "v1" } }),
  });

  /** @type {Array<{ phase: string, processed: number, total?: number }>} */
  const events = [];
  await indexWorkbook({
    workbook,
    vectorStore: store,
    embedder,
    transform: (record) => ({ metadata: { ...(record.metadata ?? {}), tag: "v2" } }),
    onProgress: (info) => events.push(info),
  });

  const upsertEvents = events.filter((e) => e.phase === "upsert");
  assert.ok(upsertEvents.length >= 2);
  const last = upsertEvents[upsertEvents.length - 1];
  assert.equal(last.processed, last.total);
  assert.ok((last.total ?? 0) > 0);
});

test("indexWorkbook reports progress across phases", async () => {
  const workbook = makeWorkbookTwoTables();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });
  await indexWorkbook({ workbook, vectorStore: store, embedder });

  // Change T1 so it must re-embed/upsert, and remove T2 to force a delete.
  const workbook2 = JSON.parse(JSON.stringify(workbook));
  workbook2.sheets[0].cells[1][0] = { v: 999 };
  workbook2.tables = workbook2.tables.filter((t) => t.name !== "T2");
  workbook2.sheets[0].cells[0][3] = null;
  workbook2.sheets[0].cells[0][4] = null;
  workbook2.sheets[0].cells[1][3] = null;
  workbook2.sheets[0].cells[1][4] = null;

  /** @type {Array<{ phase: string, processed: number, total?: number }>} */
  const events = [];
  await indexWorkbook({
    workbook: workbook2,
    vectorStore: store,
    embedder,
    embedBatchSize: 1,
    onProgress: (info) => events.push(info),
  });

  const phases = new Set(events.map((e) => e.phase));
  for (const phase of ["chunk", "hash", "embed", "upsert", "delete"]) {
    assert.ok(phases.has(phase), `Expected onProgress to include phase=${phase}`);
  }
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

test("indexWorkbook rejects when embedder returns a non-array result (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  let embedCalls = 0;
  const embedder = {
    async embedTexts() {
      embedCalls += 1;
      // Misbehaving embedder: returns a single vector instead of an array of vectors.
      return /** @type {any} */ (new Float32Array(128));
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore: store, embedder }),
    /returned a non-array result/
  );

  assert.equal(embedCalls, 1);
  assert.deepEqual(await store.list({ workbookId: workbook.id, includeVector: false }), []);
});

test("indexWorkbook rejects when embedder returns an invalid vector entry (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();

  /** @type {any[]} */
  const upserted = [];
  const vectorStore = {
    async list() {
      return [];
    },
    async upsert(records) {
      upserted.push(...records);
    },
  };

  const embedder = {
    async embedTexts(texts) {
      return texts.map((_, i) => (i === 0 ? new Float32Array(8) : undefined));
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore, embedder }),
    /returned an invalid vector/
  );

  assert.deepEqual(upserted, []);
});

test("indexWorkbook rejects when embedder returns inconsistent vector lengths (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();

  /** @type {any[]} */
  const upserted = [];
  const vectorStore = {
    async list() {
      return [];
    },
    async upsert(records) {
      upserted.push(...records);
    },
  };

  const embedder = {
    async embedTexts(texts) {
      return texts.map((_, i) => new Float32Array(i === 0 ? 8 : 7));
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore, embedder }),
    /Vector dimension mismatch/
  );

  assert.deepEqual(upserted, []);
});

test("indexWorkbook rejects when embedder returns non-finite vector values (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  const embedder = {
    async embedTexts(texts) {
      return texts.map((_, i) => {
        const vec = new Float32Array(128);
        vec[0] = i === 0 ? 0 : Number.NaN;
        return vec;
      });
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore: store, embedder }),
    /invalid vector value/
  );

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

test("indexWorkbook rejects when a batched embedder call returns the wrong vector count (no partial writes)", async () => {
  const workbook = makeWorkbookTwoTables();
  const store = new InMemoryVectorStore({ dimension: 128 });

  let embedCalls = 0;
  const embedder = {
    async embedTexts(texts) {
      embedCalls += 1;
      if (embedCalls === 1) {
        // First batch is for a single text but returns two vectors.
        return [new Float32Array(128), new Float32Array(128)];
      }
      // Second batch returns none; overall count would have "matched" without per-batch validation.
      return [];
    },
  };

  await assert.rejects(
    indexWorkbook({ workbook, vectorStore: store, embedder, embedBatchSize: 1 }),
    /returned 2 vector\(s\); expected 1/
  );

  assert.equal(embedCalls, 1);
  assert.deepEqual(await store.list({ workbookId: workbook.id, includeVector: false }), []);
});
