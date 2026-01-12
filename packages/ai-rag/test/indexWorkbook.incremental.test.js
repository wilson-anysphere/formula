import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";

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
