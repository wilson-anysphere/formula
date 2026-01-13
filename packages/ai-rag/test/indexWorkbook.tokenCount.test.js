import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";

function makeWorkbook() {
  return {
    id: "wb-token-count",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ v: "A" }, { v: "B" }],
          [{ v: "hello" }, { v: "world" }],
        ],
      },
    ],
    tables: [{ name: "T1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
  };
}

test("indexWorkbook writes metadata.tokenCount using the provided tokenCount function", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const customTokenCount = () => 42;
  await indexWorkbook({ workbook, vectorStore: store, embedder, tokenCount: customTokenCount });

  const records = await store.list({ workbookId: workbook.id, includeVector: false });
  assert.ok(records.length > 0);
  for (const rec of records) {
    assert.equal(rec.metadata.tokenCount, 42);
  }
});

