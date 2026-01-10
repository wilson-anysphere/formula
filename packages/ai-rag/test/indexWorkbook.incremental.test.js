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

test("indexWorkbook is incremental (unchanged chunks are skipped)", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  const first = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.ok(first.upserted > 0);

  const second = await indexWorkbook({ workbook, vectorStore: store, embedder });
  assert.equal(second.upserted, 0);
  assert.equal(second.deleted, 0);
});

