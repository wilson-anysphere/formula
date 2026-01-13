import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../src/store/inMemoryVectorStore.js";
import { indexWorkbook } from "../src/pipeline/indexWorkbook.js";
import { searchWorkbookRag } from "../src/retrieval/searchWorkbookRag.js";

function makeWorkbook() {
  return {
    id: "wb2",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ v: "Region" }, { v: "Revenue" }, { v: "Units" }],
          [{ v: "North" }, { v: 1200 }, { v: 10 }],
          [{ v: "South" }, { v: 800 }, { v: 9 }],
        ],
      },
      {
        name: "Sheet2",
        cells: [
          [{ v: "Employee" }, { v: "Salary" }],
          [{ v: "Alice" }, { v: 100000 }],
          [{ v: "Bob" }, { v: 120000 }],
        ],
      },
    ],
    tables: [
      { name: "RevenueByRegion", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 2 } },
      { name: "Salaries", sheetName: "Sheet2", rect: { r0: 0, c0: 0, r1: 2, c1: 1 } },
    ],
  };
}

test('query "revenue by region" ranks the revenue table above unrelated tables', async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const store = new InMemoryVectorStore({ dimension: 128 });

  await indexWorkbook({ workbook, vectorStore: store, embedder });

  const results = await searchWorkbookRag({
    queryText: "revenue by region",
    workbookId: workbook.id,
    topK: 2,
    vectorStore: store,
    embedder,
    // Keep this test focused on vector similarity ranking.
    rerank: false,
    dedupe: false,
  });

  assert.equal(results[0].metadata.title, "RevenueByRegion");
});
