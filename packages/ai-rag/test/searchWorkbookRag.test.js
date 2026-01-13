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

