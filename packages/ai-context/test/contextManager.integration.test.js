import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { HashEmbedder } from "../../ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../ai-rag/src/store/inMemoryVectorStore.js";

function makeWorkbook() {
  return {
    id: "wb3",
    sheets: [
      {
        name: "Sales",
        cells: [
          [{ v: "Region" }, { v: "Revenue" }],
          [{ v: "North" }, { v: 1000 }],
          [{ v: "South" }, { v: 2000 }],
        ],
      },
    ],
    tables: [{ name: "RevenueByRegion", sheetName: "Sales", rect: { r0: 0, c0: 0, r1: 2, c1: 1 } }],
  };
}

test('integration: query "revenue by region" retrieves the right table chunk', async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContext({ workbook, query: "revenue by region" });

  assert.match(out.promptContext, /## workbook_schema/i);
  // Schema-first: table columns (headers + inferred types) should be present even when
  // retrieval is sparse/noisy.
  assert.match(out.promptContext, /Region\s*\(string\)/i);
  assert.match(out.promptContext, /Revenue\s*\(number\)/i);

  assert.match(out.promptContext, /RevenueByRegion/);
  assert.ok(out.retrieved.length > 0);
  assert.equal(out.retrieved[0].metadata.title, "RevenueByRegion");
});

test("buildWorkbookContext: workbook_schema is present even when no chunks are retrieved", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContext({
    workbook,
    query: "anything",
    skipIndexing: true, // keep vector store empty so retrieval returns no chunks
  });

  assert.equal(out.retrieved.length, 0);
  assert.match(out.promptContext, /## workbook_schema/i);
  assert.match(out.promptContext, /Region\s*\(string\)/i);
  assert.match(out.promptContext, /Revenue\s*\(number\)/i);
});

test("buildWorkbookContext: includePromptContext=false skips prompt formatting but still retrieves chunks", async () => {
  const workbook = makeWorkbook();
  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 500,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContext({ workbook, query: "revenue by region", includePromptContext: false });

  assert.equal(out.promptContext, "");
  assert.ok(out.retrieved.length > 0);
});
