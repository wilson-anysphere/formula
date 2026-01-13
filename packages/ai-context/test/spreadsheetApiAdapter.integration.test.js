import assert from "node:assert/strict";
import test from "node:test";

import { HashEmbedder } from "../../ai-rag/src/embedding/hashEmbedder.js";
import { InMemoryVectorStore } from "../../ai-rag/src/store/inMemoryVectorStore.js";
import { ContextManager } from "../src/contextManager.js";

test("integration: buildWorkbookContextFromSpreadsheetApi retrieves relevant chunk", async () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [
        { address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: "Region" } },
        { address: { sheet: "Sheet1", row: 1, col: 2 }, cell: { value: "Revenue" } },
        { address: { sheet: "Sheet1", row: 2, col: 1 }, cell: { value: "North" } },
        { address: { sheet: "Sheet1", row: 2, col: 2 }, cell: { value: 1000 } },
        { address: { sheet: "Sheet1", row: 3, col: 1 }, cell: { value: "South" } },
        { address: { sheet: "Sheet1", row: 3, col: 2 }, cell: { value: 2000 } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContextFromSpreadsheetApi({
    spreadsheet,
    workbookId: "wb-api",
    query: "revenue by region",
  });

  assert.match(out.promptContext, /## workbook_schema/i);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";
  assert.match(schemaSection, /Region\s*\(string\)/i);
  assert.match(schemaSection, /Revenue\s*\(number\)/i);

  assert.match(out.promptContext, /DATA REGION/i);
  assert.match(out.promptContext, /Region/);
  assert.match(out.promptContext, /Revenue/);
});

test("integration: workbook_schema is still emitted when retrieval is empty (SpreadsheetApi)", async () => {
  const spreadsheet = {
    listSheets() {
      return ["Sheet1"];
    },
    listNonEmptyCells(sheet) {
      assert.equal(sheet, "Sheet1");
      return [
        { address: { sheet: "Sheet1", row: 1, col: 1 }, cell: { value: "Region" } },
        { address: { sheet: "Sheet1", row: 1, col: 2 }, cell: { value: "Revenue" } },
        { address: { sheet: "Sheet1", row: 2, col: 1 }, cell: { value: "North" } },
        { address: { sheet: "Sheet1", row: 2, col: 2 }, cell: { value: 1000 } },
        { address: { sheet: "Sheet1", row: 3, col: 1 }, cell: { value: "South" } },
        { address: { sheet: "Sheet1", row: 3, col: 2 }, cell: { value: 2000 } },
      ];
    },
  };

  const embedder = new HashEmbedder({ dimension: 128 });
  const vectorStore = new InMemoryVectorStore({ dimension: 128 });

  const cm = new ContextManager({
    tokenBudgetTokens: 800,
    workbookRag: { vectorStore, embedder, topK: 3 },
  });

  const out = await cm.buildWorkbookContextFromSpreadsheetApi({
    spreadsheet,
    workbookId: "wb-api-empty-retrieval",
    query: "does not matter",
    topK: 0, // force no retrieval (but still index the workbook)
  });

  assert.equal(out.retrieved.length, 0);
  assert.match(out.promptContext, /## workbook_schema/i);
  const schemaSection =
    out.promptContext.match(/## workbook_schema\n([\s\S]*?)(?:\n\n## [^\n]+\n|$)/i)?.[1] ?? "";
  assert.match(schemaSection, /Region\s*\(string\)/i);
  assert.match(schemaSection, /Revenue\s*\(number\)/i);
});
