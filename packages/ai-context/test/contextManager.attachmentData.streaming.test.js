import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";

function makeStubRagIndex(sheetName = "Sheet1") {
  const schema = { name: sheetName, tables: [], namedRanges: [], dataRegions: [] };
  return {
    indexSheet: async () => ({ schema, chunkCount: 0 }),
    search: async () => [],
  };
}

test("buildContext: large attachment_data range preview is bounded + includes ellipsis", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (t) => t,
    ragIndex: makeStubRagIndex(),
    cacheSheetIndex: false,
  });

  const rowCount = 50_000;
  /** @type {unknown[][]} */
  const values = new Array(rowCount);
  for (let r = 0; r < 64; r++) {
    values[r] = [`r${r}c1`, `r${r}c2`, `r${r}c3`, `r${r}c4`, `r${r}c5`];
  }

  const out = await cm.buildContext({
    sheet: { name: "Sheet1", values },
    query: "anything",
    sampleRows: 0,
    attachments: [{ type: "range", reference: "A1:E50000" }],
    limits: {
      maxContextRows: rowCount,
      maxContextCells: rowCount * 5,
    },
  });

  assert.match(out.promptContext, /## attachment_data/i);
  assert.match(out.promptContext, /… \(49970 more rows\)/);
});

test("buildContext: aborts while generating attachment_data TSV preview", async () => {
  const abortController = new AbortController();

  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (t) => t,
    ragIndex: makeStubRagIndex(),
    cacheSheetIndex: false,
  });

  let stringified = 0;
  const value = {
    toString() {
      stringified += 1;
      if (stringified === 128) abortController.abort();
      if (stringified > 128 && abortController.signal.aborted) {
        throw new Error("Cell was stringified after abort; preview generation should check signal while iterating.");
      }
      return "x";
    },
  };

  const row = new Array(200).fill(value);

  const promise = cm.buildContext({
    sheet: { name: "Sheet1", values: [row] },
    query: "anything",
    sampleRows: 0,
    attachments: [{ type: "range", reference: "A1:GR1" }],
    limits: { maxContextRows: 1, maxContextCells: 200 },
    signal: abortController.signal,
  });

  await assert.rejects(promise, { name: "AbortError" });
});

test("buildContext: out-of-bounds attachments still mention the available sheet window", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000_000,
    redactor: (t) => t,
    ragIndex: makeStubRagIndex(),
    cacheSheetIndex: false,
  });

  const out = await cm.buildContext({
    sheet: {
      name: "Sheet1",
      values: [
        ["a", "b"],
        ["c", "d"],
      ],
    },
    query: "anything",
    sampleRows: 0,
    attachments: [{ type: "range", reference: "Z100:Z200" }],
  });

  assert.match(out.promptContext, /## attachment_data/i);
  assert.match(out.promptContext, /outside the available sheet window/i);
  assert.match(out.promptContext, /\(Sheet1!A1:B2\)/);
});

