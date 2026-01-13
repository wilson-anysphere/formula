import assert from "node:assert/strict";
import test from "node:test";

import { ContextManager } from "../src/contextManager.js";
import { RagIndex } from "../src/rag.js";

function makeSheet(values, name = "Sheet1") {
  return { name, values };
}

test("buildContext: repeated calls with identical data do not re-index the sheet", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const out1 = await cm.buildContext({ sheet, query: "revenue by region" });
  const size1 = cm.ragIndex.store.size;

  // Different query to ensure we still reuse the same indexed sheet.
  const out2 = await cm.buildContext({ sheet, query: "north revenue" });
  const size2 = cm.ragIndex.store.size;

  assert.equal(indexCalls, 1);
  assert.equal(size1, 1);
  assert.equal(size2, size1);
  assert.equal(out1.retrieved[0].range, "Sheet1!A1:B3");
  assert.equal(out2.retrieved[0].range, out1.retrieved[0].range);
});

test("buildContext: cacheSheetIndex=false re-indexes even when the sheet is unchanged", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex, cacheSheetIndex: false });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  await cm.buildContext({ sheet, query: "revenue by region" });
  await cm.buildContext({ sheet, query: "revenue by region" });

  assert.equal(indexCalls, 2);
});

test("buildContext: switching sheet.origin forces re-indexing for the active store", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });

  const baseValues = [
    ["Region"],
    ["North"],
    ["South"],
  ];

  const sheetOrigin0 = { name: "Sheet1", origin: { row: 0, col: 0 }, values: baseValues };
  const out0 = await cm.buildContext({ sheet: sheetOrigin0, query: "north" });
  assert.equal(indexCalls, 1);
  assert.equal(out0.retrieved[0].range, "Sheet1!A1:A3");

  const sheetOrigin10 = { name: "Sheet1", origin: { row: 10, col: 0 }, values: baseValues };
  const out10 = await cm.buildContext({ sheet: sheetOrigin10, query: "north" });
  assert.equal(indexCalls, 2);
  assert.equal(out10.retrieved[0].range, "Sheet1!A11:A13");

  // Same origin again => cache hit (no re-index).
  await cm.buildContext({ sheet: sheetOrigin10, query: "south" });
  assert.equal(indexCalls, 2);

  // Switching back to the previous origin requires re-indexing because the underlying
  // store can only hold one set of chunks per sheet name prefix at a time.
  const out0b = await cm.buildContext({ sheet: sheetOrigin0, query: "north" });
  assert.equal(indexCalls, 3);
  assert.equal(out0b.retrieved[0].range, "Sheet1!A1:A3");
});

test("buildContext: mutated sheet data triggers re-indexing and updates stored chunks", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const out1 = await cm.buildContext({ sheet, query: "revenue by region" });
  assert.equal(indexCalls, 1);
  assert.match(out1.retrieved[0].preview, /\b1000\b/);

  // Mutate a value in-place; signature should change and force re-indexing.
  sheet.values[1][1] = 1111;
  const out2 = await cm.buildContext({ sheet, query: "revenue by region" });
  assert.equal(indexCalls, 2);
  assert.match(out2.retrieved[0].preview, /\b1111\b/);
});

test("buildContext: stale RAG chunks are removed when a sheet shrinks", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1000 });

  const twoRegions = makeSheet([
    ["Region", "Revenue", "", "Cost"],
    ["North", 1000, "", 50],
    ["South", 2000, "", 60],
  ]);
  await cm.buildContext({ sheet: twoRegions, query: "revenue" });
  assert.equal(cm.ragIndex.store.size, 2);

  const oneRegion = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);
  await cm.buildContext({ sheet: oneRegion, query: "revenue" });
  assert.equal(cm.ragIndex.store.size, 1);
});

test("buildContext: retrieved context is preserved under tight token budgets", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 20 });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const out = await cm.buildContext({ sheet, query: "revenue by region" });
  assert.match(out.promptContext, /^## retrieved\b/m);
});

test("buildContext: caps matrix size to avoid Excel-scale allocations", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });

  // 1,000 rows x 300 cols => 300,000 cells (> 200,000 cap). The ContextManager should
  // truncate columns so downstream schema + sampling work remains bounded.
  const values = Array.from({ length: 1_000 }, (_v, r) => {
    const row = Array.from({ length: 300 }, () => null);
    row[0] = r;
    return row;
  });
  const sheet = makeSheet(values);

  const out = await cm.buildContext({ sheet, query: "col1", sampleRows: 1 });
  assert.equal(out.sampledRows.length, 1);
  assert.equal(out.sampledRows[0].length, 200);
});

test("buildContext: maxContextRows option truncates rows", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextRows: 3 });
  const sheet = makeSheet([
    ["r1"],
    ["r2"],
    ["r3"],
    ["r4"],
    ["r5"],
  ]);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 100 });
  assert.equal(out.sampledRows.length, 3);
  assert.equal(out.retrieved[0].range, "Sheet1!A1:A3");
});

test("buildContext: maxContextCells option truncates columns", async () => {
  // 10 rows x 5 cols => 50 cells. Cap to 20 total => 2 cols per row (floor(20/10)=2).
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextCells: 20 });
  const values = Array.from({ length: 10 }, (_v, r) => [`r${r}c1`, `r${r}c2`, `r${r}c3`, `r${r}c4`, `r${r}c5`]);
  const sheet = makeSheet(values);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 100 });
  assert.equal(out.sampledRows.length, 10);
  assert.ok(out.sampledRows.every((row) => row.length === 2));
  assert.equal(out.retrieved[0].range, "Sheet1!A1:B10");
});

test("buildContext: maxContextCells also caps rows when maxContextRows is larger", async () => {
  // 50 rows x 2 cols => 100 cells. Cap to 10 total => 10 rows x 1 col.
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextRows: 100, maxContextCells: 10 });
  const values = Array.from({ length: 50 }, (_v, r) => [`r${r}c1`, `r${r}c2`]);
  const sheet = makeSheet(values);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 1000 });
  assert.equal(out.sampledRows.length, 10);
  assert.ok(out.sampledRows.every((row) => row.length === 1));
  assert.equal(out.retrieved[0].range, "Sheet1!A1:A10");
});

test("buildContext: maxChunkRows affects sheet-level retrieved chunk previews", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });
  const sheet = makeSheet([["a"], ["b"], ["c"], ["d"], ["e"]]);

  const outSmall = await cm.buildContext({
    sheet,
    query: "anything",
    limits: { maxChunkRows: 2 },
  });
  const previewSmall = outSmall.retrieved[0].preview;
  assert.equal(previewSmall.split("\n").length, 3);
  assert.match(previewSmall, /… \(3 more rows\)$/);

  const outLarge = await cm.buildContext({
    sheet,
    query: "anything",
    limits: { maxChunkRows: 5 },
  });
  const previewLarge = outLarge.retrieved[0].preview;
  assert.equal(previewLarge.split("\n").length, 5);
  assert.doesNotMatch(previewLarge, /… \(/);
});

test("buildContext: negative maxContextRows/maxContextCells fall back to defaults", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextRows: -1, maxContextCells: -1 });
  const sheet = makeSheet([["r1"], ["r2"], ["r3"], ["r4"], ["r5"]]);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 100 });
  assert.equal(out.sampledRows.length, 5);
  assert.equal(out.retrieved[0].range, "Sheet1!A1:A5");
});

test("buildContext: respects AbortSignal", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const abortController = new AbortController();
  abortController.abort();

  await assert.rejects(cm.buildContext({ sheet, query: "revenue", signal: abortController.signal }), {
    name: "AbortError",
  });
});
