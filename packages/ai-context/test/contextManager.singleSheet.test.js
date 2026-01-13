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

test("buildContext: concurrent calls share a single indexing pass", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  let releaseGate;
  const gate = new Promise((resolve) => {
    releaseGate = resolve;
  });
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    await gate;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const p1 = cm.buildContext({ sheet, query: "revenue by region" });
  const p2 = cm.buildContext({ sheet, query: "north revenue" });

  // Give both requests a chance to reach the indexer. Without per-sheet locking this would
  // call `indexSheet` twice before we release the gate.
  await new Promise((resolve) => setTimeout(resolve, 0));
  releaseGate();

  await Promise.all([p1, p2]);
  assert.equal(indexCalls, 1);
});

test("buildContext: aborts while waiting for an in-flight sheet index", async () => {
  const ragIndex = new RagIndex();
  let releaseGate;
  const gate = new Promise((resolve) => {
    releaseGate = resolve;
  });
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    await gate;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const p1 = cm.buildContext({ sheet, query: "revenue" });
  // Let p1 enter indexSheet and hold the per-sheet lock.
  await new Promise((resolve) => setTimeout(resolve, 0));

  const abortController = new AbortController();
  const p2 = cm.buildContext({ sheet, query: "north", signal: abortController.signal });
  // Let p2 start waiting for the lock before aborting.
  await new Promise((resolve) => setTimeout(resolve, 0));
  abortController.abort();

  await assert.rejects(p2, { name: "AbortError" });
  releaseGate();
  await p1;
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

test("buildContext: schema metadata updates (namedRanges) do not force re-indexing", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });
  const sheet = {
    name: "Sheet1",
    values: [
      ["Region", "Revenue"],
      ["North", 1000],
      ["South", 2000],
    ],
    namedRanges: [],
  };

  const out1 = await cm.buildContext({ sheet, query: "revenue" });
  assert.equal(indexCalls, 1);
  assert.deepEqual(out1.schema.namedRanges, []);

  // Update named range metadata (should update schema output, but not re-embed/re-index).
  sheet.namedRanges = [{ name: "MyRange", range: "Sheet1!A1:B1" }];
  const out2 = await cm.buildContext({ sheet, query: "revenue" });
  assert.equal(indexCalls, 1);
  assert.deepEqual(out2.schema.namedRanges, sheet.namedRanges);
});

test("buildContext: LRU cache eviction removes old sheet chunks from the in-memory store", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1000, sheetIndexCacheLimit: 2 });

  const sheetA = makeSheet(
    [
      ["Name", "Value"],
      ["A", 1],
    ],
    "SheetA",
  );
  const sheetB = makeSheet(
    [
      ["Name", "Value"],
      ["B", 2],
    ],
    "SheetB",
  );
  const sheetC = makeSheet(
    [
      ["Name", "Value"],
      ["C", 3],
    ],
    "SheetC",
  );

  await cm.buildContext({ sheet: sheetA, query: "a" });
  assert.equal(cm.ragIndex.store.size, 1);

  await cm.buildContext({ sheet: sheetB, query: "b" });
  assert.equal(cm.ragIndex.store.size, 2);

  // Indexing a 3rd distinct sheet should evict the oldest one and delete its chunks.
  await cm.buildContext({ sheet: sheetC, query: "c" });
  assert.equal(cm.ragIndex.store.size, 2);
});

test("clearSheetIndexCache: clears cached entries and can clear the in-memory store", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1000 });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  await cm.buildContext({ sheet, query: "revenue" });
  assert.equal(cm.ragIndex.store.size, 1);

  await cm.clearSheetIndexCache({ clearStore: true });
  assert.equal(cm.ragIndex.store.size, 0);

  // Subsequent calls should still work and re-index as needed.
  await cm.buildContext({ sheet, query: "revenue" });
  assert.equal(cm.ragIndex.store.size, 1);
});

test("clearSheetIndexCache: waits for in-flight indexing before clearing the store", async () => {
  const ragIndex = new RagIndex();
  let releaseGate;
  const gate = new Promise((resolve) => {
    releaseGate = resolve;
  });
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    await gate;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1000, ragIndex });
  const sheet = makeSheet([
    ["Region", "Revenue"],
    ["North", 1000],
    ["South", 2000],
  ]);

  const build = cm.buildContext({ sheet, query: "revenue" });
  // Let the buildContext call reach the gated indexSheet implementation.
  await new Promise((resolve) => setTimeout(resolve, 0));
  const clear = cm.clearSheetIndexCache({ clearStore: true });

  releaseGate();
  await Promise.allSettled([build, clear]);

  // If clear did not wait, the in-flight index could repopulate after the store was cleared.
  assert.equal(cm.ragIndex.store.size, 0);
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

test("buildContext: per-call maxContextRows override truncates rows", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextRows: 100 });
  const sheet = makeSheet([
    ["r1"],
    ["r2"],
    ["r3"],
    ["r4"],
    ["r5"],
  ]);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 100, limits: { maxContextRows: 2 } });
  assert.equal(out.sampledRows.length, 2);
  assert.equal(out.retrieved[0].range, "Sheet1!A1:A2");
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

test("buildContext: per-call maxContextCells override truncates columns", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextCells: 200_000 });
  const values = Array.from({ length: 10 }, (_v, r) => [`r${r}c1`, `r${r}c2`, `r${r}c3`]);
  const sheet = makeSheet(values);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 100, limits: { maxContextCells: 10 } });
  assert.equal(out.sampledRows.length, 10);
  assert.ok(out.sampledRows.every((row) => row.length === 1));
  assert.equal(out.retrieved[0].range, "Sheet1!A1:A10");
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

test("buildContext: maxChunkRows constructor option affects retrieved chunk previews", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxChunkRows: 2 });
  const sheet = makeSheet([["a"], ["b"], ["c"], ["d"], ["e"]]);

  const out = await cm.buildContext({ sheet, query: "anything" });
  const preview = out.retrieved[0].preview;
  assert.equal(preview.split("\n").length, 3);
  assert.match(preview, /… \(3 more rows\)$/);
});

test("buildContext: attached range previews stream rows (avoid slicing full ranges)", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 10_000 });

  const rowCount = 100;
  const values = [];
  for (let r = 0; r < rowCount; r++) {
    const row = [`Row${r + 1}`];
    let calls = 0;
    row.slice = () => {
      calls += 1;
      // ContextManager copies each row once when building the capped sheet window.
      // Any additional slice calls for rows beyond the preview limit indicate a regression
      // back to allocating full matrices for attachment previews.
      if (calls > 1 && r >= 30) {
        throw new Error(`Unexpected row.slice() call for row ${r} while building attachment preview`);
      }
      return row;
    };
    values.push(row);
  }

  const sheet = makeSheet(values);
  const out = await cm.buildContext({
    sheet,
    query: "anything",
    sampleRows: 0,
    attachments: [{ type: "range", reference: "Sheet1!A1:A100" }],
  });

  assert.match(out.promptContext, /Attached range data:/);
  assert.match(out.promptContext, /… \(70 more rows\)/);
});

test("buildContext: splitRegions indexes multiple row windows for tall sheets", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });
  const values = [];
  for (let r = 0; r < 100; r++) values.push([r === 99 ? "specialtoken" : "filler"]);
  const sheet = makeSheet(values);

  const out = await cm.buildContext({
    sheet,
    query: "specialtoken",
    limits: {
      maxChunkRows: 10,
      splitRegions: true,
      chunkRowOverlap: 0,
      maxChunksPerRegion: 20,
    },
  });

  assert.ok(cm.ragIndex.store.size > 1);
  assert.equal(out.retrieved[0].range, "Sheet1!A91:A100");
  assert.match(out.retrieved[0].preview, /\bspecialtoken\b/);
});

test("buildContext: splitRegions constructor option enables row-window retrieval without per-call limits", async () => {
  const cm = new ContextManager({
    tokenBudgetTokens: 1_000,
    maxChunkRows: 10,
    splitRegions: true,
    chunkRowOverlap: 0,
    maxChunksPerRegion: 20,
  });
  const values = [];
  for (let r = 0; r < 100; r++) values.push([r === 99 ? "specialtoken" : "filler"]);
  const sheet = makeSheet(values);

  const out = await cm.buildContext({ sheet, query: "specialtoken" });
  assert.ok(out.retrieved.length > 0);
  assert.equal(out.retrieved[0].range, "Sheet1!A91:A100");
  assert.match(out.retrieved[0].preview, /\bspecialtoken\b/);
});

test("buildContext: splitRegions default overlap/maxChunks match explicit values for cache signatures", async () => {
  const ragIndex = new RagIndex();
  let indexCalls = 0;
  const originalIndexSheet = ragIndex.indexSheet.bind(ragIndex);
  ragIndex.indexSheet = async (...args) => {
    indexCalls++;
    return originalIndexSheet(...args);
  };

  const cm = new ContextManager({ tokenBudgetTokens: 1_000, ragIndex });
  const values = [];
  for (let r = 0; r < 100; r++) values.push([`Row${r + 1}`]);
  const sheet = makeSheet(values);

  await cm.buildContext({
    sheet,
    query: "row",
    limits: { maxChunkRows: 10, splitRegions: true },
  });
  await cm.buildContext({
    sheet,
    query: "row",
    limits: { maxChunkRows: 10, splitRegions: true, chunkRowOverlap: 3, maxChunksPerRegion: 50 },
  });

  assert.equal(indexCalls, 1);
});

test("buildContext: negative maxContextRows/maxContextCells fall back to defaults", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, maxContextRows: -1, maxContextCells: -1 });
  const sheet = makeSheet([["r1"], ["r2"], ["r3"], ["r4"], ["r5"]]);

  const out = await cm.buildContext({ sheet, query: "anything", sampleRows: 100 });
  assert.equal(out.sampledRows.length, 5);
  assert.equal(out.retrieved[0].range, "Sheet1!A1:A5");
});

test("buildContext: sheetRagTopK option limits retrieved results", async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000, sheetRagTopK: 1 });
  const sheet = makeSheet([
    ["Region", "Revenue", "", "Cost"],
    ["North", 1000, "", 50],
    ["South", 2000, "", 60],
  ]);

  const out = await cm.buildContext({ sheet, query: "revenue" });
  assert.equal(out.retrieved.length, 1);
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
