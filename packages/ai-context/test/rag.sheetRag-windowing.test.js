import assert from "node:assert/strict";
import test from "node:test";
import { performance } from "node:perf_hooks";

import { chunkSheetByRegions, RagIndex, rangeToChunk } from "../src/rag.js";

test(
  "chunkSheetByRegions: streams TSV for tall regions (avoids per-row slice allocations)",
  { timeout: 10_000 },
  () => {
    const rows = 200_000;

    // `extractSheetSchema` only slices the first two rows for header heuristics.
    // Ensure later rows throw if `slice()` is called so this test detects regressions
    // back to `slice2D(...)+matrixToTsv(...)` for large regions.
    const row0 = ["x"];
    const row1 = ["x"];
    const throwingRow = ["x"];
    throwingRow.slice = () => {
      throw new Error("Unexpected row.slice() call while generating TSV chunk text");
    };

    const values = new Array(rows);
    values[0] = row0;
    values[1] = row1;
    for (let r = 2; r < rows; r++) values[r] = throwingRow;

    const sheet = { name: "Sheet1", values };

    const start = performance.now();
    const chunks = chunkSheetByRegions(sheet, { maxChunkRows: 10 });
    const took = performance.now() - start;

    assert.equal(chunks.length, 1);
    const text = chunks[0].text;

    const lines = text.split("\n");
    // 10 data lines + ellipsis
    assert.equal(lines.length, 11);
    assert.match(text, /… \(199990 more rows\)$/);

    // Generous threshold: catches order-of-magnitude regressions (e.g. accidental full-matrix copies).
    assert.ok(took < 5_000, `chunkSheetByRegions(200k x 1) took ${took.toFixed(1)}ms (expected < 5000ms)`);
  },
);

test(
  "RagIndex.indexSheet: splitRegions indexes multiple row windows for tall tables and improves search",
  { timeout: 10_000 },
  async () => {
    const values = [];
    for (let r = 0; r < 100; r++) values.push([r === 99 ? "specialtoken" : "filler"]);

    const sheet = { name: "Sheet1", values };
    const rag = new RagIndex();

    const { chunkCount } = await rag.indexSheet(sheet, {
      maxChunkRows: 10,
      splitRegions: true,
      chunkRowOverlap: 0,
      maxChunksPerRegion: 20,
    });

    // Expect more than one chunk when splitting is enabled for a tall region.
    assert.ok(chunkCount > 1, `expected > 1 chunk, got ${chunkCount}`);
    assert.equal(rag.store.size, chunkCount);

    const results = await rag.search("specialtoken", 3);
    assert.ok(results.length > 0);
    assert.match(results[0].preview, /\bspecialtoken\b/);
    // The token lives in the last row, so the best match should come from the last window.
    assert.equal(results[0].range, "Sheet1!A91:A100");
  },
);

test("rangeToChunk: preserves ragged row TSV formatting (no trailing tabs)", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      ["a", "b", "c"],
      ["d"],
    ],
  };

  const chunk = rangeToChunk(sheet, { startRow: 0, startCol: 0, endRow: 1, endCol: 2 }, { maxRows: 10 });
  assert.equal(chunk.text, "a\tb\tc\nd");
});
