import assert from "node:assert/strict";
import test from "node:test";
import { performance } from "node:perf_hooks";

import { extractSheetSchema } from "../src/schema.js";
import { ContextManager } from "../src/contextManager.js";
import { chunkSheetByRegions } from "../src/rag.js";

function elapsedMs(start) {
  return performance.now() - start;
}

function makeDenseMatrix(rows, cols) {
  const values = new Array(rows);

  const header = new Array(cols);
  for (let c = 0; c < cols; c++) header[c] = `Col${c + 1}`;
  values[0] = header;

  for (let r = 1; r < rows; r++) {
    const row = new Array(cols);
    // Keep values small; we're stress-testing traversal/allocations, not content size.
    for (let c = 0; c < cols; c++) row[c] = c;
    values[r] = row;
  }

  return values;
}

test("stress: extractSheetSchema completes on a 1000x200 dense matrix", { timeout: 10_000 }, () => {
  const sheet = { name: "Sheet1", values: makeDenseMatrix(1_000, 200) };

  const start = performance.now();
  const schema = extractSheetSchema(sheet);
  const took = elapsedMs(start);

  assert.equal(schema.name, "Sheet1");
  assert.ok(schema.dataRegions.length >= 1);

  // Generous threshold: intended to catch order-of-magnitude regressions (e.g. quadratic queueing),
  // not minor perf noise across CI runners.
  assert.ok(
    took < 2_000,
    `extractSheetSchema(1000x200) took ${took.toFixed(1)}ms (expected < 2000ms)`,
  );
});

test(
  "stress: extractSheetSchema does not attempt Excel-scale visited allocations on sparse arrays",
  { timeout: 10_000 },
  () => {
    // Simulate an Excel full-sheet selection (1,048,576 rows x 16,384 columns) without actually
    // materializing the full matrix in memory.
    //
    // This should stay bounded via `detectDataRegions` scan caps. If a future regression tries to
    // allocate `rows*cols` visited grids, it will throw (invalid typed array length / OOM) and fail.
    const values = new Array(1_048_576);
    const wideRow = new Array(16_384);
    wideRow[0] = "x";
    values[0] = wideRow;

    const start = performance.now();
    const schema = extractSheetSchema({ name: "Sheet1", values });
    const took = elapsedMs(start);

    assert.equal(schema.name, "Sheet1");
    assert.ok(schema.dataRegions.length >= 1);
    assert.ok(
      took < 2_000,
      `extractSheetSchema(excel-scale sparse) took ${took.toFixed(1)}ms (expected < 2000ms)`,
    );
  },
);

test("stress: ContextManager.buildContext completes and respects caps", { timeout: 10_000 }, async () => {
  const cm = new ContextManager({ tokenBudgetTokens: 1_000 });

  // Simulate Excel-scale row counts without allocating the whole grid. Only the first 1,000 rows
  // should be considered for context building.
  const rawValues = new Array(1_048_576);
  const wideCols = 400; // > 200_000 / 1_000 => should truncate to 200 columns.
  for (let r = 0; r < 1_000; r++) {
    const row = new Array(wideCols);
    // Keep the region connected and wide via the first row.
    if (r === 0) {
      for (let c = 0; c < wideCols; c++) row[c] = `H${c + 1}`;
    } else {
      row[0] = r;
    }
    rawValues[r] = row;
  }

  const sheet = { name: "Sheet1", values: rawValues };
  const start = performance.now();
  const out = await cm.buildContext({ sheet, query: "header", sampleRows: 1 });
  const took = elapsedMs(start);

  assert.equal(out.sampledRows.length, 1);
  assert.ok(out.sampledRows[0].length <= 200);
  assert.ok(out.schema?.dataRegions?.[0]?.columnCount <= 200);
  assert.ok(
    took < 3_000,
    `ContextManager.buildContext(excel-scale sparse) took ${took.toFixed(1)}ms (expected < 3000ms)`,
  );
});

test("stress: chunkSheetByRegions returns bounded-size chunk text", { timeout: 10_000 }, () => {
  const sheet = { name: "Sheet1", values: makeDenseMatrix(1_000, 200) };

  const start = performance.now();
  const chunks = chunkSheetByRegions(sheet, { maxChunkRows: 20 });
  const took = elapsedMs(start);

  assert.equal(chunks.length, 1);
  const chunk = chunks[0];
  const lines = chunk.text.split("\n");
  // 20 data lines + an ellipsis line when the region has more rows.
  assert.ok(lines.length <= 21, `expected <= 21 lines, got ${lines.length}`);
  assert.match(chunk.text, /\bmore rows\)$/);

  assert.ok(
    took < 2_000,
    `chunkSheetByRegions(1000x200) took ${took.toFixed(1)}ms (expected < 2000ms)`,
  );
});

