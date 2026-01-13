import assert from "node:assert/strict";
import test from "node:test";

import { chunkSheetByRegions, RagIndex, rangeToChunk } from "../src/rag.js";

test("chunkSheetByRegions: translates schema ranges to local indices when sheet.origin is set", () => {
  const sheet = {
    name: "My Sheet",
    origin: { row: 4, col: 2 }, // C5 (0-based)
    values: [
      ["Name", "Value"],
      ["Alice", 1],
      ["Bob", 2],
    ],
  };

  const chunks = chunkSheetByRegions(sheet);
  assert.equal(chunks.length, 1);
  assert.equal(chunks[0].range, "'My Sheet'!C5:D7");
  assert.match(chunks[0].text, /\bAlice\b/);
  assert.match(chunks[0].text, /\bBob\b/);
});

test("RagIndex.indexSheet: origin-offset sheets can be indexed and searched", async () => {
  const sheet = {
    name: "My Sheet",
    origin: { row: 4, col: 2 },
    values: [
      ["Name", "Value"],
      ["Alice", 1],
      ["Bob", 2],
    ],
  };

  const index = new RagIndex();
  await index.indexSheet(sheet);

  const results = await index.search("Alice", 3);
  assert.ok(results.length > 0);
  assert.equal(results[0].range, "'My Sheet'!C5:D7");
  assert.match(results[0].preview, /\bAlice\b/);
});

test("rangeToChunk: translates absolute coordinates to local indices when sheet.origin is set", () => {
  const sheet = {
    name: "My Sheet",
    origin: { row: 10, col: 5 }, // F11
    values: [
      ["A", "B"],
      ["C", "D"],
      ["E", "F"],
    ],
  };

  const chunk = rangeToChunk(
    sheet,
    { startRow: 10, startCol: 5, endRow: 12, endCol: 6 }, // F11:G13
    { maxRows: 10 },
  );

  assert.equal(chunk.range, "'My Sheet'!F11:G13");
  assert.match(chunk.text, /\bA\tB\b/);
  assert.match(chunk.text, /\bC\tD\b/);
});

test("rangeToChunk: clamps partially out-of-bounds absolute ranges to the provided matrix", () => {
  const sheet = {
    name: "My Sheet",
    origin: { row: 10, col: 5 }, // F11
    values: [
      ["A", "B"],
      ["C", "D"],
    ],
  };

  // Requested absolute range extends beyond the provided window, but intersects it.
  const chunk = rangeToChunk(
    sheet,
    { startRow: 9, startCol: 4, endRow: 12, endCol: 8 },
    { maxRows: 10 },
  );

  assert.match(chunk.text, /\bA\tB\b/);
  assert.match(chunk.text, /\bC\tD\b/);
});

test("chunk splitting: tall regions produce multiple chunks; shrinking removes stale ids", async () => {
  const makeSheet = (rows) => ({
    name: "Sheet1",
    values: Array.from({ length: rows }, (_v, rIdx) => [`Row${rIdx + 1}`, `Value${rIdx + 1}`]),
});

test("chunk splitting: repeats header row in later window chunks when a region has a header", () => {
  const values = [["Region", "Revenue"]];
  for (let i = 0; i < 20; i++) {
    values.push([`R${i + 1}`, i + 1]);
  }
  const sheet = { name: "Sheet1", values };

  const chunks = chunkSheetByRegions(sheet, { splitByRowWindows: true, maxChunkRows: 5, rowOverlap: 0 });
  assert.ok(chunks.length > 1);

  // The second window chunk should include the header row for better standalone retrieval quality.
  const second = chunks[1];
  assert.equal(second.text.split("\n")[0], "Region\tRevenue");
});

  const tallSheet = makeSheet(80);
  const tallChunks = chunkSheetByRegions(tallSheet, { splitByRowWindows: true });
  assert.ok(tallChunks.length > 1);

  // Chunk ids should be deterministic for a stable region windowing scheme.
  const tallChunksAgain = chunkSheetByRegions(tallSheet, { splitByRowWindows: true });
  assert.deepEqual(
    tallChunksAgain.map((c) => c.id),
    tallChunks.map((c) => c.id),
  );

  const index = new RagIndex();
  await index.indexSheet(tallSheet, { splitByRowWindows: true });
  const idsAfterTall = Array.from(index.store.items.keys());
  assert.ok(idsAfterTall.length > 1);

  const shortSheet = makeSheet(20);
  await index.indexSheet(shortSheet, { splitByRowWindows: true });
  const idsAfterShort = Array.from(index.store.items.keys());
  assert.ok(idsAfterShort.length < idsAfterTall.length);

  const removed = idsAfterTall.filter((id) => !idsAfterShort.includes(id));
  assert.ok(removed.length > 0);
  for (const id of removed) {
    assert.equal(index.store.items.has(id), false);
  }
});
