import assert from "node:assert/strict";
import test from "node:test";

import { chunkWorkbook } from "../src/workbook/chunkWorkbook.js";

test("chunkWorkbook: packed coord keys handle large row indices for Map-backed sheets", () => {
  const row = 999_999;
  const cells = new Map();
  cells.set(`${row},0`, { value: "A" });
  cells.set(`${row},1`, { value: "B" });

  const workbook = {
    id: "wb-packed-keys",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 1);
  assert.deepEqual(dataRegions[0].rect, { r0: row, c0: 0, r1: row, c1: 1 });
  assert.equal(dataRegions[0].cells[0][0].v, "A");
  assert.equal(dataRegions[0].cells[0][1].v, "B");
});

test("chunkWorkbook: coord packing works for very large rows (BigInt/string fallback)", () => {
  // Exceeds the Number packing limit (row * 2^20 must stay within MAX_SAFE_INTEGER).
  const row = 9_000_000_000;
  const cells = new Map();
  cells.set(`${row},0`, { value: "A" });
  cells.set(`${row},1`, { value: "B" });

  const workbook = {
    id: "wb-packed-keys-big",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 1);
  assert.deepEqual(dataRegions[0].rect, { r0: row, c0: 0, r1: row, c1: 1 });
  assert.equal(dataRegions[0].cells[0][0].v, "A");
  assert.equal(dataRegions[0].cells[0][1].v, "B");
});

test("chunkWorkbook: detectRegions connects across Number/BigInt coord key boundaries", () => {
  // This crosses the col packing boundary where `packCoordKey` switches from Number
  // keys (col < 2^20) to BigInt keys (col >= 2^20).
  const row = 0;
  const colBig = 1_048_576; // 2^20
  const colNum = colBig - 1;

  const cells = new Map();
  // Insert the BigInt-represented coordinate first to ensure traversal doesn't rely on
  // insertion order to connect the region.
  cells.set(`${row},${colBig}`, { value: "B" });
  cells.set(`${row},${colNum}`, { value: "A" });

  const workbook = {
    id: "wb-packed-keys-boundary",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 1);
  assert.deepEqual(dataRegions[0].rect, { r0: row, c0: colNum, r1: row, c1: colBig });
  assert.equal(dataRegions[0].cells[0][0].v, "A");
  assert.equal(dataRegions[0].cells[0][1].v, "B");
});

test("chunkWorkbook: detectRegions connects across BigInt/string coord key boundaries", () => {
  const row = 0;
  const maxUint32 = 2 ** 32 - 1;
  const colBigInt = maxUint32;
  const colString = maxUint32 + 1;

  const cells = new Map();
  // Insert the string-key fallback coordinate first to ensure traversal crosses
  // the packing boundary in both directions.
  cells.set(`${row},${colString}`, { value: "B" });
  cells.set(`${row},${colBigInt}`, { value: "A" });

  const workbook = {
    id: "wb-packed-keys-u32-boundary",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 1);
  assert.deepEqual(dataRegions[0].rect, { r0: row, c0: colBigInt, r1: row, c1: colString });
  assert.equal(dataRegions[0].cells[0][0].v, "A");
  assert.equal(dataRegions[0].cells[0][1].v, "B");
});

test("chunkWorkbook: detectRegions connects across Number/BigInt row packing boundary", () => {
  // This crosses the row packing boundary where `packCoordKey` switches from Number
  // keys to BigInt keys (row*2^20 must stay within MAX_SAFE_INTEGER).
  const rowNum = 8_589_934; // MAX_SAFE_PACKED_ROW
  const rowBig = rowNum + 1;
  const col = 0;

  const cells = new Map();
  // Insert the BigInt-represented coordinate first to avoid relying on insertion order.
  cells.set(`${rowBig},${col}`, { value: "B" });
  cells.set(`${rowNum},${col}`, { value: "A" });

  const workbook = {
    id: "wb-packed-keys-row-boundary",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 1);
  assert.deepEqual(dataRegions[0].rect, { r0: rowNum, c0: col, r1: rowBig, c1: col });
  assert.equal(dataRegions[0].cells[0][0].v, "A");
  assert.equal(dataRegions[0].cells[1][0].v, "B");
});
