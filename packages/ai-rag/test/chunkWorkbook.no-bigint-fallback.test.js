import assert from "node:assert/strict";
import test from "node:test";

test("chunkWorkbook: works when BigInt is unavailable (string coord keys fallback)", async () => {
  const originalBigInt = globalThis.BigInt;
  try {
    // Simulate runtimes that don't have BigInt (older browsers / JS engines).
    // `chunkWorkbook` should still parse and fall back to string coord keys.
    // @ts-ignore - test override
    globalThis.BigInt = undefined;

    const url = new URL("../src/workbook/chunkWorkbook.js", import.meta.url);
    url.search = "noBigInt=1";
    const mod = await import(url.href);
    const chunkWorkbook = mod.chunkWorkbook;

    const row = 9_000_000_000;
    const cells = new Map();
    cells.set(`${row},0`, { value: "A" });
    cells.set(`${row},1`, { value: "B" });

    const workbook = {
      id: "wb-no-bigint",
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
  } finally {
    globalThis.BigInt = originalBigInt;
  }
});

test("chunkWorkbook: detectRegions connects across Number/string packing boundary when BigInt is unavailable", async () => {
  const originalBigInt = globalThis.BigInt;
  try {
    // @ts-ignore - test override
    globalThis.BigInt = undefined;

    const url = new URL("../src/workbook/chunkWorkbook.js", import.meta.url);
    url.search = "noBigIntBoundary=1";
    const mod = await import(url.href);
    const chunkWorkbook = mod.chunkWorkbook;

    const row = 0;
    const colNum = (1 << 20) - 1;
    const colString = 1 << 20;

    const cells = new Map();
    // Insert the string-side key first to ensure the traversal crosses representations.
    cells.set(`${row},${colString}`, { value: "B" });
    cells.set(`${row},${colNum}`, { value: "A" });

    const workbook = {
      id: "wb-no-bigint-boundary",
      sheets: [{ name: "Sheet1", cells }],
      tables: [],
      namedRanges: [],
    };

    const chunks = chunkWorkbook(workbook);
    const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
    assert.equal(dataRegions.length, 1);
    assert.deepEqual(dataRegions[0].rect, { r0: row, c0: colNum, r1: row, c1: colString });
    assert.equal(dataRegions[0].cells[0][0].v, "A");
    assert.equal(dataRegions[0].cells[0][1].v, "B");
  } finally {
    globalThis.BigInt = originalBigInt;
  }
});

test("chunkWorkbook: detectRegions connects across Number/string row packing boundary when BigInt is unavailable", async () => {
  const originalBigInt = globalThis.BigInt;
  try {
    // @ts-ignore - test override
    globalThis.BigInt = undefined;

    const url = new URL("../src/workbook/chunkWorkbook.js", import.meta.url);
    url.search = "noBigIntRowBoundary=1";
    const mod = await import(url.href);
    const chunkWorkbook = mod.chunkWorkbook;

    const packColFactor = 1 << 20;
    const rowNum = Math.floor(Number.MAX_SAFE_INTEGER / packColFactor);
    const rowString = rowNum + 1;
    const col = 0;

    const cells = new Map();
    // Insert the string-side key first to ensure traversal crosses representations.
    cells.set(`${rowString},${col}`, { value: "B" });
    cells.set(`${rowNum},${col}`, { value: "A" });

    const workbook = {
      id: "wb-no-bigint-row-boundary",
      sheets: [{ name: "Sheet1", cells }],
      tables: [],
      namedRanges: [],
    };

    const chunks = chunkWorkbook(workbook);
    const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
    assert.equal(dataRegions.length, 1);
    assert.deepEqual(dataRegions[0].rect, { r0: rowNum, c0: col, r1: rowString, c1: col });
    assert.equal(dataRegions[0].cells[0][0].v, "A");
    assert.equal(dataRegions[0].cells[1][0].v, "B");
  } finally {
    globalThis.BigInt = originalBigInt;
  }
});
