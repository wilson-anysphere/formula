import assert from "node:assert/strict";
import test from "node:test";

import { chunkWorkbook } from "../src/workbook/chunkWorkbook.js";

function makeWorkbook() {
  return {
    id: "wb1",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          // Table region A1:C3
          [{ v: "Region" }, { v: "Revenue" }, { v: "Cost" }],
          [{ v: "North" }, { v: 100 }, { v: 80 }],
          [{ v: "South" }, { v: 200 }, { v: 150 }],
          // Some empty rows
          [],
          [],
          // Free-form data region B6:C7
          [null, { v: "Metric" }, { v: "Value" }],
          [null, { v: "Gross Margin" }, { v: 0.25 }],
          // Formula region E1:E3
          [{}, {}, {}, {}, { f: "=SUM(B2:B3)" }],
          [{}, {}, {}, {}, { f: "=B2/C2" }],
          [{}, {}, {}, {}, { f: "=B3/C3" }],
        ],
      },
    ],
    tables: [
      {
        name: "SalesByRegion",
        sheetName: "Sheet1",
        rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
      },
    ],
    namedRanges: [
      {
        name: "SummaryMetrics",
        sheetName: "Sheet1",
        rect: { r0: 5, c0: 1, r1: 6, c1: 2 },
      },
    ],
  };
}

test("chunkWorkbook creates table + named range chunks with stable ids", () => {
  const workbook = makeWorkbook();
  const chunks = chunkWorkbook(workbook);

  const table = chunks.find((c) => c.kind === "table" && c.title === "SalesByRegion");
  const namedRange = chunks.find((c) => c.kind === "namedRange" && c.title === "SummaryMetrics");

  assert.ok(table, "expected table chunk");
  assert.ok(namedRange, "expected named range chunk");

  assert.equal(table.id, "wb1::Sheet1::table::SalesByRegion");
  assert.equal(namedRange.id, "wb1::Sheet1::namedRange::SummaryMetrics");
});

test("chunkWorkbook detects formula-heavy regions", () => {
  const workbook = makeWorkbook();
  const chunks = chunkWorkbook(workbook);

  const formulaRegions = chunks.filter((c) => c.kind === "formulaRegion");
  assert.ok(formulaRegions.length >= 1, "expected at least one formula region");
  assert.ok(formulaRegions[0].cells.some((row) => row.some((cell) => cell.f)), "expected formulas in chunk cells");
});

test("chunkWorkbook suppresses formulaRegion chunks that overlap a dataRegion", () => {
  const workbook = {
    id: "wb1",
    sheets: [
      {
        name: "Sheet1",
        // Mixed data + formulas in one connected non-empty block A1:C3.
        cells: [
          [{ v: "Item" }, { v: "Price" }, { v: "Taxed" }],
          [{ v: "A" }, { v: 10 }, { f: "=B2*1.1" }],
          [{ v: "B" }, { v: 20 }, { f: "=B3*1.1" }],
        ],
      },
    ],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  assert.equal(chunks.filter((c) => c.kind === "dataRegion").length, 1);
  assert.equal(chunks.filter((c) => c.kind === "formulaRegion").length, 0);
});

test("chunkWorkbook still produces standalone formulaRegion chunks", () => {
  const workbook = {
    id: "wb1",
    sheets: [
      {
        name: "Sheet1",
        cells: [[{ f: "=1+1" }], [{ f: "=2+2" }]],
      },
    ],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const formulaRegions = chunks.filter((c) => c.kind === "formulaRegion");
  assert.equal(formulaRegions.length, 1);
  assert.deepEqual(formulaRegions[0].rect, { r0: 0, c0: 0, r1: 1, c1: 0 });
});

test("chunkWorkbook respects AbortSignal", () => {
  const workbook = makeWorkbook();
  // Use a large rect to ensure we exercise abort handling even when extraction would
  // otherwise be expensive.
  workbook.tables[0].rect = { r0: 0, c0: 0, r1: 999, c1: 999 };
  const abortController = new AbortController();
  abortController.abort();

  assert.throws(() => chunkWorkbook(workbook, { signal: abortController.signal }), { name: "AbortError" });
});

test("chunkWorkbook propagates AbortSignal into extractCells (abort during extraction)", () => {
  const abortController = new AbortController();
  let calls = 0;

  const workbook = {
    id: "wb1",
    sheets: [
      {
        name: "Sheet1",
        // Force extractCells to use the `getCell` callback so we can trigger abort mid-extraction.
        getCell() {
          calls += 1;
          if (calls === 10) abortController.abort();
          return null;
        },
      },
    ],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 5000 } }],
    namedRanges: [],
  };

  assert.throws(
    () =>
      chunkWorkbook(workbook, {
        signal: abortController.signal,
        extractMaxRows: 1,
        extractMaxCols: 5000,
      }),
    { name: "AbortError" }
  );
});

test("chunkWorkbook caps the number of disconnected data regions per sheet", () => {
  const map = new Map();
  // 20 disconnected 2-cell regions (each region is horizontal, with an empty row between).
  for (let i = 0; i < 20; i += 1) {
    const r = i * 2;
    map.set(`${r},0`, { v: `R${i}` });
    map.set(`${r},1`, { v: `R${i}` });
  }

  const workbook = {
    id: "wb1",
    sheets: [{ name: "Sheet1", cells: map }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook, { maxDataRegionsPerSheet: 10 });
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.equal(dataRegions.length, 10);
  assert.deepEqual(
    dataRegions.map((c) => c.rect),
    Array.from({ length: 10 }, (_, i) => ({ r0: i * 2, c0: 0, r1: i * 2, c1: 1 }))
  );
});

test("chunkWorkbook: truncation fallback still produces a deterministic chunk", () => {
  // Create >cellLimit matching cells, but ensure they're *not* connected (so without the
  // truncation fallback we would drop them as trivial 1-cell components).
  const cells = [];
  cells[0] = [];
  const cellLimit = 25;
  for (let i = 0; i < cellLimit + 5; i += 1) {
    // Space out cells so they aren't 4-neighbor connected.
    const r = i * 2;
    cells[r] = [];
    cells[r][0] = { v: `V${i}` };
  }

  const workbook = {
    id: "wb1",
    sheets: [{ name: "Sheet1", cells }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook, { detectRegionsCellLimit: cellLimit });
  const dataRegions = chunks.filter((c) => c.kind === "dataRegion");
  assert.ok(dataRegions.length >= 1, "expected at least one dataRegion chunk");
  assert.ok(
    dataRegions.some((c) => c.meta?.truncated === true),
    "expected a truncation fallback chunk"
  );
});

test("chunkWorkbook encodes id parts to avoid delimiter-in-name collisions", () => {
  const workbook = {
    id: "wb1",
    sheets: [
      { name: "A", cells: [[{ v: "x" }]] },
      { name: "A::table::B", cells: [[{ v: "y" }]] },
    ],
    tables: [
      // Old-style ids would collide:
      // wb1::A::table::B::table::C  (sheet=A, table=B::table::C)
      // wb1::A::table::B::table::C  (sheet=A::table::B, table=C)
      { name: "B::table::C", sheetName: "A", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
      { name: "C", sheetName: "A::table::B", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
    ],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  const tableChunks = chunks.filter((c) => c.kind === "table");
  assert.equal(tableChunks.length, 2, "expected two table chunks");

  const ids = new Set(tableChunks.map((c) => c.id));
  assert.equal(ids.size, 2, "expected distinct ids for colliding names");
});

test("chunkWorkbook: does not call toString on non-string sparse cell-map keys", () => {
  let calls = 0;
  const badKey = {
    toString() {
      calls += 1;
      return "0,0";
    },
  };

  const map = new Map();
  map.set(badKey, { v: "SHOULD_NOT_BE_READ" });
  map.set("0,0", { v: "A" });
  map.set("0,1", { v: "B" });

  const workbook = {
    id: "wb1",
    sheets: [{ name: "Sheet1", cells: map }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  assert.equal(calls, 0);
  assert.ok(chunks.some((c) => c.kind === "dataRegion"), "expected a dataRegion chunk");
});

test("chunkWorkbook: does not call toString on non-string formula values", () => {
  const secret = "TopSecretFormula";
  let calls = 0;
  const formula = {
    toString() {
      calls += 1;
      return secret;
    },
  };

  const workbook = {
    id: "wb1",
    sheets: [{ name: "Sheet1", cells: [[{ f: formula }], [{ f: formula }]] }],
    tables: [],
    namedRanges: [],
  };

  const chunks = chunkWorkbook(workbook);
  assert.equal(calls, 0);
  assert.ok(chunks.some((c) => c.kind === "formulaRegion"), "expected a formulaRegion chunk");
});
