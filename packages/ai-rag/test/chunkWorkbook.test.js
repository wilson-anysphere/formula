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

test("chunkWorkbook respects AbortSignal", () => {
  const workbook = makeWorkbook();
  const abortController = new AbortController();
  abortController.abort();

  assert.throws(() => chunkWorkbook(workbook, { signal: abortController.signal }), { name: "AbortError" });
});
