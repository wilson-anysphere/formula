import assert from "node:assert/strict";
import test from "node:test";

import { chunkToText } from "../src/workbook/chunkToText.js";

test("chunkToText renders labeled sample rows when a header row is detected", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
    cells: [
      [{ v: "Region" }, { v: "Revenue" }, { v: "Units" }],
      [{ v: "North" }, { v: 1200 }, { v: 10 }],
      [{ v: "South" }, { v: 800 }, { v: 5 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Region=North/);
  assert.match(text, /Revenue=1200/);
  assert.match(text, /Units=10/);
});

test("chunkToText includes formulas in labeled sample rows for header tables", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Region" }, { v: "Revenue" }],
      [{ v: "North" }, { f: "=B2*2", v: 200 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Revenue\(=B2\*2\)=200/);
});

test("chunkToText falls back to Column<N> when a header cell is empty", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "" }, { v: "Name" }],
      [{ v: 123 }, { v: "Alice" }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Column1=123/);
  assert.match(text, /Name=Alice/);
});

test("chunkToText disambiguates duplicate header names in labeled rows", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 2 },
    cells: [
      [{ v: "Value" }, { v: "Value" }, { v: "Value" }],
      [{ v: 1 }, { v: 2 }, { v: 3 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Value=1/);
  assert.match(text, /Value_2=2/);
  assert.match(text, /Value_3=3/);
});

test("chunkToText caps wide tables with an explicit truncation indicator", () => {
  const colCount = 25;
  const headers = Array.from({ length: colCount }, (_, i) => ({ v: `H${i + 1}` }));
  const row = Array.from({ length: colCount }, (_, i) => ({ v: `V${i + 1}` }));

  const chunk = {
    kind: "table",
    title: "Wide",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: colCount - 1 },
    cells: [headers, row],
  };

  const text = chunkToText(chunk, { sampleRows: 1, maxColumnsForSchema: 5, maxColumnsForRows: 5 });
  assert.ok(text.includes("… (+20 more columns)"), "expected a column truncation indicator");
  assert.ok(!text.includes("H25"), "should not list all column headers");
  assert.ok(!text.includes("V25"), "should not list all row values");
});

test("chunkToText uses the full range width when reporting truncated column counts", () => {
  const sampledCols = 50;
  const fullCols = 100;
  const headers = Array.from({ length: sampledCols }, (_, i) => ({ v: `H${i + 1}` }));
  const row = Array.from({ length: sampledCols }, (_, i) => ({ v: `V${i + 1}` }));

  const chunk = {
    kind: "table",
    title: "Wide (truncated sample)",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: fullCols - 1 },
    cells: [headers, row],
  };

  const text = chunkToText(chunk, { sampleRows: 1, maxColumnsForSchema: 5, maxColumnsForRows: 5 });
  assert.ok(text.includes("… (+95 more columns)"), "expected truncation to reflect full range width");
});

test("chunkToText includes A1-like cell addresses for formulaRegion samples", () => {
  const chunk = {
    kind: "formulaRegion",
    title: "Formula region E1:E2",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 4, r1: 1, c1: 4 },
    cells: [[{ f: "=SUM(B2:B3)" }], [{ f: "=B2/C2" }]],
  };

  const text = chunkToText(chunk);
  assert.match(text, /E1:=SUM\(B2:B3\)/);
  assert.match(text, /E2:=B2\/C2/);
});

test("chunkToText truncates long formulas inside formulaRegion samples", () => {
  const longFormula = `=${"A".repeat(80)}`;
  const chunk = {
    kind: "formulaRegion",
    title: "Formula region A1",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 0, c1: 0 },
    cells: [[{ f: longFormula }]],
  };

  const text = chunkToText(chunk);
  assert.match(text, /A1:=A{56}\.\.\./);
  assert.doesNotMatch(text, /A{80}/);
});
