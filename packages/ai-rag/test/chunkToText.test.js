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

test("chunkToText sanitizes '=' in header labels to keep key/value rows parseable", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "A=B" }, { v: "Value" }],
      [{ v: 1 }, { v: 2 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /COLUMNS: A≡B/);
  assert.match(text, /A≡B=1/);
  assert.doesNotMatch(text, /A=B=1/);
});

test("chunkToText escapes '|' characters in cell text so row separators remain unambiguous", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Region" }, { v: "Revenue" }],
      [{ v: "North|East" }, { v: 1200 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Region=North¦East/);
  assert.doesNotMatch(text, /North\|East/);
});

test("chunkToText formats Date values as ISO strings", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Date" }, { v: "Value" }],
      [{ v: new Date("2024-01-02T03:04:05.000Z") }, { v: 1 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Date=2024-01-02T03:04:05\.000Z/);
});

test("chunkToText does not call per-instance Date#toISOString overrides", () => {
  const secret = "TopSecretDate";
  const date = new Date("2024-01-02T03:04:05.000Z");
  Object.defineProperty(date, "toISOString", {
    value: () => secret,
    enumerable: false,
  });
  Object.defineProperty(date, "getTime", {
    value: () => {
      throw new Error("getTime override should not be called");
    },
    enumerable: false,
  });

  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Date" }], [{ v: date }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Date=2024-01-02T03:04:05\.000Z/);
  assert.doesNotMatch(text, new RegExp(secret));
});

test("chunkToText truncates BigInt values via the string formatting path", () => {
  const big = BigInt("1".repeat(100));
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Big" }], [{ v: big }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Big=1{57}\.\.\./);
  assert.doesNotMatch(text, /1{80}/);
});

test("chunkToText formats object cell values via JSON for stable output", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Meta" }], [{ v: { foo: "bar" } }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Meta=\{"foo":"bar"\}/);
});

test("chunkToText JSON-stringifies objects containing BigInt properties", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Meta" }], [{ v: { big: 1n } }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Meta=\{"big":"1"\}/);
});

test("chunkToText does not call toJSON on object cell values", () => {
  const secret = "TopSecretToJson";
  const value = {
    toJSON() {
      return secret;
    },
  };

  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Meta" }], [{ v: value }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.doesNotMatch(text, new RegExp(secret));
});

test("chunkToText does not call custom toString on chunk sheetName/title", () => {
  const secret = "TopSecretSheet";
  let calls = 0;
  const sheetNameObj = {
    toString() {
      calls += 1;
      return secret;
    },
  };
  const titleObj = {
    toString() {
      calls += 1;
      return secret;
    },
  };

  const chunk = {
    kind: "table",
    title: titleObj,
    sheetName: sheetNameObj,
    rect: { r0: 0, c0: 0, r1: 0, c1: 0 },
    cells: [[{ v: "A" }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.equal(calls, 0);
  assert.doesNotMatch(text, new RegExp(secret));
});

test("chunkToText formats Error values with name + message", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Result" }], [{ v: Object.assign(new Error("boom | bang"), { name: "Div0" }) }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Result=Div0: boom ¦ bang/);
  assert.doesNotMatch(text, /\|/);
});

test("chunkToText prefers object.text when present", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells: [[{ v: "Notes" }], [{ v: { text: "hello", runs: [{ text: "hello", bold: true }] } }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Notes=hello/);
  assert.doesNotMatch(text, /runs/);
});

test("chunkToText formats in-cell image values as alt text / placeholders", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Photo" }, { v: "Other" }],
      [
        { v: { type: "image", value: { imageId: "img_1", altText: "Kitten" } } },
        { v: { type: "image", value: { imageId: "img_2" } } },
      ],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Photo=Kitten/);
  assert.match(text, /Other=\[Image\]/);
  // Avoid leaking internal image payload structure into RAG text.
  assert.doesNotMatch(text, /imageId|img_1|img_2|\"type\"/);
});

test("chunkToText detects header rows when header cells contain rich values", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: { text: "Product", runs: [{ start: 0, end: 7, style: { bold: true } }] } }, { v: { type: "image", value: { imageId: "img_1", altText: "Photo" } } }],
      [{ v: "Alpha" }, { v: { type: "image", value: { imageId: "img_2" } } }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Product=Alpha/);
  assert.match(text, /Photo=\[Image\]/);
});

test("chunkToText detects header rows below a title row and preserves the title as pre-header context", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Data region A1:C3",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
    cells: [
      [{ v: "Revenue Summary" }, {}, {}],
      [{ v: "Region" }, { v: "Revenue" }, { v: "Units" }],
      [{ v: "North" }, { v: 1200 }, { v: 10 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /PRE-HEADER ROWS:/);
  assert.match(text, /Revenue Summary/);
  assert.match(text, /Region=North/);
  assert.match(text, /Revenue=1200/);
  assert.match(text, /Units=10/);
});

test("chunkToText indicates when there are additional pre-header rows", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Data region A1:B6",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 5, c1: 1 },
    cells: [
      [{ v: "Revenue Summary" }, {}],
      [{ v: "(as of 2024)" }, {}],
      [],
      [],
      [{ v: "Region" }, { v: "Revenue" }],
      [{ v: "North" }, { v: 1200 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /PRE-HEADER ROWS:/);
  assert.match(text, /… \(\+2 more pre-header rows\)/);
  assert.match(text, /Region=North/);
});

test("chunkToText surfaces non-empty pre-header rows even if earlier pre-header rows are empty", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Data region A1:B5",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 4, c1: 1 },
    cells: [
      [],
      [],
      [{ v: "Summary" }, {}],
      [{ v: "Region" }, { v: "Revenue" }],
      [{ v: "North" }, { v: 1200 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /PRE-HEADER ROWS:/);
  assert.match(text, /\bSummary\b/);
  assert.match(text, /Region=North/);
});

test("chunkToText treats a sparse header row with blank columns as a header (not as data)", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Name" }, { v: "" }],
      [{ v: "Alice" }, { v: "Seattle" }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Name=Alice/);
  assert.match(text, /Column2=Seattle/);
});

test("chunkToText does not misclassify multi-word headers as title rows", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Customer Name" }, { v: "" }],
      [{ v: "Alice" }, { v: "Seattle" }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Customer Name=Alice/);
  assert.match(text, /Column2=Seattle/);
  assert.doesNotMatch(text, /PRE-HEADER ROWS:/);
});

test("chunkToText treats single-word caption rows (e.g. 'Summary') as title rows when followed by a real header", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Data region A1:C3",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
    cells: [
      [{ v: "Summary" }, {}, {}],
      [{ v: "Region" }, { v: "Revenue" }, { v: "Units" }],
      [{ v: "North" }, { v: 1200 }, { v: 10 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /PRE-HEADER ROWS:/);
  assert.match(text, /Summary/);
  assert.match(text, /Region=North/);
});

test("chunkToText includes column truncation indicator in PRE-HEADER ROWS when table is wide", () => {
  const colCount = 25;
  const titleRow = [{ v: "Revenue Summary" }, ...Array.from({ length: colCount - 1 }, () => ({}))];
  const headerRow = Array.from({ length: colCount }, (_, i) => ({ v: `H${i + 1}` }));
  const dataRow = Array.from({ length: colCount }, (_, i) => ({ v: `V${i + 1}` }));

  const chunk = {
    kind: "dataRegion",
    title: "Wide region",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 2, c1: colCount - 1 },
    cells: [titleRow, headerRow, dataRow],
  };

  const text = chunkToText(chunk, { sampleRows: 1, maxColumnsForRows: 5, maxColumnsForSchema: 5 });
  assert.match(text, /PRE-HEADER ROWS:/);
  assert.match(text, /… \(\+20 more columns\)/);
});

test("chunkToText uses the widest sampled row when computing column counts (jagged samples)", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Data region A1:C3",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
    cells: [
      [{ v: "Revenue Summary" }],
      [{ v: "Region" }, { v: "Revenue" }, { v: "Units" }],
      [{ v: "North" }, { v: 1200 }, { v: 10 }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /COLUMNS: Region/);
  assert.match(text, /Revenue/);
  assert.match(text, /Units/);
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

test("chunkToText omits trailing '=' when a formula cell has no computed value", () => {
  const chunk = {
    kind: "table",
    title: "Example",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 1 },
    cells: [
      [{ v: "Region" }, { v: "Revenue" }],
      [{ v: "North" }, { f: "=B2*2", v: null }],
    ],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /Revenue\(=B2\*2\)/);
  assert.doesNotMatch(text, /Revenue\(=B2\*2\)=/);
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

test("chunkToText reports when sample rows are truncated relative to the full range height", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Tall region",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 9, c1: 0 }, // 10 rows
    cells: [[{ v: 1 }]], // sampled 1 row
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /… \(\+9 more rows\)/);
});

test("chunkToText tolerates sparse cell arrays with missing rows", () => {
  const cells = new Array(2);
  cells[0] = [{ v: "Header" }];
  // Leave cells[1] unset (sparse array / missing row).

  const chunk = {
    kind: "dataRegion",
    title: "Sparse",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 1, c1: 0 },
    cells,
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /SAMPLE ROWS:/);
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

test("chunkToText includes computed values for formulaRegion entries when available", () => {
  const chunk = {
    kind: "formulaRegion",
    title: "Formula region E1",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 4, r1: 0, c1: 4 },
    cells: [[{ f: "=SUM(B2:B3)", v: 300 }]],
  };

  const text = chunkToText(chunk);
  assert.match(text, /E1:=SUM\(B2:B3\)=300/);
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

test("chunkToText reports when formulaRegion samples are truncated", () => {
  const cells = Array.from({ length: 13 }, (_, r) => [{ f: `=A${r + 1}` }]);
  const chunk = {
    kind: "formulaRegion",
    title: "Formula region A1:A13",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 12, c1: 0 },
    cells,
  };

  const text = chunkToText(chunk);
  assert.match(text, /… \(\+1 more formulas\)/);
  assert.doesNotMatch(text, /\bA13:=/);
});

test("chunkToText includes computed values for non-header formula cells when available", () => {
  const chunk = {
    kind: "dataRegion",
    title: "Test",
    sheetName: "Sheet1",
    rect: { r0: 0, c0: 0, r1: 0, c1: 0 },
    cells: [[{ f: "=A1*2", v: 2 }]],
  };

  const text = chunkToText(chunk, { sampleRows: 1 });
  assert.match(text, /=A1\*2=2/);
  assert.doesNotMatch(text, /==A1\*2/);
});
