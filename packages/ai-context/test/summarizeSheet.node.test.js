import assert from "node:assert/strict";
import test from "node:test";

import { extractSheetSchema } from "../src/schema.js";
import { summarizeRegion, summarizeSheetSchema } from "../src/summarizeSheet.js";

function buildSchema() {
  return extractSheetSchema({
    name: "Sheet1",
    values: [
      ["Product", "Sales", "Active"],
      ["Alpha", 10, true],
      ["Beta", 20, false],
      [null, null, null],
      [1, 2],
      [3, 4],
    ],
    tables: [
      { name: "SalesTable", range: "A1:C3" },
      { name: "Matrix", range: "A5:B6" },
    ],
  });
}

test("summarizeSheetSchema: produces a stable, compact output string", () => {
  const schema = buildSchema();
  const summary = summarizeSheetSchema(schema);

  assert.equal(
    summary,
    [
      "sheet=[Sheet1] tables=2 regions=2 named=0",
      "T1 [SalesTable] r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
      "T2 [Matrix] r=[Sheet1!A5:B6] rows=2 cols=2 hdr=0 h=[Column1|Column2] t=[number|number]",
      "R1 r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
      "R2 r=[Sheet1!A5:B6] rows=2 cols=2 hdr=0 h=[Column1|Column2] t=[number|number]",
    ].join("\n"),
  );
});

test("summarizeSheetSchema: includes key facts (names, headers, inferred types)", () => {
  const schema = buildSchema();
  const summary = summarizeSheetSchema(schema);

  assert.ok(summary.includes("[SalesTable]"));
  assert.ok(summary.includes("r=[Sheet1!A1:C3]"));
  assert.ok(summary.includes("h=[Product|Sales|Active]"));
  assert.ok(summary.includes("t=[string|number|boolean]"));
});

test("summarizeSheetSchema: respects verbosity caps (tables + headers)", () => {
  const schema = buildSchema();
  const summary = summarizeSheetSchema(schema, { maxTables: 1, maxHeadersPerTable: 2, maxTypesPerTable: 2, maxRegions: 1 });

  // Table cap
  assert.ok(summary.includes("tables=2"));
  assert.ok(summary.includes("T1 [SalesTable]"));
  assert.ok(!summary.includes("[Matrix]"));
  assert.ok(summary.includes("T…+1"));

  // Header/type cap
  assert.ok(summary.includes("h=[Product|Sales|…+1]"));
  assert.ok(summary.includes("t=[string|number|…+1]"));

  // Region cap
  assert.ok(summary.includes("R1 r=[Sheet1!A1:C3]"));
  assert.ok(summary.includes("R…+1"));
});

test("summarizeSheetSchema: escapes delimiter characters so bracketed lists remain parseable", () => {
  const schema = extractSheetSchema({
    name: "Sheet1",
    values: [
      ["A|B", "C]D", "E\\F"],
      [1, 2, 3],
    ],
    tables: [{ name: "T]1|2", range: "A1:C2" }],
  });

  const summary = summarizeSheetSchema(schema);
  // Table name should escape `]` and `|`.
  assert.ok(summary.includes("[T\\]1\\|2]"));
  // Headers should be escaped inside the bracketed list.
  assert.ok(summary.includes("h=[A\\|B|C\\]D|E\\\\F]"));
});

test("summarizeRegion: supports both tables and data regions", () => {
  const schema = buildSchema();
  assert.equal(
    summarizeRegion(schema.tables[0]),
    "T [SalesTable] r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
  );
  assert.equal(
    summarizeRegion(schema.dataRegions[0]),
    "R r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
  );
});

test("summarizeSheetSchema: can exclude tables or regions from the sheet summary", () => {
  const schema = buildSchema();
  const noTables = summarizeSheetSchema(schema, { includeTables: false, maxRegions: 1 });
  assert.ok(noTables.includes("sheet=[Sheet1] tables=2 regions=2 named=0"));
  assert.ok(!noTables.includes("\nT1 "));
  assert.ok(noTables.includes("\nR1 "));

  const noRegions = summarizeSheetSchema(schema, { includeRegions: false, maxTables: 1 });
  assert.ok(noRegions.includes("sheet=[Sheet1] tables=2 regions=2 named=0"));
  assert.ok(noRegions.includes("\nT1 "));
  assert.ok(!noRegions.includes("\nR1 "));
});

test("summarizeSheetSchema: sorts tables/regions deterministically by range regardless of input order", () => {
  const schema = extractSheetSchema({
    name: "Sheet1",
    values: [
      ["H1", "H2", null, "X"],
      [1, 2, null, 9],
    ],
  });

  // Create an intentionally unsorted schema view (reverse the lists).
  const reversed = {
    ...schema,
    tables: schema.tables.slice().reverse(),
    dataRegions: schema.dataRegions.slice().reverse(),
  };

  const summary = summarizeSheetSchema(reversed);

  // The left-most range (A1:B2) should always appear first.
  assert.ok(summary.includes("T1 [Region1] r=[Sheet1!A1:B2]"));
  assert.ok(summary.includes("T2 [Region2] r=[Sheet1!D1:D2]"));
  assert.ok(summary.includes("R1 r=[Sheet1!A1:B2]"));
  assert.ok(summary.includes("R2 r=[Sheet1!D1:D2]"));
});

test("summarizeSheetSchema: supports maxTables=0 (only emits a truncation marker)", () => {
  const schema = buildSchema();
  const summary = summarizeSheetSchema(schema, { maxTables: 0, includeRegions: false });
  assert.ok(summary.includes("tables=2"));
  assert.ok(!summary.includes("\nT1 "));
  assert.ok(summary.includes("T…+2"));
});

test("summarizeSheetSchema: supports maxRegions=0 (only emits a truncation marker)", () => {
  const schema = buildSchema();
  const summary = summarizeSheetSchema(schema, { maxRegions: 0, includeTables: false });
  assert.ok(summary.includes("regions=2"));
  assert.ok(!summary.includes("\nR1 "));
  assert.ok(summary.includes("R…+2"));
});

test("summarizeSheetSchema: includes named ranges (and respects maxNamedRanges)", () => {
  const schema = extractSheetSchema({
    name: "Sheet1",
    values: [["A", "B"]],
    namedRanges: [
      { name: "NR1", range: "Sheet1!A1" },
      { name: "NR2", range: "Sheet1!B1" },
    ],
  });

  const summary = summarizeSheetSchema(schema, { maxTables: 0, maxRegions: 0, maxNamedRanges: 1 });
  assert.ok(summary.includes("named=2"));
  assert.ok(summary.includes("N1 [NR1] r=[Sheet1!A1]"));
  assert.ok(!summary.includes("N2 [NR2]"));
  assert.ok(summary.includes("N…+1"));
});

test("summarizeSheetSchema: can exclude named ranges from the sheet summary", () => {
  const schema = extractSheetSchema({
    name: "Sheet1",
    values: [["A", "B"]],
    namedRanges: [{ name: "NR", range: "Sheet1!A1" }],
  });

  const summary = summarizeSheetSchema(schema, { includeNamedRanges: false });
  assert.ok(summary.includes("named=1"));
  assert.ok(!summary.includes("\nN1 "));
});

test("summarizeSheetSchema: supports maxNamedRanges=0 (only emits a truncation marker)", () => {
  const schema = extractSheetSchema({
    name: "Sheet1",
    values: [["A", "B"]],
    namedRanges: [
      { name: "NR1", range: "Sheet1!A1" },
      { name: "NR2", range: "Sheet1!B1" },
    ],
  });

  const summary = summarizeSheetSchema(schema, { maxTables: 0, maxRegions: 0, maxNamedRanges: 0 });
  assert.ok(summary.includes("named=2"));
  assert.ok(!summary.includes("\nN1 "));
  assert.ok(summary.includes("N…+2"));
});

