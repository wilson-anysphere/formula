import { describe, expect, it } from "vitest";

import { extractSheetSchema } from "./schema.js";
import { summarizeRegion, summarizeSheetSchema } from "./summarizeSheet.js";

describe("summarizeSheetSchema", () => {
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

  it("produces a stable, compact output string", () => {
    const schema = buildSchema();
    const summary = summarizeSheetSchema(schema);

    expect(summary).toBe(
      [
        "sheet=[Sheet1] tables=2 regions=2 named=0",
        "T1 [SalesTable] r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
        "T2 [Matrix] r=[Sheet1!A5:B6] rows=2 cols=2 hdr=0 h=[Column1|Column2] t=[number|number]",
        "R1 r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
        "R2 r=[Sheet1!A5:B6] rows=2 cols=2 hdr=0 h=[Column1|Column2] t=[number|number]",
      ].join("\n"),
    );
  });

  it("includes key facts (names, headers, inferred types)", () => {
    const schema = buildSchema();
    const summary = summarizeSheetSchema(schema);

    expect(summary).toContain("[SalesTable]");
    expect(summary).toContain("r=[Sheet1!A1:C3]");
    expect(summary).toContain("h=[Product|Sales|Active]");
    expect(summary).toContain("t=[string|number|boolean]");
  });

  it("respects verbosity caps (tables + headers)", () => {
    const schema = buildSchema();
    const summary = summarizeSheetSchema(schema, { maxTables: 1, maxHeadersPerTable: 2, maxTypesPerTable: 2, maxRegions: 1 });

    // Table cap
    expect(summary).toContain("tables=2");
    expect(summary).toContain("T1 [SalesTable]");
    expect(summary).not.toContain("[Matrix]");
    expect(summary).toContain("T…+1");

    // Header/type cap
    expect(summary).toContain("h=[Product|Sales|…+1]");
    expect(summary).toContain("t=[string|number|…+1]");

    // Region cap
    expect(summary).toContain("R1 r=[Sheet1!A1:C3]");
    expect(summary).toContain("R…+1");
  });

  it("escapes delimiter characters so bracketed lists remain parseable", () => {
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
    expect(summary).toContain("[T\\]1\\|2]");
    // Headers should be escaped inside the bracketed list.
    expect(summary).toContain("h=[A\\|B|C\\]D|E\\\\F]");
  });

  it("summarizeRegion supports both tables and data regions", () => {
    const schema = buildSchema();
    expect(summarizeRegion(schema.tables[0]!)).toBe(
      "T [SalesTable] r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
    );
    expect(summarizeRegion(schema.dataRegions[0]!)).toBe(
      "R r=[Sheet1!A1:C3] rows=2 cols=3 hdr=1 h=[Product|Sales|Active] t=[string|number|boolean]",
    );
  });

  it("can exclude tables or regions from the sheet summary", () => {
    const schema = buildSchema();
    const noTables = summarizeSheetSchema(schema, { includeTables: false, maxRegions: 1 });
    expect(noTables).toContain("sheet=[Sheet1] tables=2 regions=2 named=0");
    expect(noTables).not.toContain("\nT1 ");
    expect(noTables).toContain("\nR1 ");

    const noRegions = summarizeSheetSchema(schema, { includeRegions: false, maxTables: 1 });
    expect(noRegions).toContain("sheet=[Sheet1] tables=2 regions=2 named=0");
    expect(noRegions).toContain("\nT1 ");
    expect(noRegions).not.toContain("\nR1 ");
  });

  it("sorts tables/regions deterministically by range regardless of input order", () => {
    const schema = extractSheetSchema({
      name: "Sheet1",
      values: [
        ["H1", "H2", null, "X"],
        [1, 2, null, 9],
      ],
    });

    // Create an intentionally unsorted schema view (reverse the lists).
    const reversed: any = {
      ...schema,
      tables: schema.tables.slice().reverse(),
      dataRegions: schema.dataRegions.slice().reverse(),
    };

    const summary = summarizeSheetSchema(reversed);

    // The left-most range (A1:B2) should always appear first.
    expect(summary).toContain("T1 [Region1] r=[Sheet1!A1:B2]");
    expect(summary).toContain("T2 [Region2] r=[Sheet1!D1:D2]");
    expect(summary).toContain("R1 r=[Sheet1!A1:B2]");
    expect(summary).toContain("R2 r=[Sheet1!D1:D2]");
  });

  it("supports maxTables=0 (only emits a truncation marker)", () => {
    const schema = buildSchema();
    const summary = summarizeSheetSchema(schema, { maxTables: 0, includeRegions: false });
    expect(summary).toContain("tables=2");
    expect(summary).not.toContain("\nT1 ");
    expect(summary).toContain("T…+2");
  });

  it("supports maxRegions=0 (only emits a truncation marker)", () => {
    const schema = buildSchema();
    const summary = summarizeSheetSchema(schema, { maxRegions: 0, includeTables: false });
    expect(summary).toContain("regions=2");
    expect(summary).not.toContain("\nR1 ");
    expect(summary).toContain("R…+2");
  });

  it("includes named ranges (and respects maxNamedRanges)", () => {
    const schema = extractSheetSchema({
      name: "Sheet1",
      values: [["A", "B"]],
      namedRanges: [
        { name: "NR1", range: "Sheet1!A1" },
        { name: "NR2", range: "Sheet1!B1" },
      ],
    });

    const summary = summarizeSheetSchema(schema, { maxTables: 0, maxRegions: 0, maxNamedRanges: 1 });
    expect(summary).toContain("named=2");
    expect(summary).toContain("N1 [NR1] r=[Sheet1!A1]");
    expect(summary).not.toContain("N2 [NR2]");
    expect(summary).toContain("N…+1");
  });
});
