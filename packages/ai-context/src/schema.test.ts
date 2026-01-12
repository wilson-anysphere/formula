import { describe, expect, it } from "vitest";

import { detectDataRegions, extractSheetSchema } from "./schema.js";

describe("extractSheetSchema", () => {
  it("detects a headered table region and infers column types", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        ["Product", "Sales", "Active"],
        ["Alpha", 10, true],
        ["Beta", 20, false]
      ],
      namedRanges: [{ name: "SalesData", range: "Sheet1!A1:C3" }]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.name).toBe("Sheet1");
    expect(schema.tables).toHaveLength(1);

    const table = schema.tables[0];
    expect(table.range).toBe("Sheet1!A1:C3");
    expect(table.rowCount).toBe(2);
    expect(table.columns.map((c: any) => ({ name: c.name, type: c.type }))).toEqual([
      { name: "Product", type: "string" },
      { name: "Sales", type: "number" },
      { name: "Active", type: "boolean" }
    ]);

    expect(schema.namedRanges).toEqual([{ name: "SalesData", range: "Sheet1!A1:C3" }]);
    expect(schema.dataRegions).toHaveLength(1);
    expect(schema.dataRegions[0].hasHeader).toBe(true);
  });

  it("detects multiple disconnected regions", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        ["A", null, null, "X"],
        [1, null, null, 9],
        [null, null, null, null]
      ]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.tables).toHaveLength(2);
    expect(schema.tables.map((t: any) => t.range)).toEqual(["Sheet1!A1:A2", "Sheet1!D1:D2"]);
  });

  it("does not treat numeric-first rows as headers", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        [1, 2],
        [3, 4]
      ]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.tables).toHaveLength(1);
    expect(schema.dataRegions[0].hasHeader).toBe(false);
    expect(schema.dataRegions[0].headers).toEqual(["Column1", "Column2"]);
  });

  it("reconciles explicit tables with implicit regions (prefer explicit names, avoid duplicates)", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        ["Product", "Sales"],
        ["Alpha", 10],
        ["Beta", 20]
      ],
      tables: [{ name: "SalesTable", range: "A1:B3" }]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.tables).toHaveLength(1);
    expect(schema.tables[0].name).toBe("SalesTable");
    expect(schema.tables[0].range).toBe("Sheet1!A1:B3");
  });

  it("keeps implicit regions that are not covered by explicit tables", () => {
    const sheet = {
      name: "Sheet1",
      values: [
        ["A", null, null, "X"],
        [1, null, null, 9],
        [null, null, null, null]
      ],
      tables: [{ name: "FirstTable", range: "Sheet1!A1:A2" }]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.tables.map((t: any) => t.name)).toEqual(["FirstTable", "Region1"]);
    expect(schema.tables.map((t: any) => t.range)).toEqual(["Sheet1!A1:A2", "Sheet1!D1:D2"]);
  });

  it("offsets A1 ranges when values are provided as a cropped window (origin)", () => {
    const sheet = {
      name: "Sheet1",
      // This 2x2 matrix represents D11:E12 in the source sheet (0-based origin at row 10, col 3).
      origin: { row: 10, col: 3 },
      values: [
        ["Header", "Value"],
        ["A", 1]
      ]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.dataRegions[0].range).toBe("Sheet1!D11:E12");
    expect(schema.tables[0].range).toBe("Sheet1!D11:E12");
  });

  it("detectDataRegions handles a large contiguous block as a single region", () => {
    const size = 50;
    const values = Array.from({ length: size }, () => Array.from({ length: size }, () => 1));

    const regions = detectDataRegions(values);
    expect(regions).toHaveLength(1);
    expect(regions[0]).toEqual({ startRow: 0, startCol: 0, endRow: size - 1, endCol: size - 1 });
  });

  it("quotes non-identifier sheet names in generated A1 ranges", () => {
    const sheet = {
      name: "My Sheet",
      values: [
        ["Product", "Sales"],
        ["Alpha", 10],
        ["Beta", 20]
      ]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.name).toBe("My Sheet");
    expect(schema.tables[0].range).toBe("'My Sheet'!A1:B3");
    expect(schema.dataRegions[0].range).toBe("'My Sheet'!A1:B3");
  });

  it("escapes single quotes in quoted sheet prefixes", () => {
    const sheet = {
      name: "Bob's Sheet",
      values: [
        ["Product", "Sales"],
        ["Alpha", 10],
        ["Beta", 20]
      ]
    };

    const schema = extractSheetSchema(sheet);
    expect(schema.tables[0].range).toBe("'Bob''s Sheet'!A1:B3");
    expect(schema.dataRegions[0].range).toBe("'Bob''s Sheet'!A1:B3");
  });

  it("detectDataRegions treats missing cells in ragged rows as empty", () => {
    const values = [
      [1, 1],
      [1]
    ];

    const regions = detectDataRegions(values);
    expect(regions).toEqual([{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]);
  });

  it("respects AbortSignal", () => {
    const abortController = new AbortController();
    abortController.abort();

    let dataRegionError: unknown = null;
    try {
      detectDataRegions([[1]], { signal: abortController.signal });
    } catch (error) {
      dataRegionError = error;
    }
    expect(dataRegionError).toMatchObject({ name: "AbortError" });

    let schemaError: unknown = null;
    try {
      extractSheetSchema(
        {
          name: "Sheet1",
          values: [[1]],
        },
        { signal: abortController.signal },
      );
    } catch (error) {
      schemaError = error;
    }
    expect(schemaError).toMatchObject({ name: "AbortError" });
  });
});
