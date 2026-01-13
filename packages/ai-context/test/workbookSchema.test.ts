import { describe, expect, it } from "vitest";

import { extractWorkbookSchema } from "../src/index.js";

describe("extractWorkbookSchema", () => {
  it("infers headers, column types, and row/column counts for workbook tables", () => {
    const workbook = {
      id: "wb1",
      sheets: [
        {
          name: "Sheet1",
          cells: [
            ["Product", "Sales", "Active"],
            ["Alpha", 10, true],
            ["Beta", 20, false],
          ],
        },
      ],
      tables: [
        {
          name: "SalesTable",
          sheetName: "Sheet1",
          rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
        },
      ],
      namedRanges: [{ name: "SalesData", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 2 } }],
    };

    const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 50 });
    expect(schema.id).toBe("wb1");
    expect(schema.sheets).toEqual([{ name: "Sheet1" }]);
    expect(schema.tables).toHaveLength(1);

    const table = schema.tables[0];
    expect(table).toMatchObject({
      name: "SalesTable",
      sheetName: "Sheet1",
      rect: { r0: 0, c0: 0, r1: 2, c1: 2 },
      rangeA1: "Sheet1!A1:C3",
      rowCount: 2,
      columnCount: 3,
    });
    expect(table.headers).toEqual(["Product", "Sales", "Active"]);
    expect(table.inferredColumnTypes).toEqual(["string", "number", "boolean"]);

    expect(schema.namedRanges).toEqual([
      { name: "SalesData", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 2 }, rangeA1: "Sheet1!A1:C3" },
    ]);
  });

  it("is deterministic (stable output independent of input ordering)", () => {
    const sheet = {
      name: "Sheet1",
      cells: [
        ["Name", "Age"],
        ["A", 1],
      ],
    };

    const tableA = { name: "A", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } };
    const tableB = { name: "B", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } };

    const schema1 = extractWorkbookSchema({ id: "wb", sheets: [sheet], tables: [tableB, tableA] });
    const schema2 = extractWorkbookSchema({ id: "wb", sheets: [sheet], tables: [tableA, tableB] });

    expect(schema1).toEqual(schema2);
    expect(schema1.tables.map((t) => t.name)).toEqual(["A", "B"]);
  });

  it("bounds sampling work for very large table rects", () => {
    let readCount = 0;
    const sheet = {
      name: "BigSheet",
      getCell(row: number, col: number) {
        readCount += 1;
        if (row > 10) throw new Error(`scanned too far: row=${row}`);
        if (row === 0) return ["H1", "H2", "H3"][col] ?? null;
        if (row === 1) return ["A", 1, true][col] ?? null;
        if (row === 2) return ["B", 2, false][col] ?? null;
        return null;
      },
    };

    const workbook = {
      id: "wb-big",
      sheets: [sheet],
      tables: [{ name: "BigTable", sheetName: "BigSheet", rect: { r0: 0, c0: 0, r1: 999_999, c1: 2 } }],
    };

    const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 2 });
    expect(schema.tables[0].name).toBe("BigTable");

    // We should only touch:
    // - header row
    // - next row for header detection
    // - a bounded number of sample rows (`maxAnalyzeRows`)
    // Each row touches 3 columns.
    expect(readCount).toBeLessThanOrEqual(30);
  });

  it("bounds sampling work for very wide table rects (maxAnalyzeCols)", () => {
    let readCount = 0;
    const sheet = {
      name: "WideSheet",
      getCell(row: number, col: number) {
        readCount += 1;
        // We should never read beyond the analyzed columns (0..2).
        if (col > 2) throw new Error(`scanned too far: col=${col}`);
        if (row === 0) return ["H1", "H2", "H3"][col] ?? null;
        if (row === 1) return ["A", 1, true][col] ?? null;
        if (row === 2) return ["B", 2, false][col] ?? null;
        return null;
      },
    };

    const workbook = {
      id: "wb-wide",
      sheets: [sheet],
      tables: [{ name: "WideTable", sheetName: "WideSheet", rect: { r0: 0, c0: 0, r1: 2, c1: 999_999 } }],
    };

    const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 2, maxAnalyzeCols: 3 });
    expect(schema.tables[0].name).toBe("WideTable");
    expect(schema.tables[0].headers).toEqual(["H1", "H2", "H3"]);
    expect(schema.tables[0].inferredColumnTypes).toEqual(["string", "number", "boolean"]);
    expect(readCount).toBeLessThanOrEqual(15);
  });

  it("treats empty object cells as empty for header/type inference", () => {
    const workbook = {
      id: "wb-empty-obj",
      sheets: [
        {
          name: "Sheet1",
          cells: [
            [{}, { v: "Sales" }],
            [{}, { v: 10 }],
          ],
        },
      ],
      tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
    };

    const schema = extractWorkbookSchema(workbook);
    expect(schema.tables).toHaveLength(1);
    expect(schema.tables[0].headers).toEqual(["Column1", "Sales"]);
    expect(schema.tables[0].inferredColumnTypes).toEqual(["empty", "number"]);
  });

  it("quotes sheet names in generated A1 ranges (tables + named ranges)", () => {
    const workbook = {
      id: "wb-quoted",
      sheets: [{ name: "Bob's Sheet", cells: [["Header", "Value"], ["A", 1]] }],
      tables: [{ name: "T", sheetName: "Bob's Sheet", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
      namedRanges: [{ name: "NR", sheetName: "Bob's Sheet", rect: { r0: 0, c0: 0, r1: 0, c1: 1 } }],
    };

    const schema = extractWorkbookSchema(workbook);
    expect(schema.tables[0].rangeA1).toBe("'Bob''s Sheet'!A1:B2");
    expect(schema.namedRanges[0].rangeA1).toBe("'Bob''s Sheet'!A1:B1");
  });

  it("supports sparse sheet cell maps (row,col keys)", () => {
    const cells = new Map<string, any>();
    cells.set("0:0", { v: "Name" });
    cells.set("0,1", { v: "Value" });
    cells.set("1,0", { v: "A" });
    cells.set("1:1", { v: 1 });

    const workbook = {
      id: "wb-map",
      sheets: [{ name: "Sheet1", cells }],
      tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
    };

    const schema = extractWorkbookSchema(workbook);
    expect(schema.tables[0]).toMatchObject({ name: "T", rangeA1: "Sheet1!A1:B2", rowCount: 1, columnCount: 2 });
    expect(schema.tables[0].headers).toEqual(["Name", "Value"]);
    expect(schema.tables[0].inferredColumnTypes).toEqual(["string", "number"]);
  });
});
