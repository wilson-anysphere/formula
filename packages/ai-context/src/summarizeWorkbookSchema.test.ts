import { describe, expect, it } from "vitest";

import { extractWorkbookSchema } from "./workbookSchema.js";
import { summarizeWorkbookSchema } from "./summarizeWorkbookSchema.js";

describe("summarizeWorkbookSchema", () => {
  it("summarizes workbook schema deterministically", () => {
    const workbook = {
      id: "wb1",
      sheets: [
        { name: "SheetB", cells: [["Name", "Age"], ["A", 1]] },
        { name: "SheetA", cells: [["Product", "Sales"], ["Alpha", 10]] },
      ],
      tables: [
        { name: "T2", sheetName: "SheetA", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } },
        { name: "T1", sheetName: "SheetB", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } },
      ],
      namedRanges: [{ name: "NR", sheetName: "SheetA", rect: { r0: 0, c0: 0, r1: 0, c1: 1 } }],
    };

    const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 5, maxAnalyzeCols: 5 });
    const summary1 = summarizeWorkbookSchema(schema);
    const summary2 = summarizeWorkbookSchema(schema);

    expect(summary1).toBe(summary2);
    expect(summary1).toMatch(/^workbook=\[wb1\] sheets=2 tables=2 named=1/m);
    expect(summary1).toMatch(/^s=\[SheetA\|SheetB\]$/m);
    expect(summary1).toMatch(/^T1 \[T2\] r=\[SheetA!A1:B2\]/m);
    expect(summary1).toMatch(/^T2 \[T1\] r=\[SheetB!A1:B2\]/m);
    expect(summary1).toMatch(/^N1 \[NR\] r=\[SheetA!A1:B1\]$/m);
  });

  it("escapes bracketed fields to keep summaries parseable", () => {
    const workbook = {
      id: "wb|1]",
      sheets: [{ name: "A|B]", cells: [["Header"], [1]] }],
      tables: [{ name: "T|1]", sheetName: "A|B]", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } }],
    };

    const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 5, maxAnalyzeCols: 5 });
    const summary = summarizeWorkbookSchema(schema, { includeNamedRanges: false });

    const lines = summary.split("\n");
    expect(lines[0]).toBe("workbook=[wb\\|1\\]] sheets=1 tables=1 named=0");
    expect(lines).toContain("s=[A\\|B\\]]");
    expect(summary).toMatch(/^T1 \[T\\\|1\\\]\] r=\['A\\\|B\\\]'\!A1:A2\]/m);
  });

  it("respects output limits", () => {
    const workbook = {
      id: "wb2",
      sheets: [{ name: "Sheet1", cells: [["A", "B"], [1, 2]] }],
      tables: [
        { name: "T1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } },
        { name: "T2", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } },
      ],
      namedRanges: [
        { name: "NR1", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
        { name: "NR2", sheetName: "Sheet1", rect: { r0: 0, c0: 1, r1: 0, c1: 1 } },
      ],
    };

    const schema = extractWorkbookSchema(workbook);
    const summary = summarizeWorkbookSchema(schema, { maxTables: 1, maxNamedRanges: 1, maxSheets: 0 });

    expect(summary).toMatch(/^workbook=\[wb2\] sheets=1 tables=2 named=2/m);
    expect(summary).toMatch(/^s=\[…\]$/m);
    expect(summary).toMatch(/^T1 /m);
    expect(summary).not.toMatch(/^T2 /m);
    expect(summary).toMatch(/^N1 /m);
    expect(summary).not.toMatch(/^N2 /m);
  });
});
