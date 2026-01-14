import assert from "node:assert/strict";
import test from "node:test";

import { summarizeWorkbookSchema } from "../src/summarizeWorkbookSchema.js";
import { extractWorkbookSchema } from "../src/workbookSchema.js";

test("summarizeWorkbookSchema: summarizes workbook schema deterministically", () => {
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

  assert.equal(summary1, summary2);
  assert.match(summary1, /^workbook=\[wb1\] sheets=2 tables=2 named=1/m);
  assert.match(summary1, /^s=\[SheetA\|SheetB\]$/m);
  assert.match(summary1, /^T1 \[T2\] r=\[SheetA!A1:B2\]/m);
  assert.match(summary1, /^T2 \[T1\] r=\[SheetB!A1:B2\]/m);
  assert.match(summary1, /^N1 \[NR\] r=\[SheetA!A1:B1\]$/m);
});

test("summarizeWorkbookSchema: escapes bracketed fields to keep summaries parseable", () => {
  const workbook = {
    id: "wb|1]",
    sheets: [{ name: "A|B]", cells: [["Header"], [1]] }],
    tables: [{ name: "T|1]", sheetName: "A|B]", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } }],
  };

  const schema = extractWorkbookSchema(workbook, { maxAnalyzeRows: 5, maxAnalyzeCols: 5 });
  const summary = summarizeWorkbookSchema(schema, { includeNamedRanges: false });

  const lines = summary.split("\n");
  assert.equal(lines[0], "workbook=[wb\\|1\\]] sheets=1 tables=1 named=0");
  assert.ok(lines.includes("s=[A\\|B\\]]"));
  assert.match(summary, /^T1 \[T\\\|1\\\]\] r=\['A\\\|B\\\]'\!A1:A2\]/m);
});

test("summarizeWorkbookSchema: respects output limits", () => {
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

  assert.match(summary, /^workbook=\[wb2\] sheets=1 tables=2 named=2/m);
  assert.match(summary, /^s=\[…\]$/m);
  assert.match(summary, /^T1 /m);
  assert.ok(!/^T2 /m.test(summary));
  assert.match(summary, /^N1 /m);
  assert.ok(!/^N2 /m.test(summary));
});

