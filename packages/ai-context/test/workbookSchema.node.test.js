import assert from "node:assert/strict";
import test from "node:test";

import { extractWorkbookSchema } from "../src/index.js";

test("extractWorkbookSchema: infers headers/types/counts for workbook tables", () => {
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
  assert.equal(schema.id, "wb1");
  assert.deepStrictEqual(schema.sheets, [{ name: "Sheet1" }]);
  assert.equal(schema.tables.length, 1);

  const table = schema.tables[0];
  assert.equal(table.name, "SalesTable");
  assert.equal(table.sheetName, "Sheet1");
  assert.deepStrictEqual(table.rect, { r0: 0, c0: 0, r1: 2, c1: 2 });
  assert.equal(table.rangeA1, "Sheet1!A1:C3");
  assert.equal(table.rowCount, 2);
  assert.equal(table.columnCount, 3);
  assert.deepStrictEqual(table.headers, ["Product", "Sales", "Active"]);
  assert.deepStrictEqual(table.inferredColumnTypes, ["string", "number", "boolean"]);

  assert.deepStrictEqual(schema.namedRanges, [
    { name: "SalesData", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 2 }, rangeA1: "Sheet1!A1:C3" },
  ]);
});

test("extractWorkbookSchema: deterministic output independent of input ordering", () => {
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

  assert.deepStrictEqual(schema1, schema2);
  assert.deepStrictEqual(
    schema1.tables.map((t) => t.name),
    ["A", "B"],
  );
});

test("extractWorkbookSchema: sorts tables/named ranges deterministically when start cell + name collide", () => {
  const workbook = {
    id: "wb-colliding",
    sheets: [{ name: "Sheet1", cells: [["H"], [1]] }],
    tables: [
      { name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } },
      { name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
    ],
    namedRanges: [
      { name: "NR", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } },
      { name: "NR", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 0 } },
    ],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(
    schema.tables.map((t) => t.rangeA1),
    ["Sheet1!A1", "Sheet1!A1:A2"],
  );
  assert.deepStrictEqual(
    schema.namedRanges.map((r) => r.rangeA1),
    ["Sheet1!A1", "Sheet1!A1:A2"],
  );
});

test("extractWorkbookSchema: accepts Map-shaped workbook metadata (tables + namedRanges)", () => {
  const tables = new Map();
  tables.set("T", { sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } });

  const namedRanges = new Map();
  namedRanges.set("NR", { sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 1 } });

  const workbook = {
    id: "wb-map-meta",
    sheets: [{ name: "Sheet1", cells: [["Name", "Value"], ["A", 1]] }],
    tables,
    namedRanges,
  };

  const schema = extractWorkbookSchema(workbook);
  assert.equal(schema.tables.length, 1);
  assert.equal(schema.tables[0].name, "T");
  assert.equal(schema.tables[0].sheetName, "Sheet1");
  assert.equal(schema.tables[0].rangeA1, "Sheet1!A1:B2");
  assert.equal(schema.tables[0].rowCount, 1);
  assert.equal(schema.tables[0].columnCount, 2);

  assert.deepStrictEqual(schema.namedRanges, [
    { name: "NR", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 0, c1: 1 }, rangeA1: "Sheet1!A1:B1" },
  ]);
});

test("extractWorkbookSchema: does not call custom toString() on Map keys", () => {
  let toStringCalls = 0;
  const keyObj = {
    toString() {
      toStringCalls += 1;
      return "TopSecretSheet";
    },
  };

  const sheets = new Map();
  sheets.set(keyObj, [[{ v: "Hello" }]]);

  const schema = extractWorkbookSchema({ id: "wb-map-key-tostring", sheets });
  assert.equal(toStringCalls, 0);
  assert.deepStrictEqual(schema.sheets, []);
});

test("extractWorkbookSchema: accepts sheet maps keyed by sheet name (values are matrices)", () => {
  const workbook = {
    id: "wb-sheet-map",
    sheets: {
      Sheet1: [["Header"], [1]],
    },
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.sheets, [{ name: "Sheet1" }]);
  assert.equal(schema.tables[0].rangeA1, "Sheet1!A1:A2");
  assert.equal(schema.tables[0].rowCount, 1);
  assert.equal(schema.tables[0].columnCount, 1);
  assert.deepStrictEqual(schema.tables[0].headers, ["Header"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["number"]);
});

test("extractWorkbookSchema: accepts sheet name lists even without cell data", () => {
  const workbook = {
    id: "wb-sheet-names",
    sheets: ["Sheet1", "Sheet2"],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 9, c1: 1 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.sheets, [{ name: "Sheet1" }, { name: "Sheet2" }]);
  assert.equal(schema.tables.length, 1);
  assert.equal(schema.tables[0].name, "T");
  assert.equal(schema.tables[0].sheetName, "Sheet1");
  assert.equal(schema.tables[0].rangeA1, "Sheet1!A1:B10");
  assert.equal(schema.tables[0].rowCount, 10);
  assert.equal(schema.tables[0].columnCount, 2);
  assert.deepStrictEqual(schema.tables[0].headers, []);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, []);
});

test("extractWorkbookSchema: bounds sampling work for very large table rects", () => {
  let readCount = 0;
  const sheet = {
    name: "BigSheet",
    getCell(row, col) {
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
  assert.equal(schema.tables[0].name, "BigTable");
  assert.ok(readCount <= 30);
});

test("extractWorkbookSchema: bounds sampling work for very wide table rects (maxAnalyzeCols)", () => {
  let readCount = 0;
  const sheet = {
    name: "WideSheet",
    getCell(row, col) {
      readCount += 1;
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
  assert.equal(schema.tables[0].name, "WideTable");
  assert.deepStrictEqual(schema.tables[0].headers, ["H1", "H2", "H3"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number", "boolean"]);
  assert.ok(readCount <= 15);
});

test("extractWorkbookSchema: treats empty object cells as empty for header/type inference", () => {
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
  assert.equal(schema.tables.length, 1);
  assert.deepStrictEqual(schema.tables[0].headers, ["Column1", "Sales"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["empty", "number"]);
});

test("extractWorkbookSchema: unwraps typed value encodings for schema inference (t:n/t:b/t:blank/t:e)", () => {
  const workbook = {
    id: "wb-typed-values",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          ["Num", "Bool", "Blank", "Err"],
          [
            { value: { t: "n", v: 1.5 }, display: "1.50" },
            { value: { t: "b", v: true }, display: "TRUE" },
            { value: { t: "blank" } },
            { value: { t: "e", v: "#DIV/0!" } },
          ],
        ],
      },
    ],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 3 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.tables[0].headers, ["Num", "Bool", "Blank", "Err"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["number", "boolean", "empty", "string"]);
});

test("extractWorkbookSchema: treats rich text + in-cell image values as strings for header/type inference", () => {
  const workbook = {
    id: "wb-rich-values",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          [{ text: "Product", runs: [{ start: 0, end: 7, style: { bold: true } }] }, { type: "image", value: { imageId: "img_1", altText: "Photo" } }, "Qty"],
          ["Alpha", { type: "image", value: { imageId: "img_2" } }, 10],
        ],
      },
    ],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 2 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.tables[0].headers, ["Product", "Photo", "Qty"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "string", "number"]);
});

test("extractWorkbookSchema: quotes sheet names in generated A1 ranges (tables + named ranges)", () => {
  const workbook = {
    id: "wb-quoted",
    sheets: [{ name: "Bob's Sheet", cells: [["Header", "Value"], ["A", 1]] }],
    tables: [{ name: "T", sheetName: "Bob's Sheet", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
    namedRanges: [{ name: "NR", sheetName: "Bob's Sheet", rect: { r0: 0, c0: 0, r1: 0, c1: 1 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.equal(schema.tables[0].rangeA1, "'Bob''s Sheet'!A1:B2");
  assert.equal(schema.namedRanges[0].rangeA1, "'Bob''s Sheet'!A1:B1");
});

test("extractWorkbookSchema: supports sparse sheet cell maps (row,col keys)", () => {
  const cells = new Map();
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
  assert.equal(schema.tables[0].name, "T");
  assert.equal(schema.tables[0].rangeA1, "Sheet1!A1:B2");
  assert.equal(schema.tables[0].rowCount, 1);
  assert.equal(schema.tables[0].columnCount, 2);
  assert.deepStrictEqual(schema.tables[0].headers, ["Name", "Value"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number"]);
});

test("extractWorkbookSchema: supports origin-offset sheet matrices (rects use absolute coordinates)", () => {
  const workbook = {
    id: "wb-origin",
    sheets: [
      {
        name: "Sheet1",
        origin: { row: 10, col: 5 },
        values: [
          ["Name", "Amount"],
          ["A", 1],
        ],
      },
    ],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 10, c0: 5, r1: 11, c1: 6 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.equal(schema.tables[0].rangeA1, "Sheet1!F11:G12");
  assert.equal(schema.tables[0].rowCount, 1);
  assert.equal(schema.tables[0].columnCount, 2);
  assert.deepStrictEqual(schema.tables[0].headers, ["Name", "Amount"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number"]);
});

test("extractWorkbookSchema: accepts range-shaped rects (startRow/startCol/endRow/endCol)", () => {
  const workbook = {
    id: "wb-range-rect",
    sheets: [{ name: "Sheet1", values: [["H1", "H2"], ["A", 1]] }],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { startRow: 0, startCol: 0, endRow: 1, endCol: 1 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.equal(schema.tables[0].rangeA1, "Sheet1!A1:B2");
  assert.equal(schema.tables[0].rowCount, 1);
  assert.equal(schema.tables[0].columnCount, 2);
  assert.deepStrictEqual(schema.tables[0].headers, ["H1", "H2"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number"]);
});

test("extractWorkbookSchema: supports sparse sheet cell object maps (row,col keys)", () => {
  const cells = {};
  cells["0:0"] = { v: "Name" };
  cells["0,1"] = { v: "Value" };
  cells["1,0"] = { v: "A" };
  cells["1:1"] = { v: 1 };

  const workbook = {
    id: "wb-obj-map",
    sheets: [{ name: "Sheet1", cells }],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.tables[0].headers, ["Name", "Value"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number"]);
});

test("extractWorkbookSchema: supports Map-like sheets (cells.get)", () => {
  const backing = new Map();
  backing.set("0:0", { v: "Name" });
  backing.set("0,1", { v: "Value" });
  backing.set("1,0", { v: "A" });
  backing.set("1:1", { v: 1 });

  const cells = {
    get(key) {
      return backing.get(key);
    },
  };

  const workbook = {
    id: "wb-map-like",
    sheets: [{ name: "Sheet1", cells }],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 1 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.tables[0].headers, ["Name", "Value"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number"]);
});

test("extractWorkbookSchema: infers formula columns when cells contain formulas", () => {
  const workbook = {
    id: "wb-formula",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          ["Item", "Price", "Tax", "Total"],
          ["A", 10, 0.1, "=B2*(1+C2)"],
          ["B", 20, 0.2, "=B3*(1+C3)"],
        ],
      },
    ],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 3 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.tables[0].headers, ["Item", "Price", "Tax", "Total"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["string", "number", "number", "formula"]);
});

test("extractWorkbookSchema: infers ISO-like date strings as dates", () => {
  const workbook = {
    id: "wb-date",
    sheets: [
      {
        name: "Sheet1",
        cells: [
          ["Date", "Amount"],
          ["2025-01-01", 10],
          ["2025-01-02", 20],
        ],
      },
    ],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 1 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.deepStrictEqual(schema.tables[0].headers, ["Date", "Amount"]);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, ["date", "number"]);
});

test("extractWorkbookSchema: respects AbortSignal", () => {
  const workbook = {
    id: "wb-abort",
    sheets: [{ name: "Sheet1", cells: [["Header"], [1]] }],
    tables: [{ name: "T", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 1, c1: 0 } }],
  };

  const abortController = new AbortController();
  abortController.abort();

  let error = null;
  try {
    extractWorkbookSchema(workbook, { signal: abortController.signal });
  } catch (err) {
    error = err;
  }

  assert.ok(error && typeof error === "object");
  assert.equal(error.name, "AbortError");
});

test("extractWorkbookSchema: includes tables even when sheet cell data is unavailable", () => {
  const workbook = {
    id: "wb-missing-sheet",
    sheets: [{ name: "Other", cells: [[1]] }],
    tables: [{ name: "T", sheetName: "Missing", rect: { r0: 0, c0: 0, r1: 9, c1: 2 } }],
  };

  const schema = extractWorkbookSchema(workbook);
  assert.equal(schema.tables.length, 1);
  assert.equal(schema.tables[0].name, "T");
  assert.equal(schema.tables[0].sheetName, "Missing");
  assert.equal(schema.tables[0].rangeA1, "Missing!A1:C10");
  assert.equal(schema.tables[0].rowCount, 10);
  assert.equal(schema.tables[0].columnCount, 3);
  assert.deepStrictEqual(schema.tables[0].headers, []);
  assert.deepStrictEqual(schema.tables[0].inferredColumnTypes, []);
});
