import assert from "node:assert/strict";
import test from "node:test";

import { detectDataRegions, extractSheetSchema } from "../src/schema.js";

test("extractSheetSchema: detects a headered table region and infers column types", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      ["Product", "Sales", "Active"],
      ["Alpha", 10, true],
      ["Beta", 20, false],
    ],
    namedRanges: [{ name: "SalesData", range: "Sheet1!A1:C3" }],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.name, "Sheet1");
  assert.equal(schema.tables.length, 1);

  const table = schema.tables[0];
  assert.equal(table.range, "Sheet1!A1:C3");
  assert.equal(table.rowCount, 2);
  assert.deepStrictEqual(
    table.columns.map((c) => ({ name: c.name, type: c.type })),
    [
      { name: "Product", type: "string" },
      { name: "Sales", type: "number" },
      { name: "Active", type: "boolean" },
    ],
  );

  assert.deepStrictEqual(schema.namedRanges, [{ name: "SalesData", range: "Sheet1!A1:C3" }]);
  assert.equal(schema.dataRegions.length, 1);
  assert.equal(schema.dataRegions[0].hasHeader, true);
});

test("extractSheetSchema: treats rich text + in-cell image values as strings", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      [
        { text: "Product", runs: [{ start: 0, end: 7, style: { bold: true } }] },
        { type: "image", value: { imageId: "img_1", altText: " Photo " } },
        "Qty",
      ],
      ["Alpha", { imageId: "img_2" }, 3],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.dataRegions.length, 1);
  assert.equal(schema.dataRegions[0].hasHeader, true);
  assert.deepStrictEqual(schema.dataRegions[0].headers, ["Product", "Photo", "Qty"]);

  assert.equal(schema.tables.length, 1);
  const table = schema.tables[0];
  assert.deepStrictEqual(
    table.columns.map((c) => ({ name: c.name, type: c.type })),
    [
      { name: "Product", type: "string" },
      { name: "Photo", type: "string" },
      { name: "Qty", type: "number" },
    ],
  );
  const photoCol = table.columns.find((c) => c.name === "Photo");
  assert.ok(photoCol, "expected Photo column");
  assert.ok(photoCol.sampleValues.includes("[Image]"));
});

test("extractSheetSchema: does not call custom toString() on non-string header cells", () => {
  const dangerous = {
    toString() {
      throw new Error("toString should not be called");
    },
  };

  const sheet = {
    name: "Sheet1",
    values: [
      [dangerous, "Name", "Age"],
      ["Alice", 30, 5],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.dataRegions.length, 1);
  assert.equal(schema.dataRegions[0].hasHeader, true);
  assert.deepStrictEqual(schema.dataRegions[0].headers, ["Column1", "Name", "Age"]);
});

test("extractSheetSchema: detects multiple disconnected regions", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      ["A", null, null, "X"],
      [1, null, null, 9],
      [null, null, null, null],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.tables.length, 2);
  assert.deepStrictEqual(
    schema.tables.map((t) => t.range),
    ["Sheet1!A1:A2", "Sheet1!D1:D2"],
  );
});

test("extractSheetSchema: does not treat numeric-first rows as headers", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      [1, 2],
      [3, 4],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.tables.length, 1);
  assert.equal(schema.dataRegions[0].hasHeader, false);
  assert.deepStrictEqual(schema.dataRegions[0].headers, ["Column1", "Column2"]);
});

test("extractSheetSchema: reconciles explicit tables with implicit regions (prefer explicit names, avoid duplicates)", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      ["Product", "Sales"],
      ["Alpha", 10],
      ["Beta", 20],
    ],
    tables: [{ name: "SalesTable", range: "A1:B3" }],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.tables.length, 1);
  assert.equal(schema.tables[0].name, "SalesTable");
  assert.equal(schema.tables[0].range, "Sheet1!A1:B3");
});

test("extractSheetSchema: keeps implicit regions that are not covered by explicit tables", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      ["A", null, null, "X"],
      [1, null, null, 9],
      [null, null, null, null],
    ],
    tables: [{ name: "FirstTable", range: "Sheet1!A1:A2" }],
  };

  const schema = extractSheetSchema(sheet);
  assert.deepStrictEqual(
    schema.tables.map((t) => t.name),
    ["FirstTable", "Region1"],
  );
  assert.deepStrictEqual(
    schema.tables.map((t) => t.range),
    ["Sheet1!A1:A2", "Sheet1!D1:D2"],
  );
});

test("extractSheetSchema: offsets A1 ranges when values are provided as a cropped window (origin)", () => {
  const sheet = {
    name: "Sheet1",
    // This 2x2 matrix represents D11:E12 in the source sheet (0-based origin at row 10, col 3).
    origin: { row: 10, col: 3 },
    values: [
      ["Header", "Value"],
      ["A", 1],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.dataRegions[0].range, "Sheet1!D11:E12");
  assert.equal(schema.tables[0].range, "Sheet1!D11:E12");
});

test("detectDataRegions: handles a large contiguous block as a single region", () => {
  const size = 50;
  const values = Array.from({ length: size }, () => Array.from({ length: size }, () => 1));

  const regions = detectDataRegions(values);
  assert.equal(regions.length, 1);
  assert.deepStrictEqual(regions[0], { startRow: 0, startCol: 0, endRow: size - 1, endCol: size - 1 });
});

test("extractSheetSchema: quotes non-identifier sheet names in generated A1 ranges", () => {
  const sheet = {
    name: "My Sheet",
    values: [
      ["Product", "Sales"],
      ["Alpha", 10],
      ["Beta", 20],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.name, "My Sheet");
  assert.equal(schema.tables[0].range, "'My Sheet'!A1:B3");
  assert.equal(schema.dataRegions[0].range, "'My Sheet'!A1:B3");
});

test("extractSheetSchema: escapes single quotes in quoted sheet prefixes", () => {
  const sheet = {
    name: "Bob's Sheet",
    values: [
      ["Product", "Sales"],
      ["Alpha", 10],
      ["Beta", 20],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.tables[0].range, "'Bob''s Sheet'!A1:B3");
  assert.equal(schema.dataRegions[0].range, "'Bob''s Sheet'!A1:B3");
});

test("extractSheetSchema: analyzes large regions using bounded sampling and preserves full row/column counts", () => {
  const rows = 4_000;
  const cols = 50;
  const headerRow = Array.from({ length: cols }, (_v, i) => {
    if (i === 0) return "Name";
    if (i === 1) return "Sales";
    if (i === 2) return "Active";
    if (i === 3) return "Date";
    if (i === 4) return "Formula";
    return `Col${i + 1}`;
  });

  const values = Array.from({ length: rows }, (_v, r) => {
    if (r === 0) return headerRow;
    const row = Array.from({ length: cols }, () => null);
    // Keep the region connected across all rows.
    row[0] = `Item${r}`;
    if (r === 1) {
      row[1] = 123;
      row[2] = true;
      row[3] = "2024-01-01";
      row[4] = "=A2*2";
    }
    // Conflicting type outside the sampled prefix should not influence inference.
    if (r === 100) {
      row[1] = "oops";
    }
    return row;
  });

  const schema = extractSheetSchema({ name: "Sheet1", values }, { maxAnalyzeRows: 5, maxSampleValuesPerColumn: 1 });

  assert.equal(schema.dataRegions.length, 1);
  assert.equal(schema.dataRegions[0].rowCount, rows - 1);
  assert.equal(schema.dataRegions[0].columnCount, cols);

  assert.equal(schema.tables.length, 1);
  const table = schema.tables[0];
  assert.equal(table.rowCount, rows - 1);
  assert.equal(table.columns.length, cols);

  const colByName = Object.fromEntries(table.columns.map((c) => [c.name, c]));
  assert.equal(colByName.Name.type, "string");
  assert.equal(colByName.Sales.type, "number");
  assert.equal(colByName.Active.type, "boolean");
  assert.equal(colByName.Date.type, "date");
  assert.equal(colByName.Formula.type, "formula");
  // sampleValues should respect maxSampleValuesPerColumn=1.
  assert.ok(colByName.Name.sampleValues.length <= 1);
  assert.ok(colByName.Sales.sampleValues.length <= 1);
});

test("detectDataRegions: treats missing cells in ragged rows as empty", () => {
  const values = [
    [1, 1],
    [1],
  ];

  const regions = detectDataRegions(values);
  assert.deepStrictEqual(regions, [{ startRow: 0, startCol: 0, endRow: 1, endCol: 1 }]);
});

test("schema extraction: respects AbortSignal", () => {
  const abortController = new AbortController();
  abortController.abort();

  let dataRegionError = null;
  try {
    detectDataRegions([[1]], { signal: abortController.signal });
  } catch (error) {
    dataRegionError = error;
  }
  assert.ok(dataRegionError && typeof dataRegionError === "object");
  assert.equal(dataRegionError.name, "AbortError");

  let schemaError = null;
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
  assert.ok(schemaError && typeof schemaError === "object");
  assert.equal(schemaError.name, "AbortError");
});
