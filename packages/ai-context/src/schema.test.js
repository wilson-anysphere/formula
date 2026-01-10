import test from "node:test";
import assert from "node:assert/strict";
import { extractSheetSchema } from "./schema.js";

test("extractSheetSchema detects a headered table region and infers column types", () => {
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

  assert.deepEqual(
    table.columns.map((c) => ({ name: c.name, type: c.type })),
    [
      { name: "Product", type: "string" },
      { name: "Sales", type: "number" },
      { name: "Active", type: "boolean" },
    ],
  );

  assert.deepEqual(schema.namedRanges, [{ name: "SalesData", range: "Sheet1!A1:C3" }]);
  assert.equal(schema.dataRegions.length, 1);
  assert.equal(schema.dataRegions[0].hasHeader, true);
});

test("extractSheetSchema detects multiple disconnected regions", () => {
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
  assert.deepEqual(
    schema.tables.map((t) => t.range),
    ["Sheet1!A1:A2", "Sheet1!D1:D2"],
  );
});

test("extractSheetSchema does not treat numeric-first rows as headers", () => {
  const sheet = {
    name: "Sheet1",
    values: [
      [1, 2],
      [3, 4],
    ],
  };

  const schema = extractSheetSchema(sheet);
  assert.equal(schema.tables.length, 1);
  const region = schema.dataRegions[0];
  assert.equal(region.hasHeader, false);
  assert.deepEqual(region.headers, ["Column1", "Column2"]);
});
