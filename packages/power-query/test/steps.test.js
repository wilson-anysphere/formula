import assert from "node:assert/strict";
import test from "node:test";

import { DataTable } from "../src/table.js";
import { applyOperation } from "../src/steps.js";

function sampleTable() {
  return DataTable.fromGrid(
    [
      ["Region", "Product", "Sales"],
      ["East", "A", 100],
      ["East", "B", 150],
      ["West", "A", 200],
      ["West", "B", 250],
    ],
    { hasHeaders: true, inferTypes: true },
  );
}

test("selectColumns keeps requested columns in order", () => {
  const table = sampleTable();
  const result = applyOperation(table, { type: "selectColumns", columns: ["Product", "Sales"] });
  assert.deepEqual(result.toGrid(), [
    ["Product", "Sales"],
    ["A", 100],
    ["B", 150],
    ["A", 200],
    ["B", 250],
  ]);
});

test("removeColumns drops requested columns", () => {
  const table = sampleTable();
  const result = applyOperation(table, { type: "removeColumns", columns: ["Product"] });
  assert.deepEqual(result.toGrid(), [
    ["Region", "Sales"],
    ["East", 100],
    ["East", 150],
    ["West", 200],
    ["West", 250],
  ]);
});

test("filterRows supports equals predicate", () => {
  const table = sampleTable();
  const result = applyOperation(table, {
    type: "filterRows",
    predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" },
  });
  assert.deepEqual(result.toGrid(), [
    ["Region", "Product", "Sales"],
    ["East", "A", 100],
    ["East", "B", 150],
  ]);
});

test("sortRows sorts by a column descending and stays stable", () => {
  const table = DataTable.fromGrid(
    [
      ["Group", "Value", "Original"],
      ["A", 1, "first"],
      ["A", 1, "second"],
      ["B", 2, "third"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const result = applyOperation(table, {
    type: "sortRows",
    sortBy: [{ column: "Value", direction: "descending" }],
  });

  assert.deepEqual(result.toGrid(), [
    ["Group", "Value", "Original"],
    ["B", 2, "third"],
    ["A", 1, "first"],
    ["A", 1, "second"],
  ]);
});

test("groupBy aggregates values per group", () => {
  const table = sampleTable();
  const result = applyOperation(table, {
    type: "groupBy",
    groupColumns: ["Region"],
    aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }],
  });

  assert.deepEqual(result.toGrid(), [
    ["Region", "Total Sales"],
    ["East", 250],
    ["West", 450],
  ]);
});

test("addColumn evaluates formulas using the sandboxed expression engine", () => {
  const table = sampleTable();
  const result = applyOperation(table, { type: "addColumn", name: "Double", formula: "=[Sales] * 2" });
  assert.deepEqual(result.toGrid(), [
    ["Region", "Product", "Sales", "Double"],
    ["East", "A", 100, 200],
    ["East", "B", 150, 300],
    ["West", "A", 200, 400],
    ["West", "B", 250, 500],
  ]);
});

test("addColumn rejects unsafe identifiers (no global access)", () => {
  const table = sampleTable();
  assert.throws(
    () => applyOperation(table, { type: "addColumn", name: "Bad", formula: "=globalThis" }),
    /Unsupported identifier 'globalThis'/,
  );
});

test("transformColumns evaluates formulas against '_' and can coerce output types", () => {
  const table = DataTable.fromGrid([["Value"], [null], [1]], { hasHeaders: true, inferTypes: false });
  const result = applyOperation(table, {
    type: "transformColumns",
    transforms: [{ column: "Value", formula: "_ == null ? 0 : _ + 1", newType: "number" }],
  });
  assert.deepEqual(result.toGrid(), [["Value"], [0], [2]]);
});

test("changeType coerces values", () => {
  const table = DataTable.fromGrid(
    [
      ["Value"],
      ["1"],
      ["2.5"],
      ["not a number"],
      [""],
    ],
    { hasHeaders: true, inferTypes: false },
  );

  const result = applyOperation(table, { type: "changeType", column: "Value", newType: "number" });
  assert.deepEqual(result.toGrid(), [["Value"], [1], [2.5], [null], [null]]);
});

test("replaceValues matches Date values by timestamp (not identity)", () => {
  const ms1 = Date.parse("2024-01-01T00:00:00.000Z");
  const ms2 = Date.parse("2024-02-01T00:00:00.000Z");
  const table = DataTable.fromGrid(
    [
      ["When"],
      [new Date(ms1)],
      [new Date(ms2)],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const result = applyOperation(table, {
    type: "replaceValues",
    column: "When",
    find: new Date(ms1),
    replace: new Date(ms2),
  });

  assert.deepEqual(result.toGrid(), [["When"], [new Date(ms2)], [new Date(ms2)]]);
});
