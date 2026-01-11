import assert from "node:assert/strict";
import test from "node:test";

import { arrowTableFromColumns } from "../../data-io/src/index.js";

import { QueryEngine } from "../src/engine.js";
import { applyOperation } from "../src/steps.js";
import { ArrowTableAdapter } from "../src/arrowTable.js";
import { DataTable } from "../src/table.js";

function sampleRows() {
  return [
    { Region: "East", Product: "A", Sales: 100 },
    { Region: "East", Product: "B", Sales: 150 },
    { Region: "West", Product: "A", Sales: 200 },
    { Region: "West", Product: "B", Sales: 250 },
  ];
}

function sampleDataTable() {
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

function sampleArrowTable() {
  const rows = sampleRows();
  return new ArrowTableAdapter(
    arrowTableFromColumns({
      Region: rows.map((r) => r.Region),
      Product: rows.map((r) => r.Product),
      Sales: rows.map((r) => r.Sales),
    }),
  );
}

test("Arrow backend: core steps match DataTable results", () => {
  const dataTable = sampleDataTable();
  const arrowTable = sampleArrowTable();

  const operations = [
    { type: "selectColumns", columns: ["Product", "Sales"] },
    { type: "removeColumns", columns: ["Product"] },
    { type: "renameColumn", oldName: "Product", newName: "Item" },
    { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
    { type: "sortRows", sortBy: [{ column: "Sales", direction: "descending" }] },
    { type: "groupBy", groupColumns: ["Region"], aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }] },
  ];

  for (const op of operations) {
    const expected = applyOperation(dataTable, op).toGrid();
    const actual = applyOperation(arrowTable, op).toGrid();
    assert.deepEqual(actual, expected);
  }
});

test("Arrow backend: changeType matches DataTable results", () => {
  const dataTable = DataTable.fromGrid(
    [
      ["Value"],
      ["1"],
      ["2.5"],
      ["not a number"],
      [""],
    ],
    { hasHeaders: true, inferTypes: false },
  );

  const arrowTable = new ArrowTableAdapter(
    arrowTableFromColumns({
      Value: ["1", "2.5", "not a number", ""],
    }),
  );

  const expected = applyOperation(dataTable, { type: "changeType", column: "Value", newType: "number" }).toGrid();
  const actual = applyOperation(arrowTable, { type: "changeType", column: "Value", newType: "number" }).toGrid();
  assert.deepEqual(actual, expected);
});

test("Arrow backend: int64/BigInt values are normalized for aggregations", () => {
  const dataTable = DataTable.fromGrid(
    [
      ["Group", "Value"],
      ["A", 1],
      ["A", 2],
      ["B", 3],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const arrowTable = new ArrowTableAdapter(
    arrowTableFromColumns({
      Group: ["A", "A", "B"],
      Value: [1n, 2n, 3n],
    }),
  );

  const op = { type: "groupBy", groupColumns: ["Group"], aggregations: [{ column: "Value", op: "sum", as: "Total" }] };
  assert.deepEqual(applyOperation(arrowTable, op).toGrid(), applyOperation(dataTable, op).toGrid());
});

test("Arrow backend: int64 group keys do not break JSON stringification", () => {
  const dataTable = DataTable.fromGrid(
    [
      ["id", "value"],
      ["9007199254740993", 1],
      ["9007199254740994", 2],
    ],
    { hasHeaders: true, inferTypes: false },
  );

  const arrowTable = new ArrowTableAdapter(
    arrowTableFromColumns({
      id: [9007199254740993n, 9007199254740994n],
      value: [1, 2],
    }),
  );

  const op = { type: "groupBy", groupColumns: ["id"], aggregations: [{ column: "value", op: "count", as: "Count" }] };
  assert.deepEqual(applyOperation(arrowTable, op).toGrid(), applyOperation(dataTable, op).toGrid());
});

test("QueryEngine: identical results across backends for a multi-step query", async () => {
  const engine = new QueryEngine();
  const dataTable = sampleDataTable();
  const arrowTable = sampleArrowTable();

  const baseQuery = {
    id: "q_test",
    name: "Test",
    source: { type: "table", table: "" },
    steps: [
      { id: "s_select", name: "Reorder", operation: { type: "selectColumns", columns: ["Sales", "Region", "Product"] } },
      { id: "s_remove", name: "Drop Product", operation: { type: "removeColumns", columns: ["Product"] } },
      { id: "s_type", name: "Sales as string", operation: { type: "changeType", column: "Sales", newType: "string" } },
      {
        id: "s_filter",
        name: "Filter East",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
      {
        id: "s_group",
        name: "Group",
        operation: { type: "groupBy", groupColumns: ["Region"], aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }] },
      },
      { id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Total Sales", direction: "descending" }] } },
    ],
  };

  const context = { tables: { data: dataTable, arrow: arrowTable } };
  const dataResult = await engine.executeQuery({ ...baseQuery, source: { type: "table", table: "data" } }, context, {});
  const arrowResult = await engine.executeQuery({ ...baseQuery, source: { type: "table", table: "arrow" } }, context, {});

  assert.deepEqual(arrowResult.toGrid(), dataResult.toGrid());
  assert.ok(arrowResult instanceof ArrowTableAdapter);
});
