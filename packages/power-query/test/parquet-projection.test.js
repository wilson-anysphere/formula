import assert from "node:assert/strict";
import test from "node:test";

import { computeParquetProjectionColumns, computeParquetRowLimit } from "../src/parquetProjection.js";

test("computeParquetProjectionColumns returns null without an explicit projection", () => {
  const steps = [
    { id: "s_filter", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } } },
  ];

  assert.equal(computeParquetProjectionColumns(steps), null);
});

test("computeParquetProjectionColumns unions referenced columns across supported ops", () => {
  const steps = [
    { id: "s_filter", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "active", operator: "equals", value: true } } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "name", "score"] } },
    { id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "score", direction: "descending" }] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["active", "id", "name", "score"]));
});

test("computeParquetProjectionColumns maps renamed columns back to parquet source names", () => {
  const steps = [
    { id: "s_rename", name: "Rename", operation: { type: "renameColumn", oldName: "score", newName: "Score" } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id", "Score"] } },
    { id: "s_type", name: "Type", operation: { type: "changeType", column: "Score", newType: "number" } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["id", "score"]));
});

test("computeParquetProjectionColumns supports groupBy + downstream references to aggregation aliases", () => {
  const steps = [
    {
      id: "s_group",
      name: "Group",
      operation: {
        type: "groupBy",
        groupColumns: ["Region"],
        aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }],
      },
    },
    { id: "s_sort", name: "Sort", operation: { type: "sortRows", sortBy: [{ column: "Total Sales", direction: "descending" }] } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["Region", "Total Sales"] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["Region", "Sales"]));
});

test("computeParquetProjectionColumns supports addColumn and does not request derived columns from parquet", () => {
  const steps = [
    { id: "s_add", name: "Add", operation: { type: "addColumn", name: "Total", formula: "=[a] + [b]" } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["Total"] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["a", "b"]));
});

test("computeParquetProjectionColumns supports transformColumns", () => {
  const steps = [
    {
      id: "s_transform",
      name: "Transform",
      operation: { type: "transformColumns", transforms: [{ column: "a", formula: "_ + 1", newType: null }] },
    },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["a"] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["a"]));
});

test("computeParquetProjectionColumns supports distinctRows + removeRowsWithErrors on an explicit schema", () => {
  const steps = [
    { id: "s_select1", name: "Select", operation: { type: "selectColumns", columns: ["a", "b"] } },
    { id: "s_distinct", name: "Distinct", operation: { type: "distinctRows", columns: null } },
    { id: "s_remove_errors", name: "Remove errors", operation: { type: "removeRowsWithErrors", columns: null } },
    { id: "s_select2", name: "Select2", operation: { type: "selectColumns", columns: ["a"] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["a", "b"]));
});

test("computeParquetProjectionColumns refuses distinctRows over unknown schema", () => {
  const steps = [
    { id: "s_distinct", name: "Distinct", operation: { type: "distinctRows", columns: null } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["a"] } },
  ];

  assert.equal(computeParquetProjectionColumns(steps), null);
});

test("computeParquetProjectionColumns supports fillDown and replaceValues", () => {
  const steps = [
    { id: "s_fill", name: "Fill", operation: { type: "fillDown", columns: ["a"] } },
    { id: "s_replace", name: "Replace", operation: { type: "replaceValues", column: "a", find: null, replace: 0 } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["a"] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["a"]));
});

test("computeParquetProjectionColumns maps renamed columns through replaceValues", () => {
  const steps = [
    { id: "s_rename", name: "Rename", operation: { type: "renameColumn", oldName: "a", newName: "A" } },
    { id: "s_replace", name: "Replace", operation: { type: "replaceValues", column: "A", find: "x", replace: "y" } },
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["A"] } },
  ];

  const cols = computeParquetProjectionColumns(steps);
  assert.ok(cols);
  assert.deepEqual(new Set(cols), new Set(["a"]));
});

test("computeParquetProjectionColumns returns null when unsupported operations are present", () => {
  const steps = [
    { id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id"] } },
    { id: "s_split", name: "Split", operation: { type: "splitColumn", column: "id", delimiter: "-" } },
  ];

  assert.equal(computeParquetProjectionColumns(steps), null);
});

test("computeParquetRowLimit pushes down execute limit for row-preserving pipelines", () => {
  const steps = [{ id: "s_select", name: "Select", operation: { type: "selectColumns", columns: ["id"] } }];
  assert.equal(computeParquetRowLimit(steps, 100), 100);
});

test("computeParquetRowLimit incorporates take() counts", () => {
  const steps = [
    { id: "s_take", name: "Take", operation: { type: "take", count: 10 } },
    { id: "s_rename", name: "Rename", operation: { type: "renameColumn", oldName: "a", newName: "b" } },
  ];
  assert.equal(computeParquetRowLimit(steps, 100), 10);
});

test("computeParquetRowLimit refuses to push down for filter/sort/group", () => {
  const steps = [
    { id: "s_filter", name: "Filter", operation: { type: "filterRows", predicate: { type: "comparison", column: "x", operator: "isNotNull" } } },
  ];
  assert.equal(computeParquetRowLimit(steps, 100), null);
});

test("computeParquetRowLimit refuses to push down for distinctRows/removeRowsWithErrors", () => {
  assert.equal(
    computeParquetRowLimit([{ id: "s1", name: "Distinct", operation: { type: "distinctRows", columns: ["a"] } }], 100),
    null,
  );
  assert.equal(
    computeParquetRowLimit([{ id: "s1", name: "Errors", operation: { type: "removeRowsWithErrors", columns: ["a"] } }], 100),
    null,
  );
});

test("computeParquetRowLimit accounts for skip/removeRows/promoteHeaders", () => {
  assert.equal(
    computeParquetRowLimit([{ id: "s_skip", name: "Skip", operation: { type: "skip", count: 5 } }], 100),
    105,
  );

  assert.equal(
    computeParquetRowLimit(
      [
        { id: "s_skip", name: "Skip", operation: { type: "skip", count: 5 } },
        { id: "s_take", name: "Take", operation: { type: "take", count: 10 } },
      ],
      100,
    ),
    15,
  );

  assert.equal(
    computeParquetRowLimit(
      [
        { id: "s_take", name: "Take", operation: { type: "take", count: 10 } },
        { id: "s_skip", name: "Skip", operation: { type: "skip", count: 5 } },
      ],
      100,
    ),
    10,
  );

  assert.equal(
    computeParquetRowLimit([{ id: "s_remove", name: "Remove", operation: { type: "removeRows", offset: 50, count: 10 } }], 100),
    110,
  );
  assert.equal(
    computeParquetRowLimit([{ id: "s_remove", name: "Remove", operation: { type: "removeRows", offset: 150, count: 10 } }], 100),
    100,
  );

  assert.equal(
    computeParquetRowLimit([{ id: "s_headers", name: "Headers", operation: { type: "promoteHeaders" } }], 100),
    101,
  );
});

test("computeParquetRowLimit refuses to push down for merge/append", () => {
  assert.equal(
    computeParquetRowLimit([{ id: "s_merge", name: "Merge", operation: { type: "merge", rightQuery: "q2", joinType: "left", leftKey: "a", rightKey: "b" } }], 100),
    null,
  );
  assert.equal(
    computeParquetRowLimit([{ id: "s_append", name: "Append", operation: { type: "append", queries: ["q2"] } }], 100),
    null,
  );
});
