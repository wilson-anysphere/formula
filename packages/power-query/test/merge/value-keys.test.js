import assert from "node:assert/strict";
import test from "node:test";

import { QueryEngine } from "../../src/engine.js";
import { ArrowTableAdapter } from "../../src/arrowTable.js";
import { DataTable } from "../../src/table.js";

/**
 * Build a minimal Arrow-like table object sufficient for `ArrowTableAdapter`.
 * This keeps the test focused on the adapter's behavior (not on `apache-arrow`'s
 * concrete JS types) and ensures date cells materialize as distinct `Date`
 * instances per read.
 *
 * @param {Record<string, unknown[]>} columns
 * @param {Record<string, string>} typeHints
 */
function makeFakeArrowTable(columns, typeHints) {
  const names = Object.keys(columns);
  const rowCount = Math.max(0, ...names.map((name) => columns[name]?.length ?? 0));

  return {
    numRows: rowCount,
    schema: {
      fields: names.map((name) => ({
        name,
        type: { toString: () => typeHints[name] ?? "Utf8" },
      })),
    },
    getChildAt: (index) => {
      const name = names[index];
      const values = columns[name] ?? [];
      return {
        length: rowCount,
        get: (rowIndex) => values[rowIndex],
      };
    },
    slice: (start, end) => {
      /** @type {Record<string, unknown[]>} */
      const sliced = {};
      for (const name of names) {
        sliced[name] = (columns[name] ?? []).slice(start, end);
      }
      return makeFakeArrowTable(sliced, typeHints);
    },
  };
}

/**
 * @param {{ left: import("../../src/table.js").ITable, right: import("../../src/table.js").ITable, leftKey: string, rightKey: string }}
 */
async function runMerge({ left, right, leftKey, rightKey }) {
  const engine = new QueryEngine();
  const rightQuery = { id: "q_right", name: "Right", source: { type: "table", table: "right" }, steps: [] };
  const leftQuery = {
    id: "q_left",
    name: "Left",
    source: { type: "table", table: "left" },
    steps: [
      {
        id: "s_merge",
        name: "Merge",
        operation: { type: "merge", rightQuery: "q_right", joinType: "inner", leftKey, rightKey },
      },
    ],
  };

  return engine.executeQuery(leftQuery, { tables: { left, right }, queries: { q_right: rightQuery } });
}

test("merge joins on Date keys across DataTable and ArrowTableAdapter sources", async () => {
  const ms = Date.parse("2020-01-01T00:00:00.000Z");

  const leftData = DataTable.fromGrid(
    [
      ["When", "Left"],
      [new Date(ms), "L1"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const rightData = DataTable.fromGrid(
    [
      ["When", "Right"],
      [new Date(ms), "R1"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const leftArrow = new ArrowTableAdapter(
    makeFakeArrowTable({ When: [ms], Left: ["L1"] }, { When: "Date64", Left: "Utf8" }),
    [
      { name: "When", type: "date" },
      { name: "Left", type: "string" },
    ],
  );

  const rightArrow = new ArrowTableAdapter(
    makeFakeArrowTable({ When: [ms], Right: ["R1"] }, { When: "Date64", Right: "Utf8" }),
    [
      { name: "When", type: "date" },
      { name: "Right", type: "string" },
    ],
  );

  const expected = [
    ["When", "Left", "Right"],
    [new Date(ms), "L1", "R1"],
  ];

  const cases = [
    { name: "DataTable x DataTable", left: leftData, right: rightData },
    { name: "DataTable x ArrowTableAdapter", left: leftData, right: rightArrow },
    { name: "ArrowTableAdapter x DataTable", left: leftArrow, right: rightData },
    { name: "ArrowTableAdapter x ArrowTableAdapter", left: leftArrow, right: rightArrow },
  ];

  for (const { name, left, right } of cases) {
    const result = await runMerge({ left, right, leftKey: "When", rightKey: "When" });
    assert.deepEqual(result.toGrid(), expected, name);
  }
});

test("merge joins on null keys (null-safe equality)", async () => {
  const left = DataTable.fromGrid(
    [
      ["Key", "Left"],
      [null, "L1"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const right = DataTable.fromGrid(
    [
      ["Key", "Right"],
      [null, "R1"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const result = await runMerge({ left, right, leftKey: "Key", rightKey: "Key" });
  assert.deepEqual(result.toGrid(), [
    ["Key", "Left", "Right"],
    [null, "L1", "R1"],
  ]);
});

test("merge joins on object keys via stable stringification", async () => {
  const left = DataTable.fromGrid(
    [
      ["Key", "Left"],
      [{ a: 1, b: 2 }, "L1"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  // Same keys/values but created in a different insertion order.
  const rightKey = { b: 2, a: 1 };
  const right = DataTable.fromGrid(
    [
      ["Key", "Right"],
      [rightKey, "R1"],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const result = await runMerge({ left, right, leftKey: "Key", rightKey: "Key" });
  assert.deepEqual(result.toGrid(), [
    ["Key", "Left", "Right"],
    [{ a: 1, b: 2 }, "L1", "R1"],
  ]);
});
