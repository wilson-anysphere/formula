import assert from "node:assert/strict";
import test from "node:test";

import { ArrowTableAdapter } from "../src/arrowTable.js";
import { DataTable } from "../src/table.js";
import { applyOperation } from "../src/steps.js";

/**
 * @param {unknown[][]} grid
 */
function gridDatesToIso(grid) {
  return grid.map((row) => row.map((cell) => (cell instanceof Date ? cell.toISOString() : cell)));
}

/**
 * Build a minimal Arrow-like table object sufficient for `ArrowTableAdapter`.
 *
 * This keeps the test independent of the optional `apache-arrow` dependency by
 * avoiding `arrowTableFromColumns`.
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

test("distinctRows matches across DataTable and Arrow backends (Dates)", () => {
  const d1 = new Date("2020-01-01T00:00:00.000Z");
  const d1b = new Date("2020-01-01T00:00:00.000Z");
  const d2 = new Date("2020-01-02T00:00:00.000Z");
  const d2b = new Date("2020-01-02T00:00:00.000Z");

  const dataTable = DataTable.fromGrid(
    [
      ["Id", "When"],
      [1, d1],
      [1, d1b],
      [1, d2],
      [1, d2b],
      [2, d1b],
      [2, d1],
    ],
    { hasHeaders: true, inferTypes: true },
  );

  const ms1 = d1.getTime();
  const ms2 = d2.getTime();
  const arrowTable = new ArrowTableAdapter(
    makeFakeArrowTable(
      {
        Id: [1, 1, 1, 1, 2, 2],
        When: [ms1, ms1, ms2, ms2, ms1, ms1],
      },
      { Id: "Int32", When: "Date64" },
    ),
  );

  const op = { type: "distinctRows", columns: null };
  const expectedIso = [
    ["Id", "When"],
    [1, d1.toISOString()],
    [1, d2.toISOString()],
    [2, d1.toISOString()],
  ];

  assert.deepEqual(gridDatesToIso(applyOperation(dataTable, op).toGrid()), expectedIso);
  assert.deepEqual(gridDatesToIso(applyOperation(arrowTable, op).toGrid()), expectedIso);
});
