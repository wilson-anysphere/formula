import assert from "node:assert/strict";
import test from "node:test";

import { arrowTableFromColumns } from "../../data-io/src/index.js";

import { ArrowTableAdapter } from "../src/arrowTable.js";
import { DataTable } from "../src/table.js";
import { applyOperation } from "../src/steps.js";

/**
 * @param {unknown[][]} grid
 */
function gridDatesToIso(grid) {
  return grid.map((row) => row.map((cell) => (cell instanceof Date ? cell.toISOString() : cell)));
}

test("distinctRows matches across DataTable and Arrow backends (Dates)", (t) => {
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

  let arrowTable;
  try {
    arrowTable = new ArrowTableAdapter(
      arrowTableFromColumns({
        Id: [1, 1, 1, 1, 2, 2],
        When: [d1, d1b, d2, d2b, d1b, d1],
      }),
    );
  } catch (err) {
    // `apache-arrow` is an optional dependency in some environments (e.g. agent sandboxes).
    // When it isn't installed, Arrow-backed tests should be skipped instead of failing.
    const message = err instanceof Error ? err.message : String(err);
    if (message.includes("optional 'apache-arrow'")) {
      t.skip(message);
      return;
    }
    throw err;
  }

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
