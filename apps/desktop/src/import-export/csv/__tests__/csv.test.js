import test from "node:test";
import assert from "node:assert/strict";

import { exportCellGridToCsv } from "../export.js";
import { parseCsv } from "../csv.js";
import { importCsvToCellGrid } from "../import.js";

test("CSV parses quoted fields, quotes, and embedded newlines", () => {
  const rows = parseCsv('name,notes\nAlice,"Line1\nLine2"\nBob,"He said ""hi"""', {
    delimiter: ",",
  });

  assert.deepEqual(rows, [
    ["name", "notes"],
    ["Alice", "Line1\nLine2"],
    ["Bob", 'He said "hi"'],
  ]);
});

test("CSV import infers column types and preserves header strings", () => {
  const csv = "id,amount,active,date\n001,10,true,2024-01-31\n002,20,false,2024-02-01\n";
  const { grid } = importCsvToCellGrid(csv, { delimiter: "," });

  // Header row stays strings.
  assert.equal(grid[0][0].value, "id");
  assert.equal(grid[0][1].value, "amount");

  // Inferred typing for subsequent rows.
  assert.equal(grid[1][0].value, "001"); // leading zeros preserved
  assert.equal(grid[1][1].value, 10);
  assert.equal(grid[1][2].value, true);
  assert.equal(typeof grid[1][3].value, "number");
  assert.equal(grid[1][3].format.numberFormat, "yyyy-mm-dd");
});

test("CSV export quotes fields when needed", () => {
  const csv = exportCellGridToCsv(
    [
      [{ value: "a" }, { value: "b,c" }],
      [{ value: 1 }, { value: true }],
    ],
    { delimiter: "," }
  );

  assert.equal(csv, 'a,"b,c"\r\n1,TRUE');
});
