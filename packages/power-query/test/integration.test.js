import assert from "node:assert/strict";
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import { QueryEngine } from "../src/engine.js";
import { InMemorySheet, writeTableToSheet } from "../src/sheet.js";

test("integration: load CSV -> filter/sort/groupBy -> write to sheet", async () => {
  const dir = await mkdtemp(path.join(os.tmpdir(), "power-query-"));
  const csvPath = path.join(dir, "sales.csv");
  await writeFile(
    csvPath,
    ["Region,Product,Sales", "East,A,100", "East,B,150", "West,A,200", "West,B,250"].join("\n"),
    "utf8",
  );

  const engine = new QueryEngine({ fileAdapter: { readText: (p) => readFile(p, "utf8") } });
  const query = {
    id: "q_sales",
    name: "Sales",
    source: { type: "csv", path: csvPath, options: { hasHeaders: true } },
    steps: [
      {
        id: "s_filter",
        name: "Filter East",
        operation: { type: "filterRows", predicate: { type: "comparison", column: "Region", operator: "equals", value: "East" } },
      },
      {
        id: "s_group",
        name: "Group by Region",
        operation: { type: "groupBy", groupColumns: ["Region"], aggregations: [{ column: "Sales", op: "sum", as: "Total Sales" }] },
      },
      {
        id: "s_sort",
        name: "Sort Desc",
        operation: { type: "sortRows", sortBy: [{ column: "Total Sales", direction: "descending" }] },
      },
    ],
    refreshPolicy: { type: "manual" },
  };

  const result = await engine.executeQuery(query, {}, {});
  assert.deepEqual(result.toGrid(), [
    ["Region", "Total Sales"],
    ["East", 250],
  ]);

  const sheet = new InMemorySheet();
  writeTableToSheet(result, sheet, { startRow: 1, startCol: 1 });

  assert.equal(sheet.getCell(1, 1), "Region");
  assert.equal(sheet.getCell(1, 2), "Total Sales");
  assert.equal(sheet.getCell(2, 1), "East");
  assert.equal(sheet.getCell(2, 2), 250);
});
