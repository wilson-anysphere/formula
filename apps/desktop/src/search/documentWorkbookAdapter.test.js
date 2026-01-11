import assert from "node:assert/strict";
import test from "node:test";

import { DocumentWorkbookAdapter } from "./documentWorkbookAdapter.js";

test("DocumentWorkbookAdapter preserves defined-name case while remaining case-insensitive", () => {
  const workbook = new DocumentWorkbookAdapter({ document: { getSheetIds: () => [] } });
  const range = { startRow: 0, endRow: 9, startCol: 0, endCol: 0 };

  workbook.defineName("SalesData", { sheetName: "Sheet1", range });

  const storedLower = workbook.getName("salesdata");
  assert.ok(storedLower);
  assert.equal(storedLower.name, "SalesData");
  assert.deepEqual(storedLower.range, range);

  const storedUpper = workbook.getName("SALESDATA");
  assert.ok(storedUpper);
  assert.equal(storedUpper.name, "SalesData");
  assert.deepEqual(storedUpper.range, range);
});

test("DocumentWorkbookAdapter schemaVersion bumps when schema changes", () => {
  const workbook = new DocumentWorkbookAdapter({ document: { getSheetIds: () => [] } });
  assert.equal(workbook.schemaVersion, 0);

  workbook.defineName("Name1", { sheetName: "Sheet1", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
  assert.equal(workbook.schemaVersion, 1);

  workbook.addTable({ name: "Table1", sheetName: "Sheet1", startRow: 0, endRow: 1, startCol: 0, endCol: 1, columns: ["A"] });
  assert.equal(workbook.schemaVersion, 2);

  workbook.clearSchema();
  assert.equal(workbook.schemaVersion, 3);
  assert.equal(workbook.names.size, 0);
  assert.equal(workbook.tables.size, 0);
});

