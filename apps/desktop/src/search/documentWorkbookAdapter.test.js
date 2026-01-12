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

test("DocumentWorkbookAdapter resolves sheets by display name via sheetNameResolver", () => {
  const doc = { getSheetIds: () => ["Sheet1", "Sheet2"] };
  const namesById = new Map([
    ["Sheet1", "Sheet1"],
    ["Sheet2", "My Sheet"],
  ]);
  const sheetNameResolver = {
    getSheetNameById: (id) => namesById.get(id) ?? null,
    getSheetIdByName: (name) => {
      const needle = String(name ?? "").trim().toLowerCase();
      if (!needle) return null;
      for (const [id, sheetName] of namesById.entries()) {
        if (sheetName.toLowerCase() === needle) return id;
      }
      return null;
    },
  };

  const workbook = new DocumentWorkbookAdapter({ document: doc, sheetNameResolver });

  const sheetNames = workbook.sheets.map((s) => s.name);
  assert.deepEqual(sheetNames, ["Sheet1", "My Sheet"]);

  const sheet = workbook.getSheet("My Sheet");
  assert.equal(sheet.sheetId, "Sheet2");

  const sheetLower = workbook.getSheet("my sheet");
  assert.equal(sheetLower.sheetId, "Sheet2");

  // Accept Excel-style quoting for sheet tokens (e.g. from sheet-qualified references).
  const sheetQuoted = workbook.getSheet("'My Sheet'");
  assert.equal(sheetQuoted.sheetId, "Sheet2");

  assert.throws(() => workbook.getSheet("MissingSheet"), /Unknown sheet/i);
});
