import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, findAll } from "../index.js";

test("search scopes: selection vs sheet vs workbook", async () => {
  const wb = new InMemoryWorkbook();
  const s1 = wb.addSheet("Sheet1");
  const s2 = wb.addSheet("Sheet2");

  // Sheet1
  s1.setValue(0, 0, "foo"); // A1
  s1.setValue(0, 1, "bar"); // B1
  s1.setValue(1, 0, "Foo"); // A2
  s1.setFormula(0, 2, "=CONCAT(\"f\",\"oo\")", { value: "foo", display: "foo" }); // C1

  // Sheet2
  s2.setValue(0, 0, "foo2"); // A1

  const sheetMatches = await findAll(wb, "foo", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "values",
    valueMode: "display",
  });
  assert.deepEqual(
    sheetMatches.map((m) => m.address),
    ["Sheet1!A1", "Sheet1!C1", "Sheet1!A2"],
  );

  const selectionMatches = await findAll(wb, "foo", {
    scope: "selection",
    currentSheetName: "Sheet1",
    selectionRanges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }], // A1:B1
  });
  assert.deepEqual(selectionMatches.map((m) => m.address), ["Sheet1!A1"]);

  const workbookMatches = await findAll(wb, "foo*", {
    scope: "workbook",
    currentSheetName: "Sheet1",
  });
  assert.deepEqual(
    workbookMatches.map((m) => m.address),
    ["Sheet1!A1", "Sheet1!C1", "Sheet1!A2", "Sheet2!A1"],
  );
});

test("search scopes: overlapping selection ranges do not produce duplicate matches", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setValue(0, 0, "foo"); // A1
  sheet.setValue(0, 1, "foo"); // B1

  const matches = await findAll(wb, "foo", {
    scope: "selection",
    currentSheetName: "Sheet1",
    selectionRanges: [
      { startRow: 0, endRow: 0, startCol: 0, endCol: 1 }, // A1:B1
      { startRow: 0, endRow: 0, startCol: 1, endCol: 2 }, // B1:C1 (overlaps on B1)
    ],
  });

  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!A1", "Sheet1!B1"]);
});

test("search order: by rows vs by columns", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setValue(0, 1, "match"); // B1
  sheet.setValue(1, 0, "match"); // A2

  const byRows = await findAll(wb, "match", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    searchOrder: "byRows",
  });
  assert.deepEqual(byRows.map((m) => m.address), ["Sheet1!B1", "Sheet1!A2"]);

  const byColumns = await findAll(wb, "match", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    searchOrder: "byColumns",
  });
  assert.deepEqual(byColumns.map((m) => m.address), ["Sheet1!A2", "Sheet1!B1"]);
});

test("look in: formulas searches formula text", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setFormula(0, 0, "=SUM(1,2,3)", { value: 6, display: "6" });
  sheet.setValue(0, 1, "SUM(1,2,3)"); // literal text

  const matches = await findAll(wb, "SUM", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "formulas",
    matchCase: true,
  });

  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!A1", "Sheet1!B1"]);
});
