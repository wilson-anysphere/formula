import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, findAll } from "../index.js";

test("merged cells: selection scope treats merged region as the top-left cell", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  // Merge A1:B2 and put the value in A1 (Excel semantics).
  sheet.mergeCells({ startRow: 0, endRow: 1, startCol: 0, endCol: 1 });
  sheet.setValue(0, 0, "foo");

  // Select B2 only – Excel still searches the merged cell (A1).
  const matches = await findAll(wb, "foo", {
    scope: "selection",
    currentSheetName: "Sheet1",
    selectionRanges: [{ startRow: 1, endRow: 1, startCol: 1, endCol: 1 }],
  });

  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!A1"]);
});

test("valueMode: display vs raw is respected for typed values", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setCell(0, 0, { value: { t: "n", v: 1.234 }, display: "1.23" });

  const displayMatches = await findAll(wb, "1.23", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "values",
    valueMode: "display",
    matchEntireCell: true,
  });
  assert.deepEqual(displayMatches.map((m) => m.address), ["Sheet1!A1"]);

  const rawMatches = await findAll(wb, "1.23", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "values",
    valueMode: "raw",
    matchEntireCell: true,
  });
  assert.deepEqual(rawMatches.map((m) => m.address), []);
});

test("error cells: display/raw formatting matches #DIV/0! string", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setCell(0, 1, { value: { t: "e", v: "#DIV/0!" } });

  const matches = await findAll(wb, "#DIV/0!", { scope: "sheet", currentSheetName: "Sheet1" });
  assert.deepEqual(matches.map((m) => m.address), ["Sheet1!B1"]);
  assert.equal(matches[0].text, "#DIV/0!");
});

test("blank typed values stringify as empty string (no [object Object])", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  // Formula cells can legitimately have a blank value.
  sheet.setFormula(0, 0, "=IF(1=2,\"x\",\"\")", { value: { t: "blank" } });

  const matches = await findAll(wb, "*", {
    scope: "selection",
    currentSheetName: "Sheet1",
    selectionRanges: [{ startRow: 0, endRow: 0, startCol: 0, endCol: 0 }],
    matchEntireCell: true,
  });

  assert.equal(matches.length, 1);
  assert.equal(matches[0].address, "Sheet1!A1");
  assert.equal(matches[0].text, "");
});

