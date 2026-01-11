import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, replaceAll, replaceNext, SearchSession, WorkbookSearchIndex } from "../index.js";

test("replace all: values mode overwrites formulas (Excel semantics)", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setFormula(0, 0, "=CONCAT(\"f\",\"oo\")", { value: "foo", display: "foo" }); // A1
  sheet.setValue(0, 1, "foo"); // B1

  const res = await replaceAll(wb, "foo", "bar", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "values",
    valueMode: "display",
  });

  assert.equal(res.replacedCells, 2);

  const a1 = sheet.getCell(0, 0);
  assert.equal(a1.formula, undefined);
  assert.equal(a1.value, "bar");

  const b1 = sheet.getCell(0, 1);
  assert.equal(b1.value, "bar");
});

test("replace all: formulas mode edits formula text without converting to values", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");

  sheet.setFormula(0, 0, "=SUM(1,2,3)", { value: 6, display: "6" });

  await replaceAll(wb, "SUM", "AVERAGE", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    lookIn: "formulas",
    matchCase: true,
  });

  const a1 = sheet.getCell(0, 0);
  assert.equal(a1.formula, "=AVERAGE(1,2,3)");
  assert.equal(a1.value, 6);
});

test("replace: numeric coercion keeps numbers when possible", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setValue(0, 0, 123);

  await replaceAll(wb, "2", "9", { scope: "sheet", currentSheetName: "Sheet1", matchCase: true });
  assert.equal(sheet.getCell(0, 0).value, 193);
});

test("replace next replaces one occurrence; replace all replaces all occurrences in a cell", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setValue(0, 0, "foofoo");

  const r1 = await replaceNext(
    wb,
    "foo",
    "x",
    { scope: "sheet", currentSheetName: "Sheet1" },
    { sheetName: "Sheet1", row: 0, col: 0 },
  );
  assert.equal(r1.replaced, true);
  assert.equal(sheet.getCell(0, 0).value, "xfoo");

  await replaceAll(wb, "foo", "x", { scope: "sheet", currentSheetName: "Sheet1" });
  assert.equal(sheet.getCell(0, 0).value, "xx");
});

test("replaceAll updates an attached WorkbookSearchIndex", async () => {
  const wb = new InMemoryWorkbook();
  const sheet = wb.addSheet("Sheet1");
  sheet.setValue(0, 0, "foo");

  const index = new WorkbookSearchIndex(wb, { autoThresholdCells: 0 });

  const builder = new SearchSession(wb, "foo", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  await builder.findNext();

  await replaceAll(wb, "foo", "bar", { scope: "sheet", currentSheetName: "Sheet1", index });

  const session = new SearchSession(wb, "bar", {
    scope: "sheet",
    currentSheetName: "Sheet1",
    index,
    indexStrategy: "always",
  });
  const match = await session.findNext();
  assert.equal(match.address, "Sheet1!A1");
  assert.equal(session.stats.indexCellsVisited, 0);
});
