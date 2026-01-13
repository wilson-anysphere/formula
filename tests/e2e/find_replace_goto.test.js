import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, parseGoTo } from "../../packages/search/index.js";
import { FindReplaceController } from "../../apps/desktop/src/panels/find-replace/findReplaceController.js";

test("e2e flow: find next, go to, replace all", async () => {
  const wb = new InMemoryWorkbook();
  const s1 = wb.addSheet("Sheet1");
  const s2 = wb.addSheet("Sheet2");

  s1.setValue(0, 0, "alpha"); // A1
  s1.setValue(0, 1, "alpha"); // B1
  s1.setValue(1, 0, "beta"); // A2
  s2.setValue(0, 0, "alpha"); // Sheet2!A1

  wb.defineName("MyRange", { sheetName: "Sheet2", range: { startRow: 0, endRow: 0, startCol: 0, endCol: 0 } });
  wb.addTable({
    name: "Table1",
    sheetName: "Sheet1",
    startRow: 0,
    endRow: 1,
    startCol: 0,
    endCol: 1,
    columns: ["ColA", "ColB"],
  });

  let activeCell = { sheetName: "Sheet1", row: 0, col: 0 };
  let selectedRange = null;

  function submitGoTo(text) {
    const parsed = parseGoTo(text, { workbook: wb, currentSheetName: activeCell.sheetName });
    if (parsed.type !== "range") return;
    const { range } = parsed;
    if (range.startRow === range.endRow && range.startCol === range.endCol) {
      activeCell = { sheetName: parsed.sheetName, row: range.startRow, col: range.startCol };
    } else {
      selectedRange = { sheetName: parsed.sheetName, range };
    }
  }

  const controller = new FindReplaceController({
    workbook: wb,
    getCurrentSheetName: () => activeCell.sheetName,
    getActiveCell: () => activeCell,
    setActiveCell: (next) => {
      activeCell = next;
    },
    getSelectionRanges: () => [{ startRow: 0, endRow: 0, startCol: 0, endCol: 1 }],
  });

  controller.scope = "workbook";
  controller.query = "alpha";

  // Find next should move from A1 to B1.
  const m1 = await controller.findNext();
  assert.equal(m1.address, "Sheet1!B1");
  assert.deepEqual(activeCell, { sheetName: "Sheet1", row: 0, col: 1 });

  // Go to a sheet-qualified ref
  submitGoTo("Sheet2!A1");
  assert.deepEqual(activeCell, { sheetName: "Sheet2", row: 0, col: 0 });

  // Go to a named range
  submitGoTo("MyRange");
  assert.deepEqual(activeCell, { sheetName: "Sheet2", row: 0, col: 0 });

  // Go to a table structured ref (column)
  submitGoTo("Table1[ColB]");
  assert.deepEqual(selectedRange, {
    sheetName: "Sheet1",
    range: { startRow: 0, endRow: 1, startCol: 1, endCol: 1 },
  });

  // Replace all
  controller.replacement = "ALPHA";
  await controller.replaceAll();

  assert.equal(s1.getCell(0, 0).value, "ALPHA");
  assert.equal(s1.getCell(0, 1).value, "ALPHA");
  assert.equal(s2.getCell(0, 0).value, "ALPHA");
  assert.equal(s1.getCell(1, 0).value, "beta");
});
