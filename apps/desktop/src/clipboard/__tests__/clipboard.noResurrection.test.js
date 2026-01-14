import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { getCellGridFromRange } from "../clipboard.js";

test("clipboard range reads do not resurrect deleted sheets", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellValue("Sheet2", "A1", 2);
  doc.deleteSheet("Sheet2");

  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);
  assert.equal(doc.getSheetMeta("Sheet2"), null);

  const grid = getCellGridFromRange(doc, "Sheet2", "A1:B2");
  assert.equal(grid.length, 2);
  assert.equal(grid[0]?.length, 2);

  // Must not recreate the deleted sheet.
  assert.deepEqual(doc.getSheetIds(), ["Sheet1"]);
  assert.equal(doc.getSheetMeta("Sheet2"), null);
});

