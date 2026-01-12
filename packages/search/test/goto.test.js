import test from "node:test";
import assert from "node:assert/strict";

import { InMemoryWorkbook, parseGoTo } from "../index.js";

test("parseGoTo canonicalizes sheet names via workbook.getSheet when available", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  // InMemoryWorkbook resolves sheets case-insensitively; parseGoTo should return the
  // canonical sheet name (as stored on the sheet object).
  const parsed = parseGoTo("sheet1!A1", { workbook: wb, currentSheetName: "Sheet1" });
  assert.equal(parsed.sheetName, "Sheet1");
  assert.deepEqual(parsed.range, { startRow: 0, endRow: 0, startCol: 0, endCol: 0 });
});

test("parseGoTo throws for unknown sheet-qualified references when workbook.getSheet is available", () => {
  const wb = new InMemoryWorkbook();
  wb.addSheet("Sheet1");

  assert.throws(() => parseGoTo("NoSuchSheet!A1", { workbook: wb, currentSheetName: "Sheet1" }), /Unknown sheet/i);
});

