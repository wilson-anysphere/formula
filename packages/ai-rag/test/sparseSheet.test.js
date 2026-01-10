import assert from "node:assert/strict";
import test from "node:test";

import { chunkWorkbook } from "../src/workbook/chunkWorkbook.js";
import { chunkToText } from "../src/workbook/chunkToText.js";

test("chunkWorkbook supports sparse sheet cell maps (Map row,col -> cell)", () => {
  const cells = new Map();
  cells.set("0,0", { value: "Region" });
  cells.set("0,1", { value: "Revenue" });
  cells.set("1,0", { value: "North" });
  cells.set("1,1", { value: 100 });
  cells.set("2,0", { value: "South" });
  cells.set("2,1", { value: 200 });
  cells.set("0,3", { formula: "SUM(B2:B3)" });
  cells.set("1,3", { formula: "B2*2" });

  const workbook = {
    id: "wb-sparse",
    sheets: [{ name: "Sheet1", cells }],
    tables: [{ name: "RevenueByRegion", sheetName: "Sheet1", rect: { r0: 0, c0: 0, r1: 2, c1: 1 } }],
  };

  const chunks = chunkWorkbook(workbook);
  const table = chunks.find((c) => c.kind === "table");
  assert.ok(table);
  assert.equal(table.title, "RevenueByRegion");

  const text = chunkToText(table, { sampleRows: 2 });
  assert.match(text, /Region/);
  assert.match(text, /Revenue/);

  const formulaRegions = chunks.filter((c) => c.kind === "formulaRegion");
  assert.ok(formulaRegions.length >= 1);
});
