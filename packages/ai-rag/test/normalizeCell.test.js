import assert from "node:assert/strict";
import test from "node:test";

import { chunkToText } from "../src/workbook/chunkToText.js";
import { normalizeCell } from "../src/workbook/normalizeCell.js";

test("normalizeCell canonicalizes formula text", () => {
  assert.deepEqual(normalizeCell({ value: null, formula: "  =  SUM(A1:A2)  " }), { f: "=SUM(A1:A2)" });
  assert.deepEqual(normalizeCell({ value: null, formula: "SUM(A1:A2)" }), { f: "=SUM(A1:A2)" });
  assert.deepEqual(normalizeCell({ value: null, formula: "=" }), {});
  assert.deepEqual(normalizeCell("   =   "), {});
});

test("chunkToText does not double-prefix formulas with '='", () => {
  const text = chunkToText(
    {
      id: "wb1::Sheet1::dataRegion::Test",
      workbookId: "wb1",
      sheetName: "Sheet1",
      kind: "dataRegion",
      title: "Test",
      rect: { r0: 0, c0: 0, r1: 0, c1: 0 },
      cells: [[{ f: "=A1*2" }]],
    },
    { sampleRows: 1 },
  );

  assert.match(text, /=A1\*2/);
  assert.doesNotMatch(text, /==A1\*2/);
});

