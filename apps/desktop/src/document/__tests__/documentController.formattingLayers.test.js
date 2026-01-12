import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("layered formatting composes non-conflicting keys (col bold + row italic)", () => {
  const doc = new DocumentController();

  // Column A (0) sets bold, Row 1 (0) sets italic.
  doc.setColFormat("Sheet1", 0, { font: { bold: true } });
  doc.setRowFormat("Sheet1", 0, { font: { italic: true } });

  const format = doc.getCellFormat("Sheet1", "A1");
  assert.equal(format.font?.bold, true);
  assert.equal(format.font?.italic, true);
});

test("layered formatting resolves conflicts deterministically (row overrides col)", () => {
  const doc = new DocumentController();

  doc.setColFormat("Sheet1", 0, { fill: { fgColor: "red" } });
  doc.setRowFormat("Sheet1", 0, { fill: { fgColor: "blue" } });

  const format = doc.getCellFormat("Sheet1", "A1");
  assert.equal(format.fill?.fgColor, "blue");
});

test("cell formatting overrides row/col formatting", () => {
  const doc = new DocumentController();

  doc.setColFormat("Sheet1", 0, { fill: { fgColor: "red" } });
  doc.setRowFormat("Sheet1", 0, { fill: { fgColor: "blue" } });
  doc.setRangeFormat("Sheet1", "A1", { fill: { fgColor: "green" } });

  const format = doc.getCellFormat("Sheet1", "A1");
  assert.equal(format.fill?.fgColor, "green");
});

