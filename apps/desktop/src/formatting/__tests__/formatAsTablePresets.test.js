import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../document/documentController.js";
import { applyFormatAsTablePreset, FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS, getFormatAsTablePreset } from "../formatAsTablePresets.js";

test("applyFormatAsTablePreset applies header formatting, banded rows, and outline borders", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [
    ["Name", "Value"],
    ["A", 1],
    ["B", 2],
  ]);

  const preset = getFormatAsTablePreset("light");
  const historyBefore = doc.history.length;
  const ok = applyFormatAsTablePreset(doc, "Sheet1", { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }, "light");
  assert.equal(ok, true);
  assert.equal(doc.history.length, historyBefore + 1);

  const headerA1 = doc.getCellFormat("Sheet1", { row: 0, col: 0 });
  const headerB1 = doc.getCellFormat("Sheet1", { row: 0, col: 1 });
  assert.equal(Boolean(headerA1.font?.bold), true);
  assert.equal(Boolean(headerB1.font?.bold), true);
  assert.equal(headerA1.fill?.fgColor, `#${preset.header.fill}`);
  assert.equal(headerA1.font?.color, `#${preset.header.fontColor}`);

  const bodyA2 = doc.getCellFormat("Sheet1", { row: 1, col: 0 });
  const bodyA3 = doc.getCellFormat("Sheet1", { row: 2, col: 0 });
  assert.equal(bodyA2.fill?.fgColor, `#${preset.bandedRows.primaryFill}`);
  assert.equal(bodyA3.fill?.fgColor, `#${preset.bandedRows.secondaryFill}`);

  // Outline borders.
  assert.equal(headerA1.border?.top?.style, preset.borders.style);
  assert.equal(headerA1.border?.left?.style, preset.borders.style);
  assert.equal(headerA1.border?.top?.color, `#${preset.borders.outlineColor}`);
  assert.equal(headerA1.border?.left?.color, `#${preset.borders.outlineColor}`);
  // Inner horizontal separators.
  assert.equal(headerA1.border?.bottom?.style, preset.borders.style);
  assert.equal(headerA1.border?.bottom?.color, `#${preset.borders.innerHorizontalColor}`);
  // Inner vertical separators.
  assert.equal(headerA1.border?.right?.style, preset.borders.style);
  assert.equal(headerA1.border?.right?.color, `#${preset.borders.innerHorizontalColor}`);
  assert.equal(bodyA2.border?.right?.style, preset.borders.style);
  assert.equal(bodyA2.border?.right?.color, `#${preset.borders.innerHorizontalColor}`);

  const bottomRight = doc.getCellFormat("Sheet1", { row: 2, col: 1 });
  assert.equal(bottomRight.border?.bottom?.style, preset.borders.style);
  assert.equal(bottomRight.border?.right?.style, preset.borders.style);
  assert.equal(bottomRight.border?.bottom?.color, `#${preset.borders.outlineColor}`);
  assert.equal(bottomRight.border?.right?.color, `#${preset.borders.outlineColor}`);
});

test("applyFormatAsTablePreset refuses overly large ranges (banding cap)", () => {
  const doc = new DocumentController();
  // Materialize the sheet so the test only measures the formatting call's history impact.
  doc.setCellValue("Sheet1", "A1", "x");

  const before = doc.history.length;
  const depthBefore = doc.batchDepth;
  // rowCount = 2*max + 3 => floor((rows - 1) / 2) = max + 1 banded-row ops (exceeds the cap).
  const rowCount = FORMAT_AS_TABLE_MAX_BANDED_ROW_OPS * 2 + 3;
  const ok = applyFormatAsTablePreset(
    doc,
    "Sheet1",
    { start: { row: 0, col: 0 }, end: { row: rowCount - 1, col: 0 } },
    "light",
  );
  assert.equal(ok, false);
  assert.equal(doc.history.length, before);
  assert.equal(doc.batchDepth, depthBefore);
});

test("applyFormatAsTablePreset does not start a nested batch when already batching", () => {
  const doc = new DocumentController();
  doc.setRangeValues("Sheet1", "A1", [
    ["Name", "Value"],
    ["A", 1],
    ["B", 2],
  ]);

  const historyBefore = doc.history.length;
  doc.beginBatch({ label: "Outer batch" });
  assert.equal(doc.batchDepth, 1);

  const ok = applyFormatAsTablePreset(doc, "Sheet1", { start: { row: 0, col: 0 }, end: { row: 2, col: 1 } }, "light");
  assert.equal(ok, true);
  // Helper should not create a nested batch.
  assert.equal(doc.batchDepth, 1);

  doc.endBatch();
  assert.equal(doc.batchDepth, 0);
  assert.equal(doc.history.length, historyBefore + 1);
});
