import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

const EXCEL_MAX_COL = 16_384 - 1;

test("setRangeFormat uses compressed format runs for huge rectangles (no per-cell materialization)", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });
  const boldStyleId = doc.styleTable.intern({ font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0);

  // 26 columns (A-Z), one interval per column.
  assert.equal(sheet.formatRunsByCol.size, 26);
  for (let col = 0; col < 26; col++) {
    const runs = sheet.formatRunsByCol.get(col);
    assert.ok(Array.isArray(runs));
    assert.equal(runs.length, 1);
    assert.equal(runs[0].startRow, 0);
    assert.equal(runs[0].endRowExclusive, 1_000_000);
    assert.equal(runs[0].styleId, boldStyleId);
  }

  assert.deepEqual(doc.getUsedRange("Sheet1"), null);
  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 999_999,
    startCol: 0,
    endCol: 25,
  });
});

test("getCellFormat incorporates run formatting even when cell.styleId === 0", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  const cell = doc.getCell("Sheet1", "M500000");
  assert.equal(cell.styleId, 0);
  const format = doc.getCellFormat("Sheet1", "M500000");
  assert.equal(format.font?.bold, true);
});

test("encodeState/applyState roundtrip preserves range-run formatting", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  const snapshot = doc.encodeState();
  const decoded =
    typeof TextDecoder !== "undefined"
      ? new TextDecoder().decode(snapshot)
      : // eslint-disable-next-line no-undef
        Buffer.from(snapshot).toString("utf8");
  const parsed = JSON.parse(decoded);
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.ok(sheet);
  assert.ok(Array.isArray(sheet.formatRunsByCol));
  assert.equal(sheet.formatRunsByCol.length, 26);
  assert.deepEqual(sheet.formatRunsByCol[0], {
    col: 0,
    runs: [{ startRow: 0, endRowExclusive: 1_000_000, format: { font: { bold: true } } }],
  });

  const restored = new DocumentController();
  restored.applyState(snapshot);
  assert.equal(restored.getCell("Sheet1", "A1").styleId, 0);
  assert.deepEqual(restored.getCellFormat("Sheet1", "A1"), { font: { bold: true } });
  assert.equal(restored.model.sheets.get("Sheet1").formatRunsByCol.size, 26);
});

test("undo/redo restores format runs precisely", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);

  doc.undo();
  assert.deepEqual(doc.getCellFormat("Sheet1", "A1"), {});
  const sheetAfterUndo = doc.model.sheets.get("Sheet1");
  assert.ok(sheetAfterUndo);
  assert.equal(sheetAfterUndo.formatRunsByCol.size, 0);

  doc.redo();
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);
  const sheetAfterRedo = doc.model.sheets.get("Sheet1");
  assert.ok(sheetAfterRedo);
  assert.equal(sheetAfterRedo.formatRunsByCol.size, 26);
});

test("overlapping formatting patches override deterministically (row formatting patches runs)", () => {
  const doc = new DocumentController();

  // Seed a range-run format in A1:A60000 (exceeds range-run threshold so we don't enumerate cells).
  doc.setRangeFormat("Sheet1", "A1:A60000", { fill: { fgColor: "red" } });
  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.cells.size, 0);
  assert.equal(sheet.formatRunsByCol.size, 1);
  assert.equal(doc.getCellFormat("Sheet1", "A2").fill?.fgColor, "red");

  // Apply a full-width row format to row 2 (0-based row=1) that conflicts with the run.
  doc.setRangeFormat(
    "Sheet1",
    { start: { row: 1, col: 0 }, end: { row: 1, col: EXCEL_MAX_COL } },
    { fill: { fgColor: "blue" } },
  );

  assert.equal(doc.getCellFormat("Sheet1", "A1").fill?.fgColor, "red");
  assert.equal(doc.getCellFormat("Sheet1", "A2").fill?.fgColor, "blue");
  assert.equal(doc.getCellFormat("Sheet1", "A3").fill?.fgColor, "red");
});
