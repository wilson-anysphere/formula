import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

const EXCEL_MAX_COL = 16_384 - 1;

test("setRangeFormat uses compressed range runs for huge rectangles without materializing cells", () => {
  const doc = new DocumentController();

  // 26 columns * 1,000,000 rows = 26,000,000 cells. This must not enumerate per cell.
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });
  const boldStyleId = doc.styleTable.intern({ font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);

  // No cells should be created for format-only range runs.
  assert.equal(sheet.cells.size, 0);

  // Range runs should be stored per-column.
  assert.equal(sheet.formatRunsByCol.size, 26);
  for (let col = 0; col < 26; col += 1) {
    const runs = sheet.formatRunsByCol.get(col);
    assert.ok(Array.isArray(runs));
    assert.equal(runs.length, 1);
    assert.equal(runs[0].startRow, 0);
    assert.equal(runs[0].endRowExclusive, 1_000_000);
    assert.equal(runs[0].styleId, boldStyleId);
  }

  // Effective formatting should apply to empty cells inside the rectangle.
  const inside = doc.getCellFormat("Sheet1", "A1");
  assert.equal(inside.font?.bold, true);

  // Cells outside the rectangle should not have the format.
  const outside = doc.getCellFormat("Sheet1", "AA1"); // column 27 (0-based 26)
  assert.equal(outside.font?.bold, undefined);

  // Styles should be interned per segment, not per cell.
  assert.equal(doc.styleTable.size, 2); // default + bold

  // Style-id tuples should include the range-run layer so caches can key correctly.
  const idsInside = doc.getCellFormatStyleIds("Sheet1", "A1");
  assert.equal(idsInside.length, 5);
  assert.equal(idsInside[3], 0); // cell layer remains empty
  assert.equal(idsInside[4], 1); // range-run layer contributes bold

  // Default used range ignores format-only regions.
  assert.equal(doc.getUsedRange("Sheet1"), null);

  // includeFormat used range should incorporate range-run formatting (without cell materialization).
  assert.deepEqual(doc.getUsedRange("Sheet1", { includeFormat: true }), {
    startRow: 0,
    endRow: 999_999,
    startCol: 0,
    endCol: 25,
  });
});

test("setRangeFormat for full-height columns patches existing range-run formatting", () => {
  const doc = new DocumentController();

  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  // Clear formatting for column A across the full sheet height. This should also clear any
  // existing range-run formatting in that column (range runs are higher precedence than col defaults).
  doc.setRangeFormat("Sheet1", "A1:A1048576", null);

  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, undefined);
  assert.equal(doc.getCellFormat("Sheet1", "B1").font?.bold, true);

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.formatRunsByCol.has(0), false);
});

test("setRangeFormat clearing a single cell removes underlying range-run formatting", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:A10", { font: { bold: true } });

  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);
  doc.setRangeFormat("Sheet1", "A1", null);

  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, undefined);
  assert.equal(doc.getCellFormat("Sheet1", "A2").font?.bold, true);
});

test("range-run formatting is undoable + redoable", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:Z1000000", { font: { bold: true } });

  const sheet = doc.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.formatRunsByCol.size, 26);

  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);
  doc.undo();
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, undefined);
  assert.equal(sheet.formatRunsByCol.size, 0);
  doc.redo();
  assert.equal(doc.getCellFormat("Sheet1", "A1").font?.bold, true);
  assert.equal(sheet.formatRunsByCol.size, 26);
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
  const sheetJson = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.ok(sheetJson);
  assert.ok(Array.isArray(sheetJson.formatRunsByCol));
  assert.equal(sheetJson.formatRunsByCol.length, 26);
  assert.deepEqual(sheetJson.formatRunsByCol[0], {
    col: 0,
    runs: [{ startRow: 0, endRowExclusive: 1_000_000, format: { font: { bold: true } } }],
  });

  const restored = new DocumentController();
  restored.applyState(snapshot);

  assert.equal(restored.getCellFormat("Sheet1", "A1").font?.bold, true);
  const sheet = restored.model.sheets.get("Sheet1");
  assert.ok(sheet);
  assert.equal(sheet.formatRunsByCol.size, 26);
});

test("range-run formatting can be overridden by later full-width row formatting (row formatting patches runs)", () => {
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
