import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("encodeState/applyState roundtrip restores cell inputs and clears history", () => {
  const doc = new DocumentController();

  doc.setCellValue("Sheet1", "A1", 1);
  doc.setCellFormula("Sheet1", "B1", "SUM(A1:A3)");
  doc.setRangeFormat("Sheet1", "A1", { bold: true });
  doc.setFrozen("Sheet1", 2, 1);
  doc.setColWidth("Sheet1", 0, 120);
  doc.setRowHeight("Sheet1", 1, 40);
  assert.equal(doc.canUndo, true);

  const snapshot = doc.encodeState();
  assert.ok(snapshot instanceof Uint8Array);

  const decoded =
    typeof TextDecoder !== "undefined"
      ? new TextDecoder().decode(snapshot)
      : // eslint-disable-next-line no-undef
        Buffer.from(snapshot).toString("utf8");
  const parsed = JSON.parse(decoded);
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.equal(sheet.defaultFormat, null);
  assert.deepEqual(sheet.rowFormats, []);
  assert.deepEqual(sheet.colFormats, []);

  const restored = new DocumentController();
  let lastChange = null;
  restored.on("change", (payload) => {
    lastChange = payload;
  });
  restored.applyState(snapshot);

  // applyState clears history and marks dirty until the host explicitly marks saved.
  assert.equal(restored.canUndo, false);
  assert.equal(restored.canRedo, false);
  assert.equal(restored.isDirty, true);

  assert.equal(lastChange?.source, "applyState");
  assert.equal(restored.getCell("Sheet1", "A1").value, 1);
  const a1 = restored.getCell("Sheet1", "A1");
  assert.deepEqual(restored.styleTable.get(a1.styleId), { bold: true });
  assert.equal(restored.getCell("Sheet1", "B1").formula, "=SUM(A1:A3)");
  assert.deepEqual(restored.getSheetView("Sheet1"), {
    frozenRows: 2,
    frozenCols: 1,
    colWidths: { "0": 120 },
    rowHeights: { "1": 40 },
  });
});

test("encodeState/applyState roundtrip preserves layered column formatting", () => {
  const doc = new DocumentController();
  doc.setRangeFormat("Sheet1", "A1:A1048576", { font: { bold: true } });

  const snapshot = doc.encodeState();
  const decoded =
    typeof TextDecoder !== "undefined"
      ? new TextDecoder().decode(snapshot)
      : // eslint-disable-next-line no-undef
        Buffer.from(snapshot).toString("utf8");
  const parsed = JSON.parse(decoded);
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.deepEqual(sheet.defaultFormat, null);
  assert.deepEqual(sheet.rowFormats, []);
  assert.deepEqual(sheet.colFormats, [{ col: 0, format: { font: { bold: true } } }]);

  const restored = new DocumentController();
  restored.applyState(snapshot);
  assert.deepEqual(restored.getCellFormat("Sheet1", "A1048576"), { font: { bold: true } });
  assert.equal(restored.model.sheets.get("Sheet1").cells.size, 0);
});

test("applyState materializes empty sheets from snapshots", () => {
  const doc = new DocumentController();
  // DocumentController lazily creates sheets on first access. This creates an empty sheet
  // that should still survive encode/apply roundtrips.
  doc.getCell("EmptySheet", "A1");

  const snapshot = doc.encodeState();
  const restored = new DocumentController();
  restored.applyState(snapshot);

  assert.deepEqual(restored.getSheetIds(), ["EmptySheet"]);
});

test("applyState removes sheets that are not present in the snapshot", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  doc.getCell("ExtraSheet", "A1"); // empty sheet

  const next = new DocumentController();
  next.setCellValue("OnlySheet", "A1", 2);

  doc.applyState(next.encodeState());

  assert.deepEqual(doc.getSheetIds(), ["OnlySheet"]);
});

test("contentVersion increments when applyState adds/removes sheets (even if empty)", () => {
  const doc = new DocumentController();
  doc.setCellValue("Sheet1", "A1", 1);
  assert.equal(doc.contentVersion, 1);
  assert.equal(doc.updateVersion, 1);

  // Add an *empty* sheet via applyState: no cell deltas, but sheet structure changes.
  const withExtraSheet = new DocumentController();
  withExtraSheet.setCellValue("Sheet1", "A1", 1);
  withExtraSheet.getCell("Sheet2", "A1"); // materialize Sheet2 without content

  doc.applyState(withExtraSheet.encodeState());
  assert.equal(doc.contentVersion, 2);
  assert.equal(doc.updateVersion, 2);
  assert.deepEqual(doc.getSheetIds().sort(), ["Sheet1", "Sheet2"]);

  // Remove the sheet via applyState: again, structure-only change.
  const withoutExtraSheet = new DocumentController();
  withoutExtraSheet.setCellValue("Sheet1", "A1", 1);

  doc.applyState(withoutExtraSheet.encodeState());
  assert.equal(doc.contentVersion, 3);
  assert.equal(doc.updateVersion, 3);
  assert.deepEqual(doc.getSheetIds().sort(), ["Sheet1"]);
});

test("encodeState/applyState preserves sheet insertion order", () => {
  const doc = new DocumentController();
  // Sheets are created lazily on first access. Create them in a non-sorted order.
  doc.getCell("Sheet2", "A1");
  doc.getCell("Sheet1", "A1");
  assert.deepEqual(doc.getSheetIds(), ["Sheet2", "Sheet1"]);

  const snapshot = doc.encodeState();

  const restored = new DocumentController();
  restored.applyState(snapshot);
  assert.deepEqual(restored.getSheetIds(), ["Sheet2", "Sheet1"]);

  // Also ensure applyState can reorder an existing controller to match the snapshot.
  const differentOrder = new DocumentController();
  differentOrder.getCell("Sheet1", "A1");
  differentOrder.getCell("Sheet2", "A1");
  assert.deepEqual(differentOrder.getSheetIds(), ["Sheet1", "Sheet2"]);

  differentOrder.applyState(snapshot);
  assert.deepEqual(differentOrder.getSheetIds(), ["Sheet2", "Sheet1"]);
});

test("update event fires on edits and undo/redo", () => {
  const doc = new DocumentController();
  let updates = 0;
  doc.on("update", () => {
    updates += 1;
  });

  doc.setCellValue("Sheet1", "A1", "x");
  doc.undo();
  doc.redo();

  assert.equal(updates, 3);
});
