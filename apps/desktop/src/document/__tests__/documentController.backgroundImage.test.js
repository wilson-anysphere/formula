import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("encodeState/applyState roundtrip preserves sheet backgroundImageId", () => {
  const doc = new DocumentController();

  const bytes = new Uint8Array([0x89, 0x50, 0x4e, 0x47]); // PNG signature prefix
  doc.setImage("bg.png", { bytes, mimeType: "image/png" });
  doc.setSheetBackgroundImageId("Sheet1", "bg.png");

  const snapshot = doc.encodeState();
  const parsed = JSON.parse(new TextDecoder().decode(snapshot));
  const sheet = parsed.sheets.find((s) => s.id === "Sheet1");
  assert.equal(sheet.backgroundImageId, "bg.png");

  const restored = new DocumentController();
  restored.applyState(snapshot);
  assert.equal(restored.getSheetBackgroundImageId("Sheet1"), "bg.png");
  assert.equal(restored.getSheetView("Sheet1").backgroundImageId, "bg.png");
});

test("sheet backgroundImageId changes flow through sheetView deltas and are undoable", () => {
  const doc = new DocumentController();

  /** @type {any} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.setSheetBackgroundImageId("Sheet1", "bg.png");
  assert.equal(doc.getSheetBackgroundImageId("Sheet1"), "bg.png");

  const deltas = Array.isArray(lastChange?.sheetViewDeltas) ? lastChange.sheetViewDeltas : [];
  const delta = deltas.find((d) => d?.sheetId === "Sheet1");
  assert.ok(delta, "expected a SheetViewDelta for Sheet1");
  assert.equal(delta.before?.backgroundImageId ?? null, null);
  assert.equal(delta.after?.backgroundImageId ?? null, "bg.png");

  doc.undo();
  assert.equal(doc.getSheetBackgroundImageId("Sheet1"), null);
  doc.redo();
  assert.equal(doc.getSheetBackgroundImageId("Sheet1"), "bg.png");

  // Clearing should also be undoable.
  doc.setSheetBackgroundImageId("Sheet1", null);
  assert.equal(doc.getSheetBackgroundImageId("Sheet1"), null);
  doc.undo();
  assert.equal(doc.getSheetBackgroundImageId("Sheet1"), "bg.png");
});

test("applyExternalSheetViewDeltas trims backgroundImageId (defensive)", () => {
  const doc = new DocumentController();

  doc.applyExternalSheetViewDeltas([
    {
      sheetId: "Sheet1",
      before: { frozenRows: 0, frozenCols: 0 },
      after: { frozenRows: 0, frozenCols: 0, backgroundImageId: " bg.png " },
    },
  ]);

  assert.equal(doc.getSheetBackgroundImageId("Sheet1"), "bg.png");
  assert.equal(doc.getSheetView("Sheet1").backgroundImageId, "bg.png");
});
