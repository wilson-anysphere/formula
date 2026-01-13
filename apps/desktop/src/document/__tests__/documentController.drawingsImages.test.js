import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../documentController.js";

test("encodeState/applyState roundtrip preserves images + drawings", () => {
  const doc = new DocumentController();

  doc.setImage("img1", { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
  doc.setSheetDrawings("Sheet1", [
    {
      id: "d1",
      zOrder: 1,
      anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
      kind: { type: "image", imageId: "img1" },
      size: { width: 100, height: 80 },
    },
  ]);

  const snapshot = doc.encodeState();
  const parsed = JSON.parse(new TextDecoder().decode(snapshot));

  assert.deepEqual(parsed.images, [{ id: "img1", mimeType: "image/png", bytesBase64: "AQID" }]);
  assert.ok(parsed.drawingsBySheet);
  assert.ok(Array.isArray(parsed.drawingsBySheet.Sheet1));
  assert.equal(parsed.drawingsBySheet.Sheet1.length, 1);

  const restored = new DocumentController();
  restored.applyState(snapshot);

  const image = restored.getImage("img1");
  assert.ok(image);
  assert.equal(image?.mimeType, "image/png");
  assert.deepEqual(Array.from(image?.bytes ?? []), [1, 2, 3]);

  assert.deepEqual(restored.getSheetDrawings("Sheet1"), doc.getSheetDrawings("Sheet1"));
});

test("undo/redo restores drawings + images", () => {
  const doc = new DocumentController();

  doc.setImage("img1", { bytes: new Uint8Array([9, 9]), mimeType: "image/png" }, { label: "Set Image" });
  doc.setSheetDrawings(
    "Sheet1",
    [
      {
        id: "d1",
        zOrder: 1,
        anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
        kind: { type: "image", imageId: "img1" },
      },
    ],
    { label: "Set Drawings" },
  );

  assert.equal(doc.canUndo, true);

  // Undo drawings.
  assert.equal(doc.undo(), true);
  assert.deepEqual(doc.getSheetDrawings("Sheet1"), []);
  assert.deepEqual(Array.from(doc.getImage("img1")?.bytes ?? []), [9, 9]);

  // Undo image store set.
  assert.equal(doc.undo(), true);
  assert.equal(doc.getImage("img1"), null);

  // Redo image store set.
  assert.equal(doc.redo(), true);
  assert.deepEqual(Array.from(doc.getImage("img1")?.bytes ?? []), [9, 9]);

  // Redo drawings.
  assert.equal(doc.redo(), true);
  assert.equal(doc.getSheetDrawings("Sheet1").length, 1);
});

test("change event includes drawingDeltas + imageDeltas", () => {
  const doc = new DocumentController();
  /** @type {any} */
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  doc.setImage("img1", { bytes: new Uint8Array([5, 6]), mimeType: "image/png" });
  assert.ok(lastChange, "expected a change event");
  assert.ok(Array.isArray(lastChange.imageDeltas));
  assert.deepEqual(lastChange.imageDeltas, [
    { imageId: "img1", before: null, after: { mimeType: "image/png", byteLength: 2 } },
  ]);

  doc.setSheetDrawings("Sheet1", [
    {
      id: "d1",
      zOrder: 1,
      anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
      kind: { type: "image", imageId: "img1" },
    },
  ]);
  assert.ok(Array.isArray(lastChange.drawingDeltas));
  assert.equal(lastChange.drawingDeltas.length, 1);
  assert.equal(lastChange.drawingDeltas[0].sheetId, "Sheet1");
  assert.equal(lastChange.drawingDeltas[0].after[0].id, "d1");
});

