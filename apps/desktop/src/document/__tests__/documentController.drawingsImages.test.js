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

test("drawing helpers support numeric ids (overlay-compatible)", () => {
  const doc = new DocumentController();

  doc.setSheetDrawings("Sheet1", [
    {
      id: 1,
      zOrder: 0,
      anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
      kind: { type: "image", imageId: "img1" },
    },
  ]);

  doc.updateDrawing("Sheet1", 1, { zOrder: 2 });
  assert.equal(doc.getSheetDrawings("Sheet1")[0]?.id, 1);
  assert.equal(doc.getSheetDrawings("Sheet1")[0]?.zOrder, 2);

  doc.deleteDrawing("Sheet1", "1");
  assert.deepEqual(doc.getSheetDrawings("Sheet1"), []);
});

test("applyState accepts formula-model style image + drawings payloads", () => {
  const snapshot = new TextEncoder().encode(
    JSON.stringify({
      schemaVersion: 1,
      sheets: [
        {
          id: "Sheet1",
          name: "Sheet1",
          visibility: "visible",
          frozenRows: 0,
          frozenCols: 0,
          cells: [],
          drawings: [
            {
              id: 1,
              z_order: 3,
              anchor: { OneCell: { from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } }, ext: { cx: 100, cy: 80 } } },
              kind: { Image: { image_id: "img1.png" } },
            },
          ],
        },
      ],
      images: {
        images: {
          "img1.png": { bytes: [1, 2, 3], content_type: "image/png" },
        },
      },
    }),
  );

  const doc = new DocumentController();
  doc.applyState(snapshot);

  const image = doc.getImage("img1.png");
  assert.ok(image);
  assert.equal(image?.mimeType, "image/png");
  assert.deepEqual(Array.from(image?.bytes ?? []), [1, 2, 3]);

  const drawings = doc.getSheetDrawings("Sheet1");
  assert.equal(drawings.length, 1);
  assert.equal(drawings[0].id, 1);
  assert.equal(drawings[0].zOrder, 3);
  assert.ok(drawings[0].anchor);
  assert.ok(drawings[0].kind);
});

test("applyExternalImageDeltas updates image store without creating undo history", () => {
  const doc = new DocumentController();
  let lastChange = null;
  doc.on("change", (payload) => {
    lastChange = payload;
  });

  assert.equal(doc.canUndo, false);
  assert.equal(doc.isDirty, false);

  doc.applyExternalImageDeltas(
    [
      {
        imageId: "img_external",
        before: null,
        after: { bytes: new Uint8Array([7, 8, 9]), mimeType: "image/png" },
      },
    ],
    { source: "collab" },
  );

  assert.equal(doc.canUndo, false);
  assert.equal(doc.isDirty, true);

  const image = doc.getImage("img_external");
  assert.ok(image);
  assert.equal(image?.mimeType, "image/png");
  assert.deepEqual(Array.from(image?.bytes ?? []), [7, 8, 9]);

  assert.ok(lastChange);
  assert.equal(lastChange.source, "collab");
  assert.deepEqual(lastChange.imageDeltas, [
    { imageId: "img_external", before: null, after: { mimeType: "image/png", byteLength: 3 } },
  ]);
});
