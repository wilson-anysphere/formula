import test from "node:test";
import assert from "node:assert/strict";

import { DocumentController } from "../../../document/documentController.js";
import { applyBranchStateToDocumentController, documentControllerToBranchState } from "../branchStateAdapter.js";

test("branchStateAdapter round-trips images + drawings via DocumentState.metadata", () => {
  const doc = new DocumentController();

  doc.setImage("img1", { bytes: new Uint8Array([1, 2, 3]), mimeType: "image/png" });
  doc.setSheetDrawings("Sheet1", [
    {
      id: "d1",
      zOrder: 1,
      anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
      kind: { type: "image", imageId: "img1" },
    },
  ]);

  const state = documentControllerToBranchState(doc);
  assert.ok(Array.isArray(state.metadata.images));
  assert.deepEqual(state.metadata.images, [{ id: "img1", mimeType: "image/png", bytesBase64: "AQID" }]);
  assert.ok(state.metadata.drawingsBySheet);
  assert.equal(state.metadata.drawingsBySheet.Sheet1?.[0]?.id, "d1");

  const restored = new DocumentController();
  applyBranchStateToDocumentController(restored, state);

  assert.deepEqual(restored.getSheetDrawings("Sheet1"), doc.getSheetDrawings("Sheet1"));
  const image = restored.getImage("img1");
  assert.ok(image);
  assert.equal(image?.mimeType, "image/png");
  assert.deepEqual(Array.from(image?.bytes ?? []), [1, 2, 3]);
});

