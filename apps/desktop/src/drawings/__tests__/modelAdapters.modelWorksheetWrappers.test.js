import assert from "node:assert/strict";
import test from "node:test";

import {
  convertModelWorkbookDrawingsToUiDrawingLayer,
  convertModelWorksheetDrawingsToUiDrawingObjects,
} from "../modelAdapters.ts";

test("convertModelWorksheetDrawingsToUiDrawingObjects accepts singleton-wrapped drawings arrays (interop)", () => {
  const modelDrawing = {
    id: 1,
    kind: { Image: { image_id: "img1.png" } },
    anchor: { Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } } },
    z_order: 0,
  };

  const worksheetObjectWrapped = { id: "Sheet1", drawings: { 0: [modelDrawing] } };
  const ui1 = convertModelWorksheetDrawingsToUiDrawingObjects(worksheetObjectWrapped);
  assert.equal(ui1.length, 1);
  assert.deepEqual(ui1[0]?.kind, { type: "image", imageId: "img1.png" });

  const worksheetArrayWrapped = { id: "Sheet1", drawings: [[modelDrawing]] };
  const ui2 = convertModelWorksheetDrawingsToUiDrawingObjects(worksheetArrayWrapped);
  assert.equal(ui2.length, 1);
  assert.deepEqual(ui2[0]?.kind, { type: "image", imageId: "img1.png" });
});

test("convertModelWorkbookDrawingsToUiDrawingLayer accepts singleton-wrapped sheets arrays (interop)", () => {
  const workbook = {
    images: { images: { "img1.png": { bytes: [1, 2, 3], content_type: "image/png" } } },
    sheets: {
      0: [
        {
          name: "Sheet1",
          drawings: [
            {
              id: 1,
              kind: { Image: { image_id: "img1.png" } },
              anchor: { Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } } },
              z_order: 0,
            },
          ],
        },
      ],
    },
  };

  const ui = convertModelWorkbookDrawingsToUiDrawingLayer(workbook);
  assert.ok(ui.images.get("img1.png"));
  assert.equal(ui.drawingsBySheetName.Sheet1?.length, 1);
  assert.deepEqual(ui.drawingsBySheetName.Sheet1?.[0]?.kind, { type: "image", imageId: "img1.png" });
});

