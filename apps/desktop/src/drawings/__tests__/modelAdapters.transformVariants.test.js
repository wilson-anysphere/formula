import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts snake_case transform keys", () => {
  const drawings = [
    {
      id: "1",
      zOrder: 0,
      kind: { type: "image", imageId: "img1" },
      anchor: { type: "cell", row: 0, col: 0 },
      size: { width: 10, height: 10 },
      transform: { rotation_deg: 30, flip_h: true, flip_v: false },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.transform, { rotationDeg: 30, flipH: true, flipV: false });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects defaults missing flip keys to false", () => {
  const drawings = [
    {
      id: "1",
      zOrder: 0,
      kind: { type: "image", imageId: "img1" },
      anchor: { type: "cell", row: 0, col: 0 },
      size: { width: 10, height: 10 },
      transform: { rotationDeg: 45 },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.transform, { rotationDeg: 45, flipH: false, flipV: false });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects ignores empty transform objects", () => {
  const drawings = [
    {
      id: "1",
      zOrder: 0,
      kind: { type: "image", imageId: "img1" },
      anchor: { type: "cell", row: 0, col: 0 },
      size: { width: 10, height: 10 },
      transform: {},
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.transform, undefined);
});

