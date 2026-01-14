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

test("convertDocumentSheetDrawingsToUiDrawingObjects derives transform from kind.rawXml when transform is absent", () => {
  const rawXml = "<xdr:sp><a:xfrm rot=\"60000\"/></xdr:sp>";
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", rawXml },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.transform?.rotationDeg, 1);
  assert.equal(ui[0]?.transform?.flipH, false);
  assert.equal(ui[0]?.transform?.flipV, false);
});

test("convertDocumentSheetDrawingsToUiDrawingObjects derives image transform from preserved xlsx.pic_xml when transform is absent", () => {
  const picXml = `
    <xdr:pic xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <xdr:spPr>
        <a:xfrm rot="5400000" flipV="1">
          <a:off x="0" y="0"/>
          <a:ext cx="1000" cy="500"/>
        </a:xfrm>
      </xdr:spPr>
    </xdr:pic>
  `;
  const drawings = [
    {
      id: "1",
      zOrder: 0,
      kind: { type: "image", imageId: "img1" },
      anchor: { type: "cell", row: 0, col: 0 },
      size: { width: 10, height: 10 },
      preserved: { "xlsx.pic_xml": picXml },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.transform, { rotationDeg: 90, flipH: false, flipV: true });
});
