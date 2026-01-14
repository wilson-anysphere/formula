import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts mixed schema: model kind enum + legacy DocumentController cell anchor", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { Image: { image_id: "img1" } },
      anchor: { type: "cell", sheetId: "Sheet1", row: 0, col: 0 },
      // Provide an explicit size so the legacy cell anchor can be promoted to a oneCell anchor.
      size: { cx: 789, cy: 321 },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.kind, { type: "image", imageId: "img1" });
  assert.deepEqual(ui[0]?.anchor, {
    type: "oneCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts mixed schema: internally-tagged kind enum + DocumentController anchor", () => {
  const rawXml = "<xdr:sp><a:xfrm rot=\"60000\"/></xdr:sp>";
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "Shape", value: { raw_xml: rawXml } },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.rawXml ?? ui[0]?.kind?.raw_xml, rawXml);
  // The formula-model kind includes a raw transform; ensure it can be extracted.
  assert.equal(ui[0]?.transform?.rotationDeg, 1);
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts mixed schema: externally-tagged kind enum with metadata keys + DocumentController anchor", () => {
  const rawXml = "<xdr:sp><a:xfrm rot=\"60000\"/></xdr:sp>";
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { label: "My Shape", Shape: { raw_xml: rawXml } },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "My Shape");
  assert.equal(ui[0]?.kind?.rawXml ?? ui[0]?.kind?.raw_xml, rawXml);
  assert.equal(ui[0]?.transform?.rotationDeg, 1);
});
