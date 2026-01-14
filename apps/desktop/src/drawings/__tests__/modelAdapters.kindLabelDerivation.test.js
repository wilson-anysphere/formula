import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects derives shape labels from rawXml when label is missing/blank", () => {
  const rawXml = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="My Shape"/></xdr:nvSpPr></xdr:sp>';
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "", rawXml },
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
});

test("convertDocumentSheetDrawingsToUiDrawingObjects derives chart labels from rawXml when label is missing", () => {
  const rawXml =
    '<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/></xdr:nvGraphicFramePr></xdr:graphicFrame>';
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "chart", relId: "rId1", rawXml },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "chart");
  assert.equal(ui[0]?.kind?.chartId, "rId1");
  assert.equal(ui[0]?.kind?.label, "Chart 1");
});

test("convertDocumentSheetDrawingsToUiDrawingObjects derives unknown graphicFrame labels from rawXml when label is missing", () => {
  const rawXml =
    '<xdr:graphicFrame><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"></a:graphicData></a:graphic></xdr:graphicFrame>';
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "unknown", rawXml },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "unknown");
  assert.equal(ui[0]?.kind?.label, "SmartArt");
});
