import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects treats chart graphicFrames as charts even when chartId is missing/unknown", () => {
  const rawXml =
    '<xdr:graphicFrame><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>' +
    "</a:graphic></xdr:graphicFrame>";
  const drawings = [
    {
      id: "1",
      zOrder: 0,
      kind: { type: "chart", chartId: "unknown", rawXml },
      anchor: { type: "cell", row: 0, col: 0 },
      size: { width: 10, height: 10 },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings, { sheetId: "Sheet1" });
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind.type, "chart");
  assert.equal(ui[0]?.kind.chartId, "Sheet1:1");
});

test("convertDocumentSheetDrawingsToUiDrawingObjects derives chartIds from singleton-wrapped anchor.sheetId (interop)", () => {
  const rawXml =
    '<xdr:graphicFrame><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>' +
    "</a:graphic></xdr:graphicFrame>";
  const drawings = [
    {
      id: "1",
      zOrder: 0,
      kind: { type: "chart", chartId: "unknown", rawXml },
      anchor: { type: "cell", sheetId: { 0: "Sheet1" }, row: 0, col: 0 },
      size: { width: 10, height: 10 },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind.type, "chart");
  assert.equal(ui[0]?.kind.chartId, "Sheet1:1");
});

test("convertDocumentSheetDrawingsToUiDrawingObjects treats chartPlaceholder graphicFrames as charts when chartId is missing/unknown", () => {
  const rawXml =
    '<xdr:graphicFrame><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>' +
    "</a:graphic></xdr:graphicFrame>";
  const drawings = [
    {
      id: "5",
      zOrder: 0,
      kind: { type: "chartPlaceholder", relId: "unknown", rawXml },
      anchor: { type: "cell", row: 0, col: 0 },
      size: { width: 10, height: 10 },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings, { sheetId: "Sheet1" });
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind.type, "chart");
  assert.equal(ui[0]?.kind.chartId, "Sheet1:5");
});
