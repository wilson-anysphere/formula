import assert from "node:assert/strict";
import test from "node:test";

import { convertModelAnchorToUiAnchor, convertModelDrawingObjectToUiDrawingObject } from "../modelAdapters.ts";

test("unwrapPossiblyTaggedEnum tolerates object-valued metadata keys alongside kind variants", () => {
  const rawXml = "<xdr:sp/>";
  const model = {
    id: 1,
    kind: { label: "My Shape", meta: { foo: 1 }, Shape: { raw_xml: rawXml } },
    anchor: { Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } } },
    z_order: 0,
  };

  const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "Sheet1" });
  assert.equal(ui.kind.type, "shape");
  assert.equal(ui.kind.label, "My Shape");
  assert.equal(ui.kind.rawXml ?? ui.kind.raw_xml, rawXml);
});

test("unwrapPossiblyTaggedEnum tolerates object-valued metadata keys alongside anchor variants", () => {
  const uiAnchor = convertModelAnchorToUiAnchor({
    meta: { foo: 1 },
    Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } },
  });

  assert.deepEqual(uiAnchor, { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 20 } });
});

test("convertModelDrawingObjectToUiDrawingObject tolerates chart placeholders missing rel ids when rawXml indicates a chart", () => {
  const rawXml =
    '<xdr:graphicFrame xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">' +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/>' +
    "</a:graphic>" +
    "</xdr:graphicFrame>";
  const model = {
    id: 2,
    kind: { ChartPlaceholder: { raw_xml: rawXml } },
    anchor: { Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } } },
    z_order: 0,
  };

  const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "Sheet1" });
  assert.equal(ui.kind.type, "chart");
  assert.equal(ui.kind.chartId, "Sheet1:2");
});

