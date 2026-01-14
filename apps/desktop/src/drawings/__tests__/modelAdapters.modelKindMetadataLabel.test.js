import assert from "node:assert/strict";
import test from "node:test";

import { convertModelDrawingObjectToUiDrawingObject } from "../modelAdapters.ts";

test("convertModelDrawingObjectToUiDrawingObject preserves kind.label metadata stored alongside externally-tagged enums", () => {
  const rawXml = '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="Derived Name"/></xdr:nvSpPr></xdr:sp>';
  const model = {
    id: 1,
    kind: { label: "My Shape", Shape: { raw_xml: rawXml } },
    anchor: { Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } } },
    z_order: 0,
  };

  const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "Sheet1" });
  assert.equal(ui.kind.type, "shape");
  assert.equal(ui.kind.label, "My Shape");
  assert.equal(ui.kind.rawXml ?? ui.kind.raw_xml, rawXml);
});

test("convertModelDrawingObjectToUiDrawingObject accepts kind tags with a *Kind suffix", () => {
  const rawXml = "<xdr:sp/>";
  const model = {
    id: 1,
    kind: { ShapeKind: { raw_xml: rawXml } },
    anchor: { Absolute: { pos: { x_emu: 0, y_emu: 0 }, ext: { cx: 10, cy: 20 } } },
    z_order: 0,
  };

  const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "Sheet1" });
  assert.equal(ui.kind.type, "shape");
  assert.equal(ui.kind.rawXml ?? ui.kind.raw_xml, rawXml);
});

