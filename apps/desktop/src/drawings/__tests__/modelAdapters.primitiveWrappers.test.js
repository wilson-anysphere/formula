import assert from "node:assert/strict";
import test from "node:test";

import { convertModelDrawingObjectToUiDrawingObject } from "../modelAdapters.ts";

test("convertModelDrawingObjectToUiDrawingObject tolerates singleton-wrapped primitive fields (interop)", () => {
  const rawXml = "<xdr:sp/>";
  const model = {
    id: 1,
    kind: { Shape: { raw_xml: { 0: rawXml } } },
    anchor: {
      Absolute: {
        pos: { x_emu: { 0: 0 }, y_emu: [0] },
        ext: { cx: { 0: 10 }, cy: [20] },
      },
    },
    z_order: { 0: 5 },
  };

  const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "Sheet1" });
  assert.equal(ui.zOrder, 5);
  assert.equal(ui.kind.type, "shape");
  assert.equal(ui.kind.rawXml ?? ui.kind.raw_xml, rawXml);
  assert.deepEqual(ui.anchor, { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 20 } });
});

test("convertModelDrawingObjectToUiDrawingObject tolerates singleton wrappers around kind/anchor enums (interop)", () => {
  const rawXml = "<xdr:sp/>";
  const model = {
    id: 1,
    kind: { 0: { Shape: { raw_xml: rawXml } } },
    anchor: {
      0: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 10, cy: 20 },
        },
      },
    },
    z_order: 0,
  };

  const ui = convertModelDrawingObjectToUiDrawingObject(model, { sheetId: "Sheet1" });
  assert.equal(ui.kind.type, "shape");
  assert.equal(ui.kind.rawXml ?? ui.kind.raw_xml, rawXml);
  assert.deepEqual(ui.anchor, { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 10, cy: 20 } });
});
