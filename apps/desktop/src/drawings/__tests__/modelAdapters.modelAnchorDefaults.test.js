import assert from "node:assert/strict";
import test from "node:test";

import { pxToEmu } from "../../shared/emu.js";

import { convertModelAnchorToUiAnchor } from "../modelAdapters.ts";

test("convertModelAnchorToUiAnchor defaults missing anchor point offsets to 0", () => {
  const anchor = {
    OneCell: {
      from: { cell: { row: 1, col: 2 } },
      ext: { cx: 10, cy: 20 },
    },
  };

  const ui = convertModelAnchorToUiAnchor(anchor);
  assert.deepEqual(ui, {
    type: "oneCell",
    from: { cell: { row: 1, col: 2 }, offset: { xEmu: 0, yEmu: 0 } },
    size: { cx: 10, cy: 20 },
  });
});

test("convertModelAnchorToUiAnchor defaults missing ext/size to 100x100px", () => {
  const anchor = {
    OneCell: {
      from: { cell: { row: 0, col: 0 }, offset: { x_emu: 0, y_emu: 0 } },
      // ext missing
    },
  };

  const ui = convertModelAnchorToUiAnchor(anchor);
  assert.deepEqual(ui.size, { cx: pxToEmu(100), cy: pxToEmu(100) });
});

test("convertModelAnchorToUiAnchor defaults missing absolute pos/ext payloads to 0,0 and 100x100px", () => {
  const anchor = { Absolute: {} };
  const ui = convertModelAnchorToUiAnchor(anchor);
  assert.deepEqual(ui, { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: pxToEmu(100), cy: pxToEmu(100) } });
});

test("convertModelAnchorToUiAnchor accepts singleton-wrapped cell refs (interop)", () => {
  const anchor = {
    OneCell: {
      from: {
        cell: { 0: { row: { 0: 1 }, col: [2] } },
        offset: { x_emu: 3, y_emu: 4 },
      },
      ext: { cx: 10, cy: 20 },
    },
  };

  const ui = convertModelAnchorToUiAnchor(anchor);
  assert.deepEqual(ui, {
    type: "oneCell",
    from: { cell: { row: 1, col: 2 }, offset: { xEmu: 3, yEmu: 4 } },
    size: { cx: 10, cy: 20 },
  });
});
