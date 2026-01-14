import assert from "node:assert/strict";
import test from "node:test";

import { convertModelAnchorToUiAnchor } from "../modelAdapters.ts";

test("convertModelAnchorToUiAnchor accepts externally-tagged Anchor variants with *Anchor suffixes", () => {
  const anchor = {
    OneCellAnchor: {
      from: { cell: { row: 1, col: 2 }, offset: { x_emu: 3, y_emu: 4 } },
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

test("convertModelAnchorToUiAnchor accepts internally-tagged Anchor variants with *Anchor suffixes", () => {
  const anchor = {
    type: "AbsoluteAnchor",
    value: {
      pos: { x_emu: 1, y_emu: 2 },
      ext: { cx: 3, cy: 4 },
    },
  };
  const ui = convertModelAnchorToUiAnchor(anchor);
  assert.deepEqual(ui, { type: "absolute", pos: { xEmu: 1, yEmu: 2 }, size: { cx: 3, cy: 4 } });
});

