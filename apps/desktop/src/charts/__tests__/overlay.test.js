import assert from "node:assert/strict";
import test from "node:test";

import { anchorToRectPx, emuToPx } from "../overlay.js";

test("emuToPx converts EMU to CSS pixels (96 dpi)", () => {
  assert.equal(emuToPx(914400), 96);
  assert.equal(emuToPx(0), 0);
});

test("anchorToRectPx converts absolute anchors", () => {
  const rect = anchorToRectPx(
    { kind: "absolute", xEmu: 914400, yEmu: 457200, cxEmu: 914400, cyEmu: 914400 },
    {}
  );
  assert.ok(rect);
  assert.equal(rect.left, 96);
  assert.equal(rect.top, 48);
  assert.equal(rect.width, 96);
  assert.equal(rect.height, 96);
});

test("anchorToRectPx converts oneCell anchors using ext", () => {
  const rect = anchorToRectPx(
    {
      kind: "oneCell",
      fromCol: 2,
      fromRow: 3,
      fromColOffEmu: 0,
      fromRowOffEmu: 0,
      cxEmu: 914400,
      cyEmu: 457200,
    },
    { defaultColWidthPx: 10, defaultRowHeightPx: 20 }
  );
  assert.ok(rect);
  assert.equal(rect.left, 20);
  assert.equal(rect.top, 60);
  assert.equal(rect.width, 96);
  assert.equal(rect.height, 48);
});

test("anchorToRectPx converts twoCell anchors using cell metrics", () => {
  const rect = anchorToRectPx(
    {
      kind: "twoCell",
      fromCol: 1,
      fromRow: 2,
      fromColOffEmu: 0,
      fromRowOffEmu: 0,
      toCol: 4,
      toRow: 5,
      toColOffEmu: 0,
      toRowOffEmu: 0,
    },
    { defaultColWidthPx: 10, defaultRowHeightPx: 20 }
  );
  assert.ok(rect);
  assert.equal(rect.left, 10);
  assert.equal(rect.top, 40);
  assert.equal(rect.width, 30);
  assert.equal(rect.height, 60);
});

