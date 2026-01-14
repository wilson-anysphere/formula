import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts OneCell anchor type variants (OneCell) without falling back to model parsing", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "OneCell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  // If we accidentally fell back to the formula-model adapter, the `label` would be dropped for
  // shape kinds (it is derived from rawXml there). Preserve the DocumentController label instead.
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "oneCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts OneCell anchor type variants (one_cell) without falling back to model parsing", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "one_cell",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "oneCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts OneCellAnchor type variants (OneCellAnchor)", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "OneCellAnchor",
        from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "oneCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts TwoCell anchor type variants (TwoCell)", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "TwoCell",
        from: { cell: { row: 0, col: 0 }, dxEmu: 0, dyEmu: 0 },
        to: { cell: { row: 1, col: 1 }, dxEmu: 0, dyEmu: 0 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "twoCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
    to: { cell: { row: 1, col: 1 }, offset: { xEmu: 0, yEmu: 0 } },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts Absolute anchor type variants (Absolute)", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "Absolute",
        pos: { x_emu: 0, y_emu: 0 },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "absolute",
    pos: { xEmu: 0, yEmu: 0 },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects unwraps internally-tagged anchors with content (type/value) without losing DocumentController kind labels", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "Absolute",
        value: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 789, cy: 321 },
        },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "absolute",
    pos: { xEmu: 0, yEmu: 0 },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects unwraps externally-tagged anchors ({ Absolute: {...} }) without losing DocumentController kind labels", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        Absolute: {
          pos: { x_emu: 0, y_emu: 0 },
          ext: { cx: 789, cy: 321 },
        },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "shape");
  assert.equal(ui[0]?.kind?.label, "Box");
  assert.deepEqual(ui[0]?.anchor, {
    type: "absolute",
    pos: { xEmu: 0, yEmu: 0 },
    size: { cx: 789, cy: 321 },
  });
});
