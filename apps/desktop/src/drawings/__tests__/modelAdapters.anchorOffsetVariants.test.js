import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts oneCell anchors whose offsets use dxEmu/dyEmu keys", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "oneCell",
        from: {
          cell: { row: 0, col: 0 },
          offset: { dxEmu: 123, dyEmu: 456 },
        },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, {
    type: "oneCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 123, yEmu: 456 } },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts oneCell anchors whose offsets are stored directly on the from payload", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "oneCell",
        from: {
          cell: { row: 0, col: 0 },
          dxEmu: 123,
          dyEmu: 456,
        },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, {
    type: "oneCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 123, yEmu: 456 } },
    size: { cx: 789, cy: 321 },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts twoCell anchors whose offsets use dxEmu/dyEmu keys", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "twoCell",
        from: {
          cell: { row: 0, col: 0 },
          offset: { dxEmu: 123, dyEmu: 456 },
        },
        to: {
          cell: { row: 1, col: 1 },
          offset: { dxEmu: 0, dyEmu: 0 },
        },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, {
    type: "twoCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 123, yEmu: 456 } },
    to: { cell: { row: 1, col: 1 }, offset: { xEmu: 0, yEmu: 0 } },
  });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts twoCell anchors whose offsets are stored directly on the from/to payloads", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "twoCell",
        from: {
          cell: { row: 0, col: 0 },
          dxEmu: 123,
          dyEmu: 456,
        },
        to: {
          cell: { row: 1, col: 1 },
          dxEmu: 0,
          dyEmu: 0,
        },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, {
    type: "twoCell",
    from: { cell: { row: 0, col: 0 }, offset: { xEmu: 123, yEmu: 456 } },
    to: { cell: { row: 1, col: 1 }, offset: { xEmu: 0, yEmu: 0 } },
  });
});
