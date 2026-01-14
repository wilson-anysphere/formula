import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects reads absolute anchors stored with pos: {xEmu,yEmu}", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 123, yEmu: 456 },
        size: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects tolerates mixed size encodings (cx null, widthEmu present)", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 123, yEmu: 456 },
        size: { cx: null, cy: null, widthEmu: 789, heightEmu: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects reads absolute anchors stored with root xEmu/yEmu keys", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        xEmu: 123,
        yEmu: 456,
        size: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects falls back to root offsets when pos offsets are invalid", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        xEmu: 123,
        yEmu: 456,
        // Some mixed schemas include `pos`, but with invalid values (e.g. null) while the root is valid.
        pos: { xEmu: null, yEmu: null },
        size: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects reads absolute anchor size from ext when size is absent", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 123, yEmu: 456 },
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects falls back to ext when size is present but invalid", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 123, yEmu: 456 },
        size: {},
        ext: { cx: 789, cy: 321 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects reads absolute anchor size from root cx/cy keys", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "shape", label: "Box" },
      anchor: {
        type: "absolute",
        xEmu: 123,
        yEmu: 456,
        cx: 789,
        cy: 321,
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.anchor, { type: "absolute", pos: { xEmu: 123, yEmu: 456 }, size: { cx: 789, cy: 321 } });
});
