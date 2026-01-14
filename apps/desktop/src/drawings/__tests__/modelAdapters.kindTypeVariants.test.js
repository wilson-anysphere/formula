import assert from "node:assert/strict";
import test from "node:test";

import { convertDocumentSheetDrawingsToUiDrawingObjects } from "../modelAdapters.ts";

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts kind.type case variants (Shape) and preserves label", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "Shape", label: "Box" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.deepEqual(ui[0]?.kind, { type: "shape", label: "Box" });
});

test("convertDocumentSheetDrawingsToUiDrawingObjects accepts kind.type underscore variants (chart_placeholder)", () => {
  const drawings = [
    {
      id: "7",
      zOrder: 0,
      kind: { type: "chart_placeholder", relId: "rId1", label: "Chart" },
      anchor: {
        type: "absolute",
        pos: { xEmu: 0, yEmu: 0 },
        size: { cx: 100, cy: 50 },
      },
    },
  ];

  const ui = convertDocumentSheetDrawingsToUiDrawingObjects(drawings);
  assert.equal(ui.length, 1);
  assert.equal(ui[0]?.kind?.type, "chart");
  assert.equal(ui[0]?.kind?.chartId, "rId1");
});

