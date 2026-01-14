import assert from "node:assert/strict";
import test from "node:test";

import { mergeDocumentStates } from "./merge.js";

test("mergeDocumentStates: merges sheet view mergedRanges + drawings without clobbering independent edits", () => {
  const base = {
    schemaVersion: 1,
    sheets: {
      order: ["Sheet1"],
      metaById: {
        Sheet1: {
          id: "Sheet1",
          name: "Sheet1",
          view: {
            frozenRows: 0,
            frozenCols: 0,
            mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }],
            drawings: [
              {
                id: 1,
                zOrder: 0,
                kind: { type: "image", imageId: "img-1" },
                anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
              },
            ],
          },
        },
      },
    },
    cells: { Sheet1: {} },
    metadata: {},
    namedRanges: {},
    comments: {},
  };

  // Ours moves the existing drawing and adds a new merged range.
  const ours = structuredClone(base);
  ours.sheets.metaById.Sheet1.view.drawings[0].anchor.pos.xEmu = 10;
  ours.sheets.metaById.Sheet1.view.mergedRanges.push({ startRow: 2, endRow: 3, startCol: 0, endCol: 1 });

  // Theirs inserts a new drawing and adds a different merged range.
  const theirs = structuredClone(base);
  theirs.sheets.metaById.Sheet1.view.drawings.push({
    id: 2,
    zOrder: 1,
    kind: { type: "image", imageId: "img-2" },
    anchor: { type: "absolute", pos: { xEmu: 5, yEmu: 5 }, size: { cx: 2, cy: 2 } },
  });
  theirs.sheets.metaById.Sheet1.view.mergedRanges.push({ startRow: 4, endRow: 5, startCol: 0, endCol: 1 });

  const result = mergeDocumentStates({ base, ours, theirs });
  assert.equal(result.conflicts.length, 0);

  assert.deepEqual(result.merged.sheets.metaById.Sheet1.view.drawings, [
    {
      id: 1,
      zOrder: 0,
      kind: { type: "image", imageId: "img-1" },
      anchor: { type: "absolute", pos: { xEmu: 10, yEmu: 0 }, size: { cx: 1, cy: 1 } },
    },
    {
      id: 2,
      zOrder: 1,
      kind: { type: "image", imageId: "img-2" },
      anchor: { type: "absolute", pos: { xEmu: 5, yEmu: 5 }, size: { cx: 2, cy: 2 } },
    },
  ]);

  assert.deepEqual(result.merged.sheets.metaById.Sheet1.view.mergedRanges, [
    { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    { startRow: 2, endRow: 3, startCol: 0, endCol: 1 },
    { startRow: 4, endRow: 5, startCol: 0, endCol: 1 },
  ]);
});

