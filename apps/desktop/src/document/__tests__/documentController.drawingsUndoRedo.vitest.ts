import { describe, expect, it } from "vitest";

import { DocumentController } from "../documentController.js";

describe("DocumentController drawings undo/redo", () => {
  it("setSheetDrawings is undoable and survives encodeState/applyState", () => {
    const doc = new DocumentController();

    const drawings = [
      {
        id: "d1",
        kind: { type: "image", imageId: "image_1" },
        anchor: {
          type: "oneCell",
          from: { cell: { row: 0, col: 0 }, offset: { xEmu: 0, yEmu: 0 } },
          size: { cx: 914_400, cy: 914_400 },
        },
        zOrder: 0,
      },
    ];

    expect(doc.getSheetDrawings("Sheet1")).toEqual([]);
    doc.setSheetDrawings("Sheet1", drawings, { label: "Insert Picture" });

    expect(doc.canUndo).toBe(true);
    expect(doc.getSheetDrawings("Sheet1")).toEqual(drawings);

    // Returned arrays should be detached from the internal model.
    const mutated = doc.getSheetDrawings("Sheet1");
    mutated.push({ id: 999 });
    expect(doc.getSheetDrawings("Sheet1")).toEqual(drawings);

    doc.undo();
    expect(doc.getSheetDrawings("Sheet1")).toEqual([]);

    doc.redo();
    expect(doc.getSheetDrawings("Sheet1")).toEqual(drawings);

    const snapshot = doc.encodeState();
    const restored = new DocumentController();
    restored.applyState(snapshot);
    expect(restored.getSheetDrawings("Sheet1")).toEqual(drawings);
  });
});
