// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (source filtering)", () => {
  it("does not write DocumentController changes with source=collab back into Yjs (prevents echo from full binder)", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    const before = document.getSheetView(sheetId) as any;
    const after = { ...before, frozenRows: 2, frozenCols: 1 };
    document.applyExternalSheetViewDeltas([{ sheetId, before, after }], { source: "collab" });

    expect(sheetMap.get("view")).toBe(undefined);
    expect(sheetMap.get("frozenRows")).toBe(undefined);
    expect(sheetMap.get("frozenCols")).toBe(undefined);

    const drawings = [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "shape", label: "Rect" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ];
    document.applyExternalDrawingDeltas([{ sheetId, before: [], after: drawings }], { source: "collab" });

    expect(sheetMap.get("view")).toBe(undefined);
    expect(sheetMap.get("drawings")).toBe(undefined);

    binder.destroy();
  });

  it("does not write DocumentController applyState snapshots back into Yjs", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    // Build a snapshot with non-default view state and apply it into the bound controller.
    const other = new DocumentController();
    other.addSheet({ sheetId, name: "Sheet1" });
    other.setSheetBackgroundImageId(sheetId, "bg.png");
    other.setMergedRanges(sheetId, [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);
    other.setSheetDrawings(sheetId, [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 5, yEmu: 5 }, size: { cx: 2, cy: 3 } },
      },
    ]);
    const snapshot = other.encodeState();

    document.applyState(snapshot);

    expect(sheetMap.get("view")).toBe(undefined);
    expect(sheetMap.get("backgroundImageId")).toBe(undefined);
    expect(sheetMap.get("mergedRanges")).toBe(undefined);
    expect(sheetMap.get("drawings")).toBe(undefined);

    binder.destroy();
  });
});

