// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (duplicate sheets)", () => {
  it("writes drawings to all duplicate sheet entries (so later schema normalization doesn't drop them)", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetA = new Y.Map<any>();
    sheetA.set("id", sheetId);
    sheetA.set("name", "Sheet1-A");
    const sheetB = new Y.Map<any>();
    sheetB.set("id", sheetId);
    sheetB.set("name", "Sheet1-B");
    sheets.push([sheetA, sheetB]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    const drawings = [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ];
    document.setSheetDrawings(sheetId, drawings);

    const viewA = sheetA.get("view") as Y.Map<any>;
    const viewB = sheetB.get("view") as Y.Map<any>;
    expect(viewA).toBeInstanceOf(Y.Map);
    expect(viewB).toBeInstanceOf(Y.Map);
    expect(viewA.get("drawings")).toEqual(drawings);
    expect(viewB.get("drawings")).toEqual(drawings);
    // Back-compat mirror.
    expect(sheetA.get("drawings")).toEqual(drawings);
    expect(sheetB.get("drawings")).toEqual(drawings);

    binder.destroy();
  });

  it("reads the last duplicate sheet entry by index (matches ensureWorkbookSchema pruning behavior)", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetA = new Y.Map<any>();
    sheetA.set("id", sheetId);
    sheetA.set("name", "Sheet1-A");
    const sheetB = new Y.Map<any>();
    sheetB.set("id", sheetId);
    sheetB.set("name", "Sheet1-B");
    sheets.push([sheetA, sheetB]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    const drawingsA = [
      {
        id: "drawing-a",
        zOrder: 0,
        kind: { type: "image", imageId: "img-a" },
        anchor: { type: "absolute", pos: { xEmu: 1, yEmu: 1 }, size: { cx: 1, cy: 1 } },
      },
    ];
    const drawingsB = [
      {
        id: "drawing-b",
        zOrder: 0,
        kind: { type: "image", imageId: "img-b" },
        anchor: { type: "absolute", pos: { xEmu: 2, yEmu: 2 }, size: { cx: 1, cy: 1 } },
      },
    ];

    doc.transact(() => {
      (sheetA.get("view") as Y.Map<any> | undefined)?.set?.("drawings", drawingsA);
      (sheetB.get("view") as Y.Map<any> | undefined)?.set?.("drawings", drawingsB);
      // If the view doesn't exist yet, create it (matches how other code writes).
      if (!sheetA.get("view")) {
        const view = new Y.Map();
        view.set("drawings", drawingsA);
        sheetA.set("view", view);
      }
      if (!sheetB.get("view")) {
        const view = new Y.Map();
        view.set("drawings", drawingsB);
        sheetB.set("view", view);
      }
    });

    // The binder reads the last duplicate by index (sheetB).
    expect((document as any).getSheetDrawings(sheetId)).toEqual(drawingsB);

    binder.destroy();
  });
});

