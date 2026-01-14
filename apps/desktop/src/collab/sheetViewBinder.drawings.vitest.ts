// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (drawings)", () => {
  it("syncs local drawings updates into Yjs + applies remote drawings updates back into DocumentController", () => {
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

    const drawings = [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ];
    document.setSheetDrawings(sheetId, drawings);

    const viewMap = sheetMap.get("view") as Y.Map<any>;
    expect(viewMap).toBeInstanceOf(Y.Map);
    expect(viewMap.get("drawings")).toEqual(drawings);
    // Back-compat mirror.
    expect(sheetMap.get("drawings")).toEqual(drawings);

    const remoteDrawings = [
      {
        id: "drawing-2",
        zOrder: 1,
        kind: { type: "shape", label: "Rect" },
        anchor: { type: "absolute", pos: { xEmu: 10, yEmu: 10 }, size: { cx: 2, cy: 3 } },
      },
    ];

    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      remoteViewMap.set("drawings", remoteDrawings);
    });

    expect((document as any).getSheetDrawings(sheetId)).toEqual(remoteDrawings);

    binder.destroy();
  });

  it("hydrates drawings stored as a Y.Array (legacy/experimental encoding)", () => {
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

    const yDrawings = new Y.Array<any>();
    yDrawings.push([
      {
        id: "drawing-yarray",
        zOrder: 0,
        kind: { type: "shape", label: "From Y.Array" },
        anchor: { type: "absolute", pos: { xEmu: 1, yEmu: 2 }, size: { cx: 3, cy: 4 } },
      },
    ]);

    doc.transact(() => {
      // Use the top-level key to simulate older docs that didn't nest under `view`.
      sheetMap.set("drawings", yDrawings);
    });

    expect((document as any).getSheetDrawings(sheetId)).toEqual([
      {
        id: "drawing-yarray",
        zOrder: 0,
        kind: { type: "shape", label: "From Y.Array" },
        anchor: { type: "absolute", pos: { xEmu: 1, yEmu: 2 }, size: { cx: 3, cy: 4 } },
      },
    ]);

    binder.destroy();
  });

  it("ignores remote drawings with excessively long ids (defensive)", () => {
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

    // Start with one valid drawing so we can observe the remote invalid update clearing it.
    document.setSheetDrawings(sheetId, [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "image", imageId: "img-1" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ]);
    expect((document as any).getSheetDrawings(sheetId).length).toBe(1);

    const tooLongId = "x".repeat(5000);
    doc.transact(() => {
      const viewMap = sheetMap.get("view") as Y.Map<any>;
      viewMap.set("drawings", [
        {
          id: tooLongId,
          zOrder: 0,
          kind: { type: "shape", label: "bad" },
          anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
        },
      ]);
    });

    // The binder should ignore the invalid entry and clear drawings (DocumentController also rejects
    // long ids), rather than throwing or storing the oversized payload.
    expect((document as any).getSheetDrawings(sheetId)).toEqual([]);

    binder.destroy();
  });
});
