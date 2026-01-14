// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (mergedRanges)", () => {
  it("syncs local mergedRanges updates into Yjs + applies remote updates back into DocumentController", () => {
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

    const mergedRanges = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 2 }];
    document.setMergedRanges(sheetId, mergedRanges);

    const viewMap = sheetMap.get("view") as Y.Map<any>;
    expect(viewMap).toBeInstanceOf(Y.Map);
    expect(viewMap.get("mergedRanges")).toEqual(mergedRanges);
    // Back-compat mirror: legacy key + top-level mirrors.
    expect(viewMap.get("mergedCells")).toEqual(mergedRanges);
    expect(sheetMap.get("mergedRanges")).toEqual(mergedRanges);
    expect(sheetMap.get("mergedCells")).toEqual(mergedRanges);

    // Local removal should delete both preferred + legacy keys.
    document.setMergedRanges(sheetId, []);
    expect(viewMap.get("mergedRanges")).toBe(undefined);
    expect(viewMap.get("mergedCells")).toBe(undefined);
    expect(sheetMap.get("mergedRanges")).toBe(undefined);
    expect(sheetMap.get("mergedCells")).toBe(undefined);

    // Remote update: binder should normalize reversed coordinates and preserve ordering.
    const remoteMergedRanges = [
      { startRow: 3, endRow: 2, startCol: 4, endCol: 1 },
      { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
    ];
    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      remoteViewMap.set("mergedRanges", remoteMergedRanges);
      sheetMap.set("mergedRanges", remoteMergedRanges);
    });

    expect(document.getMergedRanges(sheetId)).toEqual([
      { startRow: 0, endRow: 1, startCol: 0, endCol: 1 },
      { startRow: 2, endRow: 3, startCol: 1, endCol: 4 },
    ]);

    // Remote deletion should clear merged ranges from the DocumentController.
    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      remoteViewMap.delete("mergedRanges");
      remoteViewMap.delete("mergedCells");
      sheetMap.delete("mergedRanges");
      sheetMap.delete("mergedCells");
    });

    expect(document.getMergedRanges(sheetId)).toEqual([]);

    binder.destroy();
  });

  it("hydrates mergedRanges stored as a Y.Array (legacy/experimental encoding)", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheets.push([sheetMap]);

    const yRanges = new Y.Array<any>();
    yRanges.push([{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);
    doc.transact(() => {
      const view = new Y.Map<any>();
      view.set("mergedRanges", yRanges);
      sheetMap.set("view", view);
    });

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    expect(document.getMergedRanges(sheetId)).toEqual([{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);

    binder.destroy();
  });

  it("hydrates mergedRanges from the legacy mergedCells key when mergedRanges is absent", () => {
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

    const legacyMergedCells = [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }];
    doc.transact(() => {
      const viewMap = new Y.Map<any>();
      viewMap.set("mergedCells", legacyMergedCells);
      sheetMap.set("view", viewMap);
    });

    expect(document.getMergedRanges(sheetId)).toEqual(legacyMergedCells);

    binder.destroy();
  });
});
