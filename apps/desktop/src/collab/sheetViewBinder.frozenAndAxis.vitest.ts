// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (frozen + axis overrides)", () => {
  it("syncs frozen pane counts between DocumentController and Yjs", () => {
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

    document.setFrozen(sheetId, 2, 1);

    const viewMap = sheetMap.get("view") as Y.Map<any>;
    expect(viewMap).toBeInstanceOf(Y.Map);
    // Preferred nested view keys.
    expect(viewMap.get("frozenRows")).toBe(2);
    expect(viewMap.get("frozenCols")).toBe(1);
    // Back-compat top-level mirrors.
    expect(sheetMap.get("frozenRows")).toBe(2);
    expect(sheetMap.get("frozenCols")).toBe(1);

    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      remoteViewMap.set("frozenRows", 4);
      remoteViewMap.set("frozenCols", 3);
      sheetMap.set("frozenRows", 4);
      sheetMap.set("frozenCols", 3);
    });

    expect(document.getSheetView(sheetId).frozenRows).toBe(4);
    expect(document.getSheetView(sheetId).frozenCols).toBe(3);

    binder.destroy();
  });

  it("syncs colWidths/rowHeights axis overrides between DocumentController and Yjs", () => {
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

    // Local updates -> Yjs.
    document.setColWidth(sheetId, 1, 120);
    document.setColWidth(sheetId, 3, 200);
    document.setRowHeight(sheetId, 0, 40);

    const viewMap = sheetMap.get("view") as Y.Map<any>;
    const colWidths = viewMap.get("colWidths") as Y.Map<any>;
    const rowHeights = viewMap.get("rowHeights") as Y.Map<any>;

    expect(colWidths).toBeInstanceOf(Y.Map);
    expect(rowHeights).toBeInstanceOf(Y.Map);

    expect(colWidths.get("1")).toBe(120);
    expect(colWidths.get("3")).toBe(200);
    expect(rowHeights.get("0")).toBe(40);

    // Local removal -> deletes the specific key.
    document.resetColWidth(sheetId, 1);
    expect(colWidths.get("1")).toBe(undefined);
    expect(colWidths.get("3")).toBe(200);

    // Remote updates -> DocumentController.
    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      const remoteColWidths = remoteViewMap.get("colWidths") as Y.Map<any>;
      const remoteRowHeights = remoteViewMap.get("rowHeights") as Y.Map<any>;
      remoteColWidths.set("2", 150);
      remoteColWidths.delete("3");
      remoteRowHeights.set("1", 60);
      remoteRowHeights.delete("0");
    });

    expect(document.getSheetView(sheetId).colWidths).toEqual({ "2": 150 });
    expect(document.getSheetView(sheetId).rowHeights).toEqual({ "1": 60 });

    binder.destroy();
  });

  it("hydrates legacy top-level axis overrides when `view` is missing", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheetMap.set("colWidths", { "1": 100 });
    sheetMap.set("rowHeights", { "0": 25 });
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    expect(document.getSheetView(sheetId).colWidths).toEqual({ "1": 100 });
    expect(document.getSheetView(sheetId).rowHeights).toEqual({ "0": 25 });

    binder.destroy();
  });

  it("hydrates legacy axis overrides encoded as arrays when `view` is missing", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    // Support both tuple and `{index,size}` shapes.
    sheetMap.set("colWidths", [
      [1, 100],
      [2, 200],
      [-1, 50], // ignored
      ["bad", 50], // ignored
      [3, 0], // ignored (non-positive)
    ]);
    sheetMap.set("rowHeights", [
      { index: 0, size: 25 },
      { index: 2, size: 30 },
      { index: -1, size: 10 }, // ignored
      { index: 1, size: 0 }, // ignored (non-positive)
    ]);
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    expect(document.getSheetView(sheetId).colWidths).toEqual({ "1": 100, "2": 200 });
    expect(document.getSheetView(sheetId).rowHeights).toEqual({ "0": 25, "2": 30 });

    binder.destroy();
  });
});
