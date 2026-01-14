// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (read-only sessions)", () => {
  it("does not write local sheet view changes into Yjs when session.isReadOnly() is true, but still applies remote updates", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => true } as any,
      documentController: document,
    });

    // Local changes should update the local DocumentController state, but should not persist
    // into the shared Yjs doc (viewer/commenter mode).
    document.setFrozen(sheetId, 2, 1);
    document.setColWidth(sheetId, 1, 120);
    document.setMergedRanges(sheetId, [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);

    expect(document.getSheetView(sheetId).frozenRows).toBe(2);
    expect(document.getSheetView(sheetId).frozenCols).toBe(1);
    expect(document.getSheetView(sheetId).colWidths).toEqual({ "1": 120 });
    expect(document.getMergedRanges(sheetId)).toEqual([{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);

    // But nothing should be written into Yjs.
    expect(sheetMap.get("view")).toBe(undefined);
    expect(sheetMap.get("frozenRows")).toBe(undefined);
    expect(sheetMap.get("colWidths")).toBe(undefined);
    expect(sheetMap.get("mergedRanges")).toBe(undefined);

    // Remote change should still apply back into the DocumentController.
    doc.transact(() => {
      const viewMap = new Y.Map<any>();
      viewMap.set("frozenRows", 4);
      viewMap.set("frozenCols", 3);
      viewMap.set("mergedRanges", [{ startRow: 2, endRow: 3, startCol: 1, endCol: 2 }]);
      const colWidths = new Y.Map<any>();
      colWidths.set("2", 150);
      viewMap.set("colWidths", colWidths);
      sheetMap.set("view", viewMap);
    });

    expect(document.getSheetView(sheetId).frozenRows).toBe(4);
    expect(document.getSheetView(sheetId).frozenCols).toBe(3);
    expect(document.getSheetView(sheetId).colWidths).toEqual({ "2": 150 });
    expect(document.getMergedRanges(sheetId)).toEqual([{ startRow: 2, endRow: 3, startCol: 1, endCol: 2 }]);

    binder.destroy();
  });
});

