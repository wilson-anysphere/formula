// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (plain object `view` encoding)", () => {
  it("hydrates from a plain-object `view` and preserves unknown keys when migrating to Y.Map", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheetMap.set("view", {
      frozenRows: 1,
      frozenCols: 2,
      colWidths: { "1": 100 },
      rowHeights: { "0": 25 },
      mergedRanges: [{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }],
      // Simulate non-binder keys (e.g. drawings payloads from other clients/features).
      customKey: { foo: "bar" },
    });
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    // Initial hydration should read from the object view.
    expect(document.getSheetView(sheetId).frozenRows).toBe(1);
    expect(document.getSheetView(sheetId).frozenCols).toBe(2);
    expect(document.getSheetView(sheetId).colWidths).toEqual({ "1": 100 });
    expect(document.getSheetView(sheetId).rowHeights).toEqual({ "0": 25 });
    expect(document.getMergedRanges(sheetId)).toEqual([{ startRow: 0, endRow: 1, startCol: 0, endCol: 1 }]);

    // The binder should not eagerly rewrite `view` on initial hydration.
    expect(sheetMap.get("view")).not.toBeInstanceOf(Y.Map);

    // A local update should migrate `view` into a nested Y.Map while preserving unknown keys.
    document.setColWidth(sheetId, 2, 150);

    const viewMap = sheetMap.get("view") as Y.Map<any>;
    expect(viewMap).toBeInstanceOf(Y.Map);
    expect(viewMap.get("customKey")).toEqual({ foo: "bar" });

    const colWidths = viewMap.get("colWidths") as Y.Map<any>;
    const rowHeights = viewMap.get("rowHeights") as Y.Map<any>;
    expect(colWidths).toBeInstanceOf(Y.Map);
    expect(rowHeights).toBeInstanceOf(Y.Map);

    expect(colWidths.get("1")).toBe(100);
    expect(colWidths.get("2")).toBe(150);
    expect(rowHeights.get("0")).toBe(25);

    binder.destroy();
  });
});

