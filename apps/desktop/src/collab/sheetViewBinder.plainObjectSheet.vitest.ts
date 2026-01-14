// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (plain object sheet entries)", () => {
  it("hydrates from plain-object sheet entries but skips writing back (defensive)", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<any>("sheets");

    const sheetId = "sheet-1";
    doc.transact(() => {
      sheets.push([
        {
          id: sheetId,
          view: { frozenRows: 1, frozenCols: 2 },
        },
      ]);
    });

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    expect(document.getSheetView(sheetId).frozenRows).toBe(1);
    expect(document.getSheetView(sheetId).frozenCols).toBe(2);

    // Local changes should not throw even though the backing sheet entry is a plain object.
    document.setFrozen(sheetId, 3, 4);

    // The binder should skip writing back into the plain-object sheet entry.
    const entry = sheets.get(0);
    expect(entry.view.frozenRows).toBe(1);
    expect(entry.view.frozenCols).toBe(2);

    binder.destroy();
  });

  it("does not crash when DocumentController emits drawingDeltas for a sheet stored as a plain object in Yjs", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<any>("sheets");

    const sheetA = "sheet-1";
    const sheetB = "sheet-2";

    doc.transact(() => {
      sheets.push([
        { id: sheetA, view: { frozenRows: 0, frozenCols: 0 } },
        (() => {
          const map = new Y.Map<any>();
          map.set("id", sheetB);
          return map;
        })(),
      ]);
    });

    const document = new DocumentController();
    document.addSheet({ sheetId: sheetA, name: "Sheet1" });
    document.addSheet({ sheetId: sheetB, name: "Sheet2" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    // Force a non-empty drawingsBySheet entry so deleteSheet emits drawingDeltas.
    (document as any).drawingsBySheet.set(sheetA, [
      {
        id: "drawing-1",
        zOrder: 0,
        kind: { type: "shape", label: "Rect" },
        anchor: { type: "absolute", pos: { xEmu: 0, yEmu: 0 }, size: { cx: 1, cy: 1 } },
      },
    ]);

    expect(() => document.deleteSheet(sheetA)).not.toThrow();

    binder.destroy();
  });
});
