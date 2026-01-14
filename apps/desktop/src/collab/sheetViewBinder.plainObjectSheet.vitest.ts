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
});

