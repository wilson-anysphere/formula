// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (no resurrection)", () => {
  it("does not recreate deleted sheets when the Yjs sheet list contains a stale entry", () => {
    const ydoc = new Y.Doc();
    const sheets = ydoc.getArray<Y.Map<any>>("sheets");

    const sheet1 = new Y.Map<any>();
    sheet1.set("id", "Sheet1");
    const sheet2 = new Y.Map<any>();
    sheet2.set("id", "Sheet2");
    sheets.push([sheet1, sheet2]);

    const document = new DocumentController();
    document.addSheet({ sheetId: "Sheet1", name: "Sheet1" });
    document.addSheet({ sheetId: "Sheet2", name: "Sheet2" });
    document.deleteSheet("Sheet2");
    expect(document.getSheetIds()).toEqual(["Sheet1"]);
    expect(document.getSheetMeta("Sheet2")).toBeNull();

    const binder = bindSheetViewToCollabSession({
      session: { doc: ydoc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    // Initial hydration should not resurrect the deleted sheet.
    expect(document.getSheetIds()).toEqual(["Sheet1"]);
    expect(document.getSheetMeta("Sheet2")).toBeNull();

    // Remote updates targeting the stale entry should also be ignored.
    ydoc.transact(() => {
      const view = new Y.Map<any>();
      view.set("frozenRows", 2);
      view.set("frozenCols", 1);
      sheet2.set("view", view);
      // Back-compat mirrors.
      sheet2.set("frozenRows", 2);
      sheet2.set("frozenCols", 1);
    });

    expect(document.getSheetIds()).toEqual(["Sheet1"]);
    expect(document.getSheetMeta("Sheet2")).toBeNull();

    binder.destroy();
  });
});

