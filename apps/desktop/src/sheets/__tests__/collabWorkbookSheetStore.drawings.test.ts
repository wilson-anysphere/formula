import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { CollabWorkbookSheetStore, listSheetsFromCollabSession } from "../collabWorkbookSheetStore";

describe("CollabWorkbookSheetStore (drawings id hardening)", () => {
  it("filters oversized drawing ids when cloning sheet metadata during move()", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    doc.transact(() => {
      const sheet1 = new Y.Map();
      sheet1.set("id", "Sheet1");
      sheet1.set("name", "Sheet1");
      sheet1.set("view", {
        drawings: [
          { id: "x".repeat(5000), zOrder: 0 },
          { id: "  ok  ", zOrder: 1 },
          { id: 1, zOrder: 2 },
        ],
      });

      const sheet2 = new Y.Map();
      sheet2.set("id", "Sheet2");
      sheet2.set("name", "Sheet2");

      sheets.push([sheet1, sheet2]);
    });

    const session = {
      sheets,
      transactLocal: (fn: () => void) => doc.transact(fn),
    };

    const keyRef = { value: "" };
    const store = new CollabWorkbookSheetStore(
      session as any,
      listSheetsFromCollabSession(session as any),
      keyRef,
    );

    // Moving a sheet clones the sheet entry map; the clone should not include pathological ids.
    store.move("Sheet1", 1);

    const moved = sheets.get(1) as any;
    expect(moved?.get?.("id")).toBe("Sheet1");
    const view = moved?.get?.("view") as any;
    expect(view && typeof view === "object").toBe(true);
    expect(Array.isArray(view.drawings)).toBe(true);
    expect(view.drawings.map((d: any) => d.id)).toEqual(["ok", 1]);
  });
});

