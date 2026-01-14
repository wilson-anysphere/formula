// @vitest-environment jsdom
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { DocumentController } from "../document/documentController.js";
import { bindSheetViewToCollabSession } from "./sheetViewBinder";

describe("bindSheetViewToCollabSession (backgroundImageId)", () => {
  it("syncs local background image updates into Yjs + applies remote updates back into DocumentController", () => {
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

    document.setSheetBackgroundImageId(sheetId, "bg.png");

    const viewMap = sheetMap.get("view") as Y.Map<any>;
    expect(viewMap).toBeInstanceOf(Y.Map);
    expect(viewMap.get("backgroundImageId")).toBe("bg.png");
    // Back-compat mirror.
    expect(sheetMap.get("backgroundImageId")).toBe("bg.png");

    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      remoteViewMap.set("backgroundImageId", "bg2.png");
      sheetMap.set("backgroundImageId", "bg2.png");
    });

    expect(document.getSheetBackgroundImageId(sheetId)).toBe("bg2.png");

    doc.transact(() => {
      const remoteViewMap = sheetMap.get("view") as Y.Map<any>;
      remoteViewMap.delete("backgroundImageId");
      sheetMap.delete("backgroundImageId");
    });

    expect(document.getSheetBackgroundImageId(sheetId)).toBe(null);

    binder.destroy();
  });
});

