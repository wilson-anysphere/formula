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

  it("hydrates backgroundImageId stored as Y.Text (mixed/legacy encoding)", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheets.push([sheetMap]);

    const yText = new Y.Text();
    yText.insert(0, "bg-ytext.png");
    doc.transact(() => {
      const view = new Y.Map<any>();
      view.set("backgroundImageId", yText);
      sheetMap.set("view", view);
    });

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    expect(document.getSheetBackgroundImageId(sheetId)).toBe("bg-ytext.png");

    binder.destroy();
  });

  it("hydrates backgroundImageId stored under legacy background_image key and reacts to nested view updates", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    const viewMap = new Y.Map<any>();
    sheetMap.set("view", viewMap);
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins: new Set(), isReadOnly: () => false } as any,
      documentController: document,
    });

    doc.transact(() => {
      viewMap.set("background_image", "bg-legacy.png");
    });

    expect(document.getSheetBackgroundImageId(sheetId)).toBe("bg-legacy.png");

    binder.destroy();
  });

  it("does not remove a provided origin token from session.localOrigins on destroy", () => {
    const doc = new Y.Doc();
    const sheets = doc.getArray<Y.Map<any>>("sheets");

    const sheetId = "sheet-1";
    const sheetMap = new Y.Map<any>();
    sheetMap.set("id", sheetId);
    sheets.push([sheetMap]);

    const document = new DocumentController();
    document.addSheet({ sheetId, name: "Sheet1" });

    const sharedOrigin = { type: "shared-origin" };
    const localOrigins = new Set<any>([sharedOrigin]);
    const binder = bindSheetViewToCollabSession({
      session: { doc, sheets, localOrigins, isReadOnly: () => false } as any,
      documentController: document,
      origin: sharedOrigin,
    });

    binder.destroy();

    expect(localOrigins.has(sharedOrigin)).toBe(true);
  });
});
