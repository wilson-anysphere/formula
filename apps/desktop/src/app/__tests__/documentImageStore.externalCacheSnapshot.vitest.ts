/**
 * @vitest-environment jsdom
 */

import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { IndexedDbImageStore } from "../../drawings/persistence/indexedDbImageStore";
import { DocumentImageStore } from "../spreadsheetApp";

describe("DocumentImageStore (external mode) + DocumentController image cache", () => {
  it("hydrates external image bytes without storing them in DocumentController snapshots", () => {
    const doc = new DocumentController();
    const persisted = new IndexedDbImageStore("wb_test");
    const store = new DocumentImageStore(doc, persisted, { mode: "external", source: "collab" });

    let changeCount = 0;
    doc.on("change", () => {
      changeCount += 1;
    });

    store.set({ id: "img1", mimeType: "image/png", bytes: new Uint8Array([1, 2, 3]) });

    expect(changeCount).toBe(1);
    expect(doc.isDirty).toBe(false);

    const snapshot = JSON.parse(new TextDecoder().decode(doc.encodeState()));
    expect(snapshot.images).toBeUndefined();

    const entry = doc.getImage("img1");
    expect(entry).toBeTruthy();
    expect(entry?.mimeType).toBe("image/png");
    expect(Array.from(entry?.bytes ?? [])).toEqual([1, 2, 3]);
  });
});

