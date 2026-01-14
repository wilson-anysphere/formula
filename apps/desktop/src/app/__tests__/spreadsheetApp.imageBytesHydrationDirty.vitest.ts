/**
 * @vitest-environment jsdom
 */

import { describe, expect, it } from "vitest";

import { DocumentController } from "../../document/documentController.js";
import { IndexedDbImageStore } from "../../drawings/persistence/indexedDbImageStore";
import { DocumentImageStore } from "../spreadsheetApp";

describe("DocumentImageStore external image hydration", () => {
  it("does not mark DocumentController dirty when collab image bytes are written", () => {
    const doc = new DocumentController();
    doc.markSaved();
    expect(doc.isDirty).toBe(false);

    const persisted = new IndexedDbImageStore("test-workbook");
    const store = new DocumentImageStore(doc, persisted, { mode: "external", source: "collab" });

    const id = "image-1.png";
    const bytes = new Uint8Array([1, 2, 3, 4]);
    const mimeType = "image/png";

    store.set({ id, bytes, mimeType });

    expect(doc.isDirty).toBe(false);
    expect(doc.getImage(id)).toEqual({ bytes, mimeType });
  });
});

