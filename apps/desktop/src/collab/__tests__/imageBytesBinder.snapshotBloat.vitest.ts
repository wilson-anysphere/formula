/**
 * @vitest-environment jsdom
 */

import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { bindImageBytesToCollabSession } from "../imageBytesBinder";
import { DocumentController } from "../../document/documentController.js";
import { IndexedDbImageStore } from "../../drawings/persistence/indexedDbImageStore";
import { DocumentImageStore } from "../../app/spreadsheetApp";

describe("imageBytesBinder + DocumentController imageCache", () => {
  it("hydrates bytes without storing them in DocumentController snapshots", () => {
    const doc = new DocumentController();
    expect(doc.isDirty).toBe(false);

    const persisted = new IndexedDbImageStore(`wb_${Date.now()}_${Math.random().toString(16).slice(2)}`);
    const images = new DocumentImageStore(doc, persisted, { mode: "external", source: "collab" });

    const ydoc = new Y.Doc();
    const metadata = ydoc.getMap("metadata");
    const imagesMap = new Y.Map();
    metadata.set("drawingImages", imagesMap);
    imagesMap.set("img-1", {
      mimeType: "image/png",
      bytesBase64: Buffer.from([1, 2, 3, 4]).toString("base64"),
    });

    const session = { doc: ydoc, metadata, localOrigins: new Set<any>() } as any;
    const binder = bindImageBytesToCollabSession({ session, images });

    expect(doc.getImage("img-1")).toBeTruthy();
    expect(doc.isDirty).toBe(false);

    const snapshot = JSON.parse(new TextDecoder().decode(doc.encodeState()));
    expect(snapshot.images).toBeUndefined();

    binder.destroy();
    images.dispose();
  });
});

