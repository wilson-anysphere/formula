import { describe, expect, it } from "vitest";
import * as Y from "yjs";

import { bindImageBytesToCollabSession } from "../imageBytesBinder";
import type { ImageEntry, ImageStore } from "../../drawings/types";

function createMemoryImageStore(): ImageStore & { map: Map<string, ImageEntry> } {
  const map = new Map<string, ImageEntry>();
  return {
    map,
    get(id: string) {
      return map.get(id);
    },
    set(entry: ImageEntry) {
      map.set(entry.id, entry);
    },
  };
}

describe("imageBytesBinder", () => {
  it("propagates image bytes between collaborators via Yjs metadata", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const storeA = createMemoryImageStore();
    const storeB = createMemoryImageStore();

    const sessionA = { doc, metadata, localOrigins: new Set<any>() } as any;
    const sessionB = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binderA = bindImageBytesToCollabSession({ session: sessionA, images: storeA });
    const binderB = bindImageBytesToCollabSession({ session: sessionB, images: storeB });

    // Binder should be read-only on startup and avoid creating empty Yjs roots.
    expect(metadata.get("drawingImages")).toBeUndefined();

    const image: ImageEntry = { id: "img-1", mimeType: "image/png", bytes: new Uint8Array([1, 2, 3, 4]) };
    storeA.set(image);
    binderA.onLocalImageInserted(image);

    const received = storeB.get("img-1");
    expect(received).toBeTruthy();
    expect(received?.mimeType).toBe("image/png");
    expect(Array.from(received?.bytes ?? [])).toEqual([1, 2, 3, 4]);

    binderA.destroy();
    binderB.destroy();
  });

  it("is idempotent and does not re-hydrate on unrelated metadata changes", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const storeA = createMemoryImageStore();
    const storeB = createMemoryImageStore();
    let storeBSets = 0;
    storeB.set = (entry: ImageEntry) => {
      storeBSets += 1;
      storeB.map.set(entry.id, entry);
    };

    const sessionA = { doc, metadata, localOrigins: new Set<any>() } as any;
    const sessionB = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binderA = bindImageBytesToCollabSession({ session: sessionA, images: storeA });
    const binderB = bindImageBytesToCollabSession({ session: sessionB, images: storeB });

    const image: ImageEntry = { id: "img-1", mimeType: "image/png", bytes: new Uint8Array([1, 2, 3, 4]) };
    binderA.onLocalImageInserted(image);
    expect(storeBSets).toBe(1);

    doc.transact(() => {
      metadata.set("unrelatedKey", "value");
    });

    // Unrelated metadata changes should not cause repeated base64 decoding / ImageStore writes.
    expect(storeBSets).toBe(1);

    binderA.destroy();
    binderB.destroy();
  });

  it("skips base64 decode/write when bytes already exist locally", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    // Seed the collaborative drawingImages map with one entry.
    const imagesMap = new Y.Map();
    metadata.set("drawingImages", imagesMap);
    imagesMap.set("img-1", {
      mimeType: "image/png",
      bytesBase64: Buffer.from([1, 2, 3, 4]).toString("base64"),
    });

    const store = createMemoryImageStore();
    // Simulate a local persistence layer having already populated bytes.
    store.map.set("img-1", { id: "img-1", mimeType: "image/png", bytes: new Uint8Array([1, 2, 3, 4]) });
    let sets = 0;
    store.set = () => {
      sets += 1;
    };

    const session = { doc, metadata, localOrigins: new Set<any>() } as any;
    const binder = bindImageBytesToCollabSession({ session, images: store });

    expect(sets).toBe(0);

    binder.destroy();
  });

  it("does not propagate oversized images", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const storeA = createMemoryImageStore();
    const storeB = createMemoryImageStore();

    const sessionA = { doc, metadata, localOrigins: new Set<any>() } as any;
    const sessionB = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binderA = bindImageBytesToCollabSession({ session: sessionA, images: storeA });
    const binderB = bindImageBytesToCollabSession({ session: sessionB, images: storeB });

    const bytes = new Uint8Array(1_000_000 + 1);
    const image: ImageEntry = { id: "img-big", mimeType: "image/png", bytes };
    storeA.set(image);
    binderA.onLocalImageInserted(image);

    expect(storeB.get("img-big")).toBeUndefined();

    const drawingImages = metadata.get("drawingImages") as any;
    expect(drawingImages?.get?.("img-big")).toBeUndefined();

    binderA.destroy();
    binderB.destroy();
  });
});
