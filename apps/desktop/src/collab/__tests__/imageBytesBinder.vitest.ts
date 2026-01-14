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
    delete(id: string) {
      map.delete(id);
    },
    clear() {
      map.clear();
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

  it("does not remove a provided origin token from session.localOrigins on destroy", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");
    const store = createMemoryImageStore();
    const sharedOrigin = { type: "shared-origin" };
    const localOrigins = new Set<any>([sharedOrigin]);
    const session = { doc, metadata, localOrigins } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store, origin: sharedOrigin });
    binder.destroy();

    expect(localOrigins.has(sharedOrigin)).toBe(true);
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

  it("evicts oldest images when exceeding maxImages", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store, maxImages: 2 });

    const img1: ImageEntry = { id: "img-1", mimeType: "image/png", bytes: new Uint8Array([1]) };
    const img2: ImageEntry = { id: "img-2", mimeType: "image/png", bytes: new Uint8Array([2]) };
    const img3: ImageEntry = { id: "img-3", mimeType: "image/png", bytes: new Uint8Array([3]) };

    binder.onLocalImageInserted(img1);
    binder.onLocalImageInserted(img2);
    binder.onLocalImageInserted(img3);

    const imagesMap = metadata.get("drawingImages") as any;
    expect(imagesMap?.size).toBe(2);
    expect(imagesMap?.get?.("img-1")).toBeUndefined();
    expect(imagesMap?.get?.("img-2")).toBeTruthy();
    expect(imagesMap?.get?.("img-3")).toBeTruthy();

    binder.destroy();
  });

  it("hydrates from a plain-object drawingImages container", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const bytes = new Uint8Array([7, 8, 9]);
    metadata.set("drawingImages", {
      "img-1": {
        mimeType: "image/png",
        bytesBase64: Buffer.from(bytes).toString("base64"),
      },
    });

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store });

    const hydrated = store.get("img-1");
    expect(hydrated).toBeTruthy();
    expect(hydrated?.mimeType).toBe("image/png");
    expect(Array.from(hydrated?.bytes ?? [])).toEqual([7, 8, 9]);

    // Binder should not eagerly rewrite the container just to read it.
    expect(metadata.get("drawingImages")).toEqual({
      "img-1": {
        mimeType: "image/png",
        bytesBase64: Buffer.from(bytes).toString("base64"),
      },
    });

    binder.destroy();
  });

  it("converts a record drawingImages container into a Y.Map when publishing new bytes", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    // Legacy/plain-object container.
    metadata.set("drawingImages", {
      existing: {
        mimeType: "image/png",
        bytesBase64: Buffer.from([1]).toString("base64"),
      },
    });

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;
    const binder = bindImageBytesToCollabSession({ session, images: store });

    const next: ImageEntry = { id: "new", mimeType: "image/png", bytes: new Uint8Array([2, 3]) };
    binder.onLocalImageInserted(next);

    const container = metadata.get("drawingImages");
    expect(container).toBeInstanceOf(Y.Map);
    const imagesMap = container as Y.Map<any>;
    expect(imagesMap.get("existing")).toBeTruthy();
    expect(imagesMap.get("new")).toBeTruthy();

    binder.destroy();
  });

  it("decodes base64url strings (no padding)", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const bytes = new Uint8Array([0xfb, 0xff, 0xff]); // base64 includes "+" and "/"
    const base64 = Buffer.from(bytes).toString("base64");
    const base64url = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");

    metadata.set("drawingImages", {
      "img-1": {
        mimeType: "image/png",
        bytesBase64: base64url,
      },
    });

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store });

    const hydrated = store.get("img-1");
    expect(hydrated).toBeTruthy();
    expect(Array.from(hydrated?.bytes ?? [])).toEqual(Array.from(bytes));

    binder.destroy();
  });

  it("ignores invalid base64 payloads without throwing", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    // Invalid base64: length % 4 === 1 is never valid.
    metadata.set("drawingImages", {
      "img-1": {
        mimeType: "image/png",
        bytesBase64: "a",
      },
    });

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store });

    expect(store.get("img-1")).toBeUndefined();

    binder.destroy();
  });

  it("enforces maxImageBytes when hydrating remote base64 payloads", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    metadata.set("drawingImages", {
      "img-1": {
        mimeType: "image/png",
        bytesBase64: Buffer.from([1, 2, 3]).toString("base64"), // 3 bytes
      },
    });

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store, maxImageBytes: 2 });

    expect(store.get("img-1")).toBeUndefined();

    binder.destroy();
  });

  it("hydrates from a nested Y.Map per image id entry", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const bytes = new Uint8Array([10, 20, 30]);
    const imagesMap = new Y.Map();
    metadata.set("drawingImages", imagesMap);

    const entryMap = new Y.Map();
    entryMap.set("mimeType", "image/png");
    entryMap.set("bytesBase64", Buffer.from(bytes).toString("base64"));
    imagesMap.set("img-1", entryMap);

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store });

    const hydrated = store.get("img-1");
    expect(hydrated).toBeTruthy();
    expect(hydrated?.mimeType).toBe("image/png");
    expect(Array.from(hydrated?.bytes ?? [])).toEqual([10, 20, 30]);

    binder.destroy();
  });

  it("hydrates from a direct Uint8Array payload (legacy variant)", () => {
    const doc = new Y.Doc();
    const metadata = doc.getMap("metadata");

    const imagesMap = new Y.Map();
    metadata.set("drawingImages", imagesMap);
    imagesMap.set("img-1", new Uint8Array([1, 2, 3]));

    const store = createMemoryImageStore();
    const session = { doc, metadata, localOrigins: new Set<any>() } as any;

    const binder = bindImageBytesToCollabSession({ session, images: store });

    const hydrated = store.get("img-1");
    expect(hydrated).toBeTruthy();
    expect(hydrated?.mimeType).toBe("application/octet-stream");
    expect(Array.from(hydrated?.bytes ?? [])).toEqual([1, 2, 3]);

    binder.destroy();
  });
});  
