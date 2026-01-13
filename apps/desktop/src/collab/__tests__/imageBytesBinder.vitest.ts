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

