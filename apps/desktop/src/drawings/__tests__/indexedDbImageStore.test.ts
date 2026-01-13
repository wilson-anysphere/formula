/**
 * @vitest-environment jsdom
 */

import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";

import { IndexedDbImageStore } from "../persistence/indexedDbImageStore";

describe("IndexedDbImageStore", () => {
  it("setAsync then getAsync returns the same bytes", async () => {
    const workbookId = `wb_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const store = new IndexedDbImageStore(workbookId);

    const entry = {
      id: "image_1",
      mimeType: "image/png",
      bytes: new Uint8Array([1, 2, 3, 4, 5]),
    };

    await store.setAsync(entry);

    // Ensure we're reading from IndexedDB (not just the in-memory cache).
    store.clearMemory();
    expect(store.get(entry.id)).toBeUndefined();

    const loaded = await store.getAsync(entry.id);
    expect(loaded).toBeTruthy();
    expect(loaded?.mimeType).toBe("image/png");
    expect(Array.from(loaded!.bytes)).toEqual(Array.from(entry.bytes));
  });
});

