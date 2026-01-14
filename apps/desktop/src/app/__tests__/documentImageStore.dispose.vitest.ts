/**
 * @vitest-environment jsdom
 */

import "fake-indexeddb/auto";
import { describe, expect, it } from "vitest";

import type { DocumentController } from "../../document/documentController.js";
import { IndexedDbImageStore } from "../../drawings/persistence/indexedDbImageStore";
import { DocumentImageStore } from "../spreadsheetApp";

describe("DocumentImageStore.dispose", () => {
  it("prevents getAsync from repopulating in-memory caches after dispose", async () => {
    const workbookId = `wb_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const persisted = new IndexedDbImageStore(workbookId);
    const entry = { id: "image_1", mimeType: "image/png", bytes: new Uint8Array([1, 2, 3]) };
    await persisted.setAsync(entry);
    persisted.clearMemory();

    // Provide a minimal DocumentController-like stub (DocumentImageStore uses dynamic access).
    const doc = {
      images: new Map<string, unknown>(),
      imageCache: new Map<string, unknown>(),
    } as unknown as DocumentController;

    const store = new DocumentImageStore(doc, persisted);
    expect(store.get(entry.id)).toBeUndefined();

    const promise = store.getAsync(entry.id);
    store.dispose();

    const loaded = await promise;
    expect(loaded).toBeUndefined();
    // Ensure neither the DocumentImageStore fallback nor the persisted store's memory cache was repopulated.
    expect(store.get(entry.id)).toBeUndefined();
    expect(persisted.get(entry.id)).toBeUndefined();
  });
});

