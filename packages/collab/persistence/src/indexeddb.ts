import * as Y from "yjs";
import { IndexeddbPersistence } from "y-indexeddb";

import type { CollabPersistence, CollabPersistenceBinding } from "./index.js";

type Entry = {
  doc: Y.Doc;
  persistence: IndexeddbPersistence;
};

/**
 * Browser (IndexedDB) persistence using `y-indexeddb`.
 *
 * In tests, this works with `fake-indexeddb` by installing the IndexedDB globals
 * (`globalThis.indexedDB`, `globalThis.IDBKeyRange`, ...).
 */
export class IndexedDbCollabPersistence implements CollabPersistence {
  private readonly entries = new Map<string, Entry>();

  private getOrCreate(docId: string, doc: Y.Doc): IndexeddbPersistence {
    const existing = this.entries.get(docId);
    if (existing) {
      if (existing.doc === doc) return existing.persistence;
      existing.persistence.destroy();
      this.entries.delete(docId);
    }

    const persistence = new IndexeddbPersistence(docId, doc);
    this.entries.set(docId, { doc, persistence });
    return persistence;
  }

  async load(docId: string, doc: Y.Doc): Promise<void> {
    const persistence = this.getOrCreate(docId, doc);
    await persistence.whenSynced;
  }

  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding {
    const persistence = this.getOrCreate(docId, doc);
    return {
      destroy: async () => {
        persistence.destroy();
        this.entries.delete(docId);
      },
    };
  }

  async clear(docId: string): Promise<void> {
    const entry = this.entries.get(docId);
    if (entry) {
      await entry.persistence.clearData();
      entry.persistence.destroy();
      this.entries.delete(docId);
      return;
    }

    const tmpDoc = new Y.Doc();
    const tmp = new IndexeddbPersistence(docId, tmpDoc);
    await tmp.clearData();
    tmp.destroy();
    tmpDoc.destroy();
  }
}

