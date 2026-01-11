import * as Y from "yjs";
import { IndexeddbPersistence } from "y-indexeddb";

import type { CollabPersistence, CollabPersistenceBinding } from "./index.js";

type Entry = {
  doc: Y.Doc;
  persistence: IndexeddbPersistence;
  destroyed: Promise<void>;
  resolveDestroyed: () => void;
  onDocDestroy: () => void;
};

/**
 * Browser (IndexedDB) persistence using `y-indexeddb`.
 *
 * In tests, this works with `fake-indexeddb` by installing the IndexedDB globals
 * (`globalThis.indexedDB`, `globalThis.IDBKeyRange`, ...).
 */
export class IndexedDbCollabPersistence implements CollabPersistence {
  private readonly entries = new Map<string, Entry>();

  private destroyEntry(docId: string, entry: Entry): void {
    entry.resolveDestroyed();
    entry.doc.off("destroy", entry.onDocDestroy);
    entry.persistence.destroy();
    this.entries.delete(docId);
  }

  private getOrCreateEntry(docId: string, doc: Y.Doc): Entry {
    const existing = this.entries.get(docId);
    if (existing) {
      if (existing.doc === doc) return existing;
      this.destroyEntry(docId, existing);
    }

    const persistence = new IndexeddbPersistence(docId, doc);
    let resolveDestroyed: () => void = () => {};
    const destroyed = new Promise<void>((resolve) => {
      resolveDestroyed = resolve;
    });

    const entry: Entry = {
      doc,
      persistence,
      destroyed,
      resolveDestroyed,
      onDocDestroy: () => {},
    };

    entry.onDocDestroy = () => {
      // y-indexeddb's `whenSynced` promise can hang forever if the persistence
      // instance is destroyed before the initial sync completes. Ensure any
      // pending `load()` calls are unblocked when the Y.Doc lifecycle ends.
      this.destroyEntry(docId, entry);
    };
    doc.on("destroy", entry.onDocDestroy);

    this.entries.set(docId, entry);
    return entry;
  }

  async load(docId: string, doc: Y.Doc): Promise<void> {
    const entry = this.getOrCreateEntry(docId, doc);
    await Promise.race([Promise.resolve(entry.persistence.whenSynced), entry.destroyed]);
  }

  bind(docId: string, doc: Y.Doc): CollabPersistenceBinding {
    const entry = this.getOrCreateEntry(docId, doc);
    return {
      destroy: async () => {
        this.destroyEntry(docId, entry);
      },
    };
  }

  async clear(docId: string): Promise<void> {
    const entry = this.entries.get(docId);
    if (entry) {
      entry.resolveDestroyed();
      await entry.persistence.clearData();
      this.destroyEntry(docId, entry);
      return;
    }

    const tmpDoc = new Y.Doc();
    const tmp = new IndexeddbPersistence(docId, tmpDoc);
    await tmp.clearData();
    tmp.destroy();
    tmpDoc.destroy();
  }
}
