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

const UPDATES_STORE_NAME = "updates";

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

  /**
   * Best-effort durability flush for IndexedDB persistence.
   *
   * `y-indexeddb` writes updates to IndexedDB asynchronously and does not expose a
   * built-in API to await the completion of those writes. For Formula's collab
   * persistence contract (`CollabPersistence.flush`), we implement flush by
   * storing a full document snapshot update into the same `updates` object store
   * used by `y-indexeddb`.
   *
   * This guarantees that all in-memory state at the time of the call is durable
   * (even if some incremental updates are still in flight).
   */
  async flush(docId: string): Promise<void> {
    const entry = this.entries.get(docId);
    if (!entry) return;

    const ready = await Promise.race([
      Promise.resolve(entry.persistence.whenSynced).then(() => "synced" as const),
      entry.destroyed.then(() => "destroyed" as const),
    ]);
    if (ready === "destroyed") return;

    const db = entry.persistence.db;
    if (!db) return;

    const snapshot = Y.encodeStateAsUpdate(entry.doc);

    await new Promise<void>((resolve, reject) => {
      try {
        const tx = (db as any).transaction([UPDATES_STORE_NAME], "readwrite");
        const finishError = () => reject((tx as any).error ?? new Error("IndexedDB flush transaction failed"));
        const finishOk = () => resolve();

        // Prefer EventTarget listeners when available.
        if (typeof tx?.addEventListener === "function") {
          tx.addEventListener("complete", finishOk, { once: true });
          tx.addEventListener("error", finishError, { once: true });
          tx.addEventListener("abort", finishError, { once: true });
        } else {
          (tx as any).oncomplete = finishOk;
          (tx as any).onerror = finishError;
          (tx as any).onabort = finishError;
        }

        const store = (tx as any).objectStore(UPDATES_STORE_NAME);
        store.add(snapshot);
      } catch (err) {
        reject(err);
      }
    });
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
