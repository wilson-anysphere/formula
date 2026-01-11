import type * as Y from "yjs";
import { IndexeddbPersistence } from "y-indexeddb";

import type { OfflinePersistenceHandle } from "./types.ts";

function deleteDatabase(name: string): Promise<void> {
  const idb: IDBFactory | undefined = (globalThis as any).indexedDB;
  if (!idb?.deleteDatabase) {
    return Promise.reject(new Error("indexedDB.deleteDatabase is unavailable in this environment"));
  }

  return new Promise((resolve, reject) => {
    const request = idb.deleteDatabase(name);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error ?? new Error("indexedDB.deleteDatabase failed"));
    request.onblocked = () => reject(new Error("indexedDB.deleteDatabase was blocked"));
  });
}

export function attachIndexeddbPersistence(doc: Y.Doc, opts: { key: string }): OfflinePersistenceHandle {
  const persistence = new IndexeddbPersistence(opts.key, doc) as any;

  let destroyed = false;

  // y-indexeddb uses `whenSynced` for "load complete".
  const whenLoaded = async () => {
    if (destroyed) return;
    await Promise.resolve(persistence.whenSynced);
  };

  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    persistence.destroy?.();
    doc.off("destroy", destroy);
  };

  const clear = async () => {
    if (destroyed) return;

    // Prefer the library's API when available.
    if (typeof persistence.clearData === "function") {
      destroyed = true;
      doc.off("destroy", destroy);
      await persistence.clearData();
      return;
    }
    const ctor: any = IndexeddbPersistence as any;
    if (typeof ctor.clearData === "function") {
      destroyed = true;
      doc.off("destroy", destroy);
      await ctor.clearData(opts.key);
      return;
    }

    // Last resort: delete the database directly.
    // Note: if the connection is still open, deletion can be blocked.
    destroy();
    await deleteDatabase(opts.key);
  };

  doc.on("destroy", destroy);

  return { whenLoaded, destroy, clear };
}
