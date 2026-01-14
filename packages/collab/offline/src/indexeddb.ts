import * as Y from "yjs";

import type { OfflinePersistenceHandle } from "./types.ts";

const persistenceOrigin = Symbol("formula.collab-offline.indexeddb");

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

function openDatabase(name: string): Promise<IDBDatabase> {
  const idb: IDBFactory | undefined = (globalThis as any).indexedDB;
  if (!idb?.open) {
    return Promise.reject(new Error("indexedDB.open is unavailable in this environment"));
  }

  return new Promise((resolve, reject) => {
    const request = idb.open(name, 1);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains("updates")) {
        db.createObjectStore("updates", { autoIncrement: true });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("indexedDB.open failed"));
    request.onblocked = () => reject(new Error("indexedDB.open was blocked"));
  });
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
  });
}

async function readAllUpdates(db: IDBDatabase): Promise<Uint8Array[]> {
  const tx = db.transaction("updates", "readonly");
  const store = tx.objectStore("updates");
  const out: Uint8Array[] = [];

  await new Promise<void>((resolve, reject) => {
    const req = store.openCursor();
    req.onsuccess = () => {
      const cursor = req.result;
      if (!cursor) {
        resolve();
        return;
      }

      const value: unknown = cursor.value;
      if (value instanceof Uint8Array) {
        out.push(value);
      } else if (value instanceof ArrayBuffer) {
        out.push(new Uint8Array(value));
      } else if (ArrayBuffer.isView(value)) {
        out.push(new Uint8Array(value.buffer, value.byteOffset, value.byteLength));
      } else if (
        value &&
        typeof value === "object" &&
        "update" in value &&
        (value as any).update instanceof Uint8Array
      ) {
        out.push((value as any).update);
      }

      cursor.continue();
    };
    req.onerror = () => reject(req.error ?? new Error("IndexedDB cursor failed"));
  });

  await transactionDone(tx);
  return out;
}

async function appendUpdate(db: IDBDatabase, update: Uint8Array): Promise<void> {
  const tx = db.transaction("updates", "readwrite");
  tx.objectStore("updates").add(update);
  await transactionDone(tx);
}

async function clearUpdates(db: IDBDatabase): Promise<void> {
  const tx = db.transaction("updates", "readwrite");
  tx.objectStore("updates").clear();
  await transactionDone(tx);
}

export function attachIndexeddbPersistence(doc: Y.Doc, opts: { key: string }): OfflinePersistenceHandle {
  let destroyed = false;
  let isLoading = false;
  const bufferedUpdates: Uint8Array[] = [];
  let db: IDBDatabase | null = null;
  let writeQueue = Promise.resolve();

  const enqueueWrite = (update: Uint8Array) => {
    // IndexedDB writes are best-effort. Avoid unhandled rejections for callers that don't await
    // `whenLoaded()` (or when writes fail after initial load).
    writeQueue = writeQueue
      .catch(() => {
        // keep queue alive
      })
      .then(async () => {
        if (destroyed) return;
        if (!db) return;
        try {
          await appendUpdate(db, update);
        } catch {
          // Best-effort: ignore persistence write failures (private browsing, blocked IDB, etc).
        }
      });
  };

  const updateHandler = (update: Uint8Array, origin: unknown) => {
    if (destroyed) return;
    if (origin === persistenceOrigin) return;
    if (isLoading) {
      bufferedUpdates.push(update);
      return;
    }

    enqueueWrite(update);
  };

  doc.on("update", updateHandler);

  let loadPromise: Promise<void> | null = null;
  const whenLoaded = async () => {
    if (loadPromise) return loadPromise;
    loadPromise = (async () => {
      if (destroyed) return;

      isLoading = true;
      try {
        db = await openDatabase(opts.key);
        const updates = await readAllUpdates(db);
        for (const update of updates) {
          if (destroyed) return;
          Y.applyUpdate(doc, update, persistenceOrigin);
        }

        if (!destroyed && updates.length === 0) {
          try {
            await appendUpdate(db, Y.encodeStateAsUpdate(doc));
            bufferedUpdates.length = 0;
          } catch {
            // Best-effort: failure to write a baseline should not prevent loading.
          }
        }
      } finally {
        isLoading = false;
      }

      if (destroyed) return;

      if (bufferedUpdates.length > 0) {
        for (const update of bufferedUpdates.splice(0, bufferedUpdates.length)) {
          enqueueWrite(update);
        }
      }

      await writeQueue;
    })();

    return loadPromise;
  };

  const destroy = () => {
    if (destroyed) return;
    destroyed = true;
    doc.off("update", updateHandler);
    doc.off("destroy", destroy);
    const toClose = db;
    db = null;
    try {
      toClose?.close();
    } catch {
    }
  };

  const clear = async () => {
    if (destroyed) return;

    const name = opts.key;
    destroy();

    try {
      const tmpDb = await openDatabase(name);
      await clearUpdates(tmpDb);
      tmpDb.close();
      return;
    } catch {
    }

    await deleteDatabase(name);
  };

  doc.on("destroy", destroy);

  return { whenLoaded, destroy, clear };
}
