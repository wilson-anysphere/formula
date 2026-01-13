import type { ImageEntry, ImageStore } from "../types";

type StoredImageRecord = {
  id: string;
  mimeType: string;
  /**
   * Stored as an ArrayBuffer so the structured clone payload is compact and
   * doesn't depend on typed-array subclass support.
   */
  bytes: ArrayBuffer;
};

type IndexedDbImageStoreOptions = {
  /**
   * Database name prefix. The workbook id is appended to avoid cross-workbook
   * collisions.
   */
  dbPrefix?: string;
  storeName?: string;
  version?: number;
};

function copyBytes(bytes: Uint8Array): ArrayBuffer {
  // `Uint8Array.buffer` can include unrelated bytes if the view has an offset.
  // Copy so the persisted payload is exactly the image bytes.
  const out = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(out).set(bytes);
  return out;
}

function requestToPromise<T>(req: IDBRequest<T>): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
  });
}

function transactionDone(tx: IDBTransaction): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
  });
}

/**
 * Persistent image-bytes store backed by IndexedDB, keyed by `imageId`.
 *
 * The database is namespaced per workbook to avoid cross-workbook collisions.
 *
 * NOTE: Synchronous `get`/`set` operations only consult the in-memory cache.
 * Use `getAsync`/`setAsync` for IndexedDB access.
 */
export class IndexedDbImageStore implements ImageStore {
  private readonly cache = new Map<string, ImageEntry>();
  private readonly dbName: string;
  private readonly storeName: string;
  private readonly version: number;
  private dbPromise: Promise<IDBDatabase> | null = null;

  constructor(
    private readonly workbookId: string,
    opts: IndexedDbImageStoreOptions = {},
  ) {
    const prefix = opts.dbPrefix ?? "formula:drawings:images:";
    // Encode to keep the DB name stable even if workbook ids include `/` etc.
    const workbookToken = encodeURIComponent(String(workbookId ?? "").trim() || "local-workbook");
    this.dbName = `${prefix}${workbookToken}`;
    this.storeName = opts.storeName ?? "images";
    this.version = opts.version ?? 1;
  }

  get(id: string): ImageEntry | undefined {
    return this.cache.get(id);
  }

  set(entry: ImageEntry): void {
    this.cache.set(entry.id, entry);
    // Best-effort persistence: failures should never block image insertion.
    void this.setAsync(entry).catch(() => {});
  }

  /**
   * Clear the in-memory cache without touching IndexedDB.
   *
   * Useful for tests and for simulating a reload where only the persistent
   * store remains.
   */
  clearMemory(): void {
    this.cache.clear();
  }

  async getAsync(id: string): Promise<ImageEntry | undefined> {
    const cached = this.cache.get(id);
    if (cached) return cached;

    const db = await this.openDb().catch(() => null);
    if (!db) return undefined;

    try {
      const tx = db.transaction(this.storeName, "readonly");
      const done = transactionDone(tx);
      const store = tx.objectStore(this.storeName);
      const record = await requestToPromise<StoredImageRecord | undefined>(store.get(id));
      await done.catch(() => {});
      if (!record) return undefined;
      const entry: ImageEntry = {
        id: String(record.id),
        mimeType: String(record.mimeType ?? "application/octet-stream"),
        bytes: new Uint8Array(record.bytes),
      };
      this.cache.set(entry.id, entry);
      return entry;
    } catch {
      return undefined;
    }
  }

  async setAsync(entry: ImageEntry): Promise<void> {
    const db = await this.openDb();
    const tx = db.transaction(this.storeName, "readwrite");
    const done = transactionDone(tx);
    const store = tx.objectStore(this.storeName);
    const record: StoredImageRecord = {
      id: entry.id,
      mimeType: entry.mimeType,
      bytes: copyBytes(entry.bytes),
    };
    store.put(record);
    await done;
  }

  async deleteAsync(imageId: string): Promise<void> {
    const db = await this.openDb();
    const tx = db.transaction(this.storeName, "readwrite");
    const done = transactionDone(tx);
    const store = tx.objectStore(this.storeName);
    store.delete(imageId);
    await done;
    this.cache.delete(imageId);
  }

  /**
   * Best-effort garbage collection: remove any records not present in `keep`.
   */
  async garbageCollectAsync(keep: Iterable<string>): Promise<void> {
    const keepSet = new Set(Array.from(keep, (id) => String(id)));
    const db = await this.openDb().catch(() => null);
    if (!db) return;

    try {
      const tx = db.transaction(this.storeName, "readwrite");
      const done = transactionDone(tx);
      const store = tx.objectStore(this.storeName);
      const allKeys = await requestToPromise<IDBValidKey[]>(store.getAllKeys());
      for (const key of allKeys) {
        const id = String(key);
        if (keepSet.has(id)) continue;
        store.delete(key);
        this.cache.delete(id);
      }
      await done;
    } catch {
      // Best-effort.
    }
  }

  private openDb(): Promise<IDBDatabase> {
    if (this.dbPromise) return this.dbPromise;

    const indexedDb = (globalThis as any).indexedDB as IDBFactory | undefined;
    if (!indexedDb) {
      this.dbPromise = Promise.reject(new Error("IndexedDB unavailable"));
      return this.dbPromise;
    }

    this.dbPromise = new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDb.open(this.dbName, this.version);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName, { keyPath: "id" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error("IndexedDB open failed"));
    });

    return this.dbPromise;
  }
}
