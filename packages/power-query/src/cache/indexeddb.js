import { fnv1a64 } from "./key.js";

/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

/**
 * Minimal IndexedDB-backed cache store (browser environments).
 *
 * This implementation is intentionally tiny and only targets the needs of
 * `CacheManager` (get/set/delete/clear). Host applications can replace it with a
 * richer implementation if needed.
 */
export class IndexedDBCacheStore {
  /**
   * @param {{ dbName?: string, storeName?: string }} [options]
   */
  constructor(options = {}) {
    this.dbName = options.dbName ?? "power-query-cache";
    this.storeName = options.storeName ?? "entries";
    /** @type {Promise<IDBDatabase> | null} */
    this.dbPromise = null;
  }

  /**
   * @returns {Promise<IDBDatabase>}
   */
  async open() {
    if (this.dbPromise) return this.dbPromise;
    this.dbPromise = new Promise((resolve, reject) => {
      if (typeof indexedDB === "undefined") {
        reject(new Error("IndexedDB is not available in this environment"));
        return;
      }
      const req = indexedDB.open(this.dbName, 1);
      req.onerror = () => reject(req.error ?? new Error("IndexedDB open failed"));
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName);
        }
      };
      req.onsuccess = () => resolve(req.result);
    });
    return this.dbPromise;
  }

  /**
   * @param {(store: IDBObjectStore) => void} fn
   * @returns {Promise<void>}
   */
  async withWrite(fn) {
    const db = await this.open();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, "readwrite");
      tx.oncomplete = () => resolve(undefined);
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
      fn(tx.objectStore(this.storeName));
    });
  }

  /**
   * @param {(store: IDBObjectStore) => IDBRequest} fn
   * @returns {Promise<any>}
   */
  async withRead(fn) {
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, "readonly");
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
      const req = fn(tx.objectStore(this.storeName));
      req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
      req.onsuccess = () => resolve(req.result);
    });
  }

  /**
   * @param {string} key
   * @returns {string}
   */
  normalizeKey(key) {
    // Keys can be large; normalize to a fixed-size identifier.
    return fnv1a64(key);
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    const normalized = this.normalizeKey(key);
    const result = await this.withRead((store) => store.get(normalized));
    return result ?? null;
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    const normalized = this.normalizeKey(key);
    await this.withWrite((store) => {
      store.put(entry, normalized);
    });
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    const normalized = this.normalizeKey(key);
    await this.withWrite((store) => {
      store.delete(normalized);
    });
  }

  async clear() {
    await this.withWrite((store) => {
      store.clear();
    });
  }

  /**
   * Proactively delete expired entries.
   *
   * CacheManager deletes expired keys on access, but long-lived caches can benefit
   * from an occasional sweep to free storage.
   *
   * @param {number} [nowMs]
   */
  async pruneExpired(nowMs = Date.now()) {
    const db = await this.open();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, "readwrite");
      tx.oncomplete = () => resolve(undefined);
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));

      const store = tx.objectStore(this.storeName);
      const req = store.openCursor();
      req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
      req.onsuccess = () => {
        const cursor = req.result;
        if (!cursor) return;

        const entry = cursor.value;
        const expiresAtMs =
          entry && typeof entry === "object" && "expiresAtMs" in entry
            ? // @ts-ignore - runtime access
              entry.expiresAtMs
            : null;

        if (typeof expiresAtMs === "number" && expiresAtMs <= nowMs) {
          try {
            cursor.delete();
          } catch {
            // ignore delete failures; continue iterating
          }
        }

        cursor.continue();
      };
    });
  }
}
