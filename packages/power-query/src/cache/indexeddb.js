import { fnv1a64 } from "./key.js";

/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

/**
 * @typedef {{
 *   entry: CacheEntry;
 *   sizeBytes: number;
 *   lastAccessMs: number;
 * }} IndexedDBCacheRecord
 */

const SIZE_MARKER_KEY = "__pq_cache_bytes";

/**
 * @param {string} text
 */
function utf8ByteLength(text) {
  if (typeof Buffer !== "undefined") {
    return Buffer.byteLength(text, "utf8");
  }
  if (typeof TextEncoder !== "undefined") {
    return new TextEncoder().encode(text).byteLength;
  }
  // Fallback: assume 1 byte per code unit.
  return text.length;
}

/**
 * Approximate the serialized size of a cache entry.
 *
 * We want something cheap and deterministic (not perfect). The heuristic:
 * - JSON byte length for the non-binary parts
 * - + `byteLength` for any `Uint8Array` payloads
 *
 * @param {CacheEntry} entry
 */
function estimateEntrySizeBytes(entry) {
  let binaryBytes = 0;
  let json = "";
  try {
    json = JSON.stringify(entry, (_key, value) => {
      if (value instanceof Uint8Array) {
        binaryBytes += value.byteLength;
        return { [SIZE_MARKER_KEY]: value.byteLength };
      }
      if (value instanceof ArrayBuffer) {
        binaryBytes += value.byteLength;
        return { [SIZE_MARKER_KEY]: value.byteLength };
      }
      // Node Buffers define a `toJSON()` hook that runs before replacers, so we may
      // see the `{ type: "Buffer", data: number[] }` shape here instead of the
      // original `Uint8Array`. Treat it as binary to avoid JSON bloat.
      if (
        value &&
        typeof value === "object" &&
        !Array.isArray(value) &&
        // @ts-ignore - runtime inspection
        value.type === "Buffer" &&
        // @ts-ignore - runtime inspection
        Array.isArray(value.data)
      ) {
        // @ts-ignore - runtime inspection
        const bytes = Uint8Array.from(value.data);
        binaryBytes += bytes.byteLength;
        return { [SIZE_MARKER_KEY]: bytes.byteLength };
      }
      if (typeof value === "bigint") {
        return value.toString();
      }
      return value;
    });
  } catch {
    json = "";
  }

  return utf8ByteLength(json) + binaryBytes;
}

/**
 * @param {any} value
 * @returns {value is IndexedDBCacheRecord}
 */
function isCacheRecord(value) {
  return (
    value &&
    typeof value === "object" &&
    !Array.isArray(value) &&
    "entry" in value &&
    "sizeBytes" in value &&
    "lastAccessMs" in value
  );
}

/**
 * Minimal IndexedDB-backed cache store (browser environments).
 *
 * This implementation is intentionally tiny and only targets the needs of
 * `CacheManager` (get/set/delete/clear). Host applications can replace it with a
 * richer implementation if needed.
 */
export class IndexedDBCacheStore {
  /**
   * @param {{ dbName?: string, storeName?: string, now?: () => number }} [options]
   */
  constructor(options = {}) {
    this.dbName = options.dbName ?? "power-query-cache";
    this.storeName = options.storeName ?? "entries";
    this.now = options.now ?? (() => Date.now());
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
    const db = await this.open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, "readwrite");
      /** @type {CacheEntry | null} */
      let entryToReturn = null;

      tx.oncomplete = () => resolve(entryToReturn);
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
      tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));

      const store = tx.objectStore(this.storeName);
      const req = store.get(normalized);
      req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
      req.onsuccess = () => {
        const result = req.result;
        if (!result) return;

        /** @type {CacheEntry | null} */
        let entry = null;
        /** @type {number | null} */
        let sizeBytes = null;

        if (isCacheRecord(result)) {
          entry = result.entry ?? null;
          sizeBytes = typeof result.sizeBytes === "number" ? result.sizeBytes : null;
        } else {
          entry = result;
          sizeBytes = null;
        }

        if (!entry) return;
        entryToReturn = entry;

        const updated = {
          entry,
          sizeBytes: sizeBytes ?? estimateEntrySizeBytes(entry),
          lastAccessMs: this.now(),
        };

        // Best-effort metadata update: do not let access-time updates turn reads into failures.
        try {
          const putReq = store.put(updated, normalized);
          putReq.onerror = (event) => {
            if (event && typeof event.preventDefault === "function") event.preventDefault();
          };
        } catch {
          // ignore
        }
      };
    });
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    const normalized = this.normalizeKey(key);
    const record = {
      entry,
      sizeBytes: estimateEntrySizeBytes(entry),
      lastAccessMs: this.now(),
    };
    await this.withWrite((store) => {
      store.put(record, normalized);
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

        const value = cursor.value;
        const entry = isCacheRecord(value) ? value.entry : value;
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

  /**
   * Prune expired entries and enforce optional entry/byte quotas using LRU eviction.
   *
   * @param {{ nowMs: number, maxEntries?: number, maxBytes?: number }} options
   */
  async prune(options) {
    const maxEntries = options.maxEntries;
    const maxBytes = options.maxBytes;

    if (maxEntries == null && maxBytes == null) {
      // Still allow pruning expired entries when no explicit quotas are supplied.
    }

    /** @type {Array<{ id: any, value: any }>} */
    let records = [];
    try {
      const db = await this.open();
      records = await new Promise((resolve, reject) => {
        /** @type {Array<{ id: any, value: any }>} */
        const out = [];
        const tx = db.transaction(this.storeName, "readonly");
        tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
        const store = tx.objectStore(this.storeName);
        const req = store.openCursor();
        req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
        req.onsuccess = () => {
          const cursor = req.result;
          if (!cursor) {
            resolve(out);
            return;
          }
          out.push({ id: cursor.key, value: cursor.value });
          cursor.continue();
        };
      });
    } catch {
      return;
    }

    /** @type {{ id: any, entry: CacheEntry, sizeBytes: number, lastAccessMs: number }[]} */
    const items = [];

    for (const record of records) {
      const id = record.id;
      const value = record.value;

      /** @type {CacheEntry | null} */
      let entry = null;
      /** @type {number | null} */
      let sizeBytes = null;
      /** @type {number | null} */
      let lastAccessMs = null;

      if (isCacheRecord(value)) {
        entry = value.entry ?? null;
        sizeBytes = typeof value.sizeBytes === "number" ? value.sizeBytes : null;
        lastAccessMs = typeof value.lastAccessMs === "number" ? value.lastAccessMs : null;
      } else {
        entry = value;
      }

      if (!entry) continue;

      const normalizedSize = sizeBytes ?? estimateEntrySizeBytes(entry);
      const normalizedLastAccess =
        lastAccessMs ??
        (typeof entry.createdAtMs === "number" ? entry.createdAtMs : 0);

      items.push({ id, entry, sizeBytes: normalizedSize, lastAccessMs: normalizedLastAccess });
    }

    /** @type {Set<any>} */
    const deleteIds = new Set();

    /** @type {{ id: any, entry: CacheEntry, sizeBytes: number, lastAccessMs: number }[]} */
    const remaining = [];
    let totalBytes = 0;

    for (const item of items) {
      if (item.entry.expiresAtMs != null && item.entry.expiresAtMs <= options.nowMs) {
        deleteIds.add(item.id);
      } else {
        remaining.push(item);
        totalBytes += item.sizeBytes;
      }
    }

    let totalEntries = remaining.length;

    remaining.sort((a, b) => {
      if (a.lastAccessMs !== b.lastAccessMs) return a.lastAccessMs - b.lastAccessMs;
      const aCreated = typeof a.entry.createdAtMs === "number" ? a.entry.createdAtMs : 0;
      const bCreated = typeof b.entry.createdAtMs === "number" ? b.entry.createdAtMs : 0;
      if (aCreated !== bCreated) return aCreated - bCreated;
      return String(a.id).localeCompare(String(b.id));
    });

    let idx = 0;
    while (
      (maxEntries != null && totalEntries > maxEntries) ||
      (maxBytes != null && totalBytes > maxBytes)
    ) {
      const victim = remaining[idx++];
      if (!victim) break;
      if (deleteIds.has(victim.id)) continue;
      deleteIds.add(victim.id);
      totalEntries -= 1;
      totalBytes -= victim.sizeBytes;
    }

    if (deleteIds.size === 0) return;

    try {
      await this.withWrite((store) => {
        for (const id of deleteIds) {
          try {
            store.delete(id);
          } catch {
            // ignore per-key failures; best-effort pruning
          }
        }
      });
    } catch {
      // Ignore transaction failures (concurrent prune / blocked transactions).
    }
  }
}
