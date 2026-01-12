import { fnv1a64 } from "../cache/key.js";

/**
 * @typedef {import("../table.js").Column} Column
 */

/**
 * @typedef {{
 *   kind?: "memory" | "indexeddb" | "opfs";
 *   dbName?: string;
 *   storeName?: string;
 * }} SpillStoreOptions
 */

/**
 * @typedef {{
 *   key: string;
 *   rows: unknown[][];
 * }} SpillBatchRecord
 */

/**
 * @typedef {{
 *   rowsWritten: number;
 *   batchesWritten: number;
 *   bytesWritten: number;
 * }} SpillStats
 */

/**
 * @typedef {{
 *   putBatch: (key: string, rows: unknown[][]) => Promise<void>;
 *   iterateBatches: (key: string) => AsyncIterable<unknown[][]>;
 *   iterateRows: (key: string) => AsyncIterable<unknown[]>;
 *   clear: (key: string) => Promise<void>;
 *   clearPrefix: (prefix: string) => Promise<void>;
 *   stats: SpillStats;
 * }} SpillStore
 */

/**
 * Approximate byte length of a string.
 * @param {string} text
 */
function utf8ByteLength(text) {
  if (typeof Buffer !== "undefined") {
    return Buffer.byteLength(text, "utf8");
  }
  if (typeof TextEncoder !== "undefined") {
    return new TextEncoder().encode(text).byteLength;
  }
  return text.length;
}

/**
 * Best-effort size estimate for a batch.
 * @param {unknown[][]} rows
 */
function estimateBatchBytes(rows) {
  try {
    return utf8ByteLength(JSON.stringify(rows, (_k, v) => (typeof v === "bigint" ? v.toString() : v)));
  } catch {
    return rows.length * 16;
  }
}

export class MemorySpillStore {
  constructor() {
    /** @type {Map<string, unknown[][][]>} */
    this.batchesByKey = new Map();
    /** @type {SpillStats} */
    this.stats = { rowsWritten: 0, batchesWritten: 0, bytesWritten: 0 };
  }

  /**
   * @param {string} key
   * @param {unknown[][]} rows
   */
  async putBatch(key, rows) {
    const list = this.batchesByKey.get(key) ?? [];
    list.push(rows);
    this.batchesByKey.set(key, list);
    this.stats.rowsWritten += rows.length;
    this.stats.batchesWritten += 1;
    this.stats.bytesWritten += estimateBatchBytes(rows);
  }

  /**
   * @param {string} key
   * @returns {AsyncIterable<unknown[][]>}
   */
  async *iterateBatches(key) {
    const batches = this.batchesByKey.get(key) ?? [];
    for (const batch of batches) {
      yield batch;
    }
  }

  /**
   * @param {string} key
   * @returns {AsyncIterable<unknown[]>}
   */
  async *iterateRows(key) {
    for await (const batch of this.iterateBatches(key)) {
      for (const row of batch) {
        yield row;
      }
    }
  }

  /**
   * @param {string} key
   */
  async clear(key) {
    this.batchesByKey.delete(key);
  }

  /**
   * @param {string} prefix
   */
  async clearPrefix(prefix) {
    for (const key of this.batchesByKey.keys()) {
      if (key.startsWith(prefix)) {
        this.batchesByKey.delete(key);
      }
    }
  }
}

/**
 * Minimal IndexedDB-backed spill store.
 *
 * This intentionally mirrors the "append batches per key" access pattern used by
 * the streaming v2 out-of-core operators.
 */
export class IndexedDbSpillStore {
  /**
   * @param {{ dbName?: string, storeName?: string }} [options]
   */
  constructor(options = {}) {
    this.dbName = options.dbName ?? "power-query-spill";
    this.storeName = options.storeName ?? "spillBatches";
    /** @type {Promise<IDBDatabase> | null} */
    this.dbPromise = null;
    /** @type {SpillStats} */
    this.stats = { rowsWritten: 0, batchesWritten: 0, bytesWritten: 0 };
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
          const store = db.createObjectStore(this.storeName, { keyPath: "id", autoIncrement: true });
          store.createIndex("byKey", "key", { unique: false });
        }
      };
      req.onsuccess = () => resolve(req.result);
    });
    return this.dbPromise;
  }

  /**
   * @param {"readonly" | "readwrite"} mode
   * @param {(store: IDBObjectStore) => IDBRequest | void} fn
   */
  async withStore(mode, fn) {
    const db = await this.open();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(this.storeName, mode);
      tx.oncomplete = () => resolve(undefined);
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
      tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
      const store = tx.objectStore(this.storeName);
      const req = fn(store);
      if (req) {
        req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));
      }
    });
  }

  /**
   * @param {string} key
   * @param {unknown[][]} rows
   */
  async putBatch(key, rows) {
    await this.withStore("readwrite", (store) => store.add({ key, rows }));
    this.stats.rowsWritten += rows.length;
    this.stats.batchesWritten += 1;
    this.stats.bytesWritten += estimateBatchBytes(rows);
  }

  /**
   * @param {string} key
   * @returns {AsyncIterable<unknown[][]>}
   */
  async *iterateBatches(key) {
    const db = await this.open();
    const tx = db.transaction(this.storeName, "readonly");
    const store = tx.objectStore(this.storeName).index("byKey");
    const req = store.openCursor(IDBKeyRange.only(key));

    /**
     * @returns {Promise<IDBCursorWithValue | null>}
     */
    const awaitCursorEvent = async () =>
      new Promise((resolve, reject) => {
        req.onerror = () => reject(req.error ?? new Error("IndexedDB cursor failed"));
        req.onsuccess = () => resolve(req.result);
      });

    let completed = false;
    try {
      let cursor = await awaitCursorEvent();
      while (cursor) {
        const batch = cursor.value?.rows ?? [];
        yield Array.isArray(batch) ? batch : [];
        const next = awaitCursorEvent();
        cursor.continue();
        cursor = await next;
      }
      completed = true;
    } finally {
      if (!completed) {
        // If the consumer stops early, abort the transaction to stop the cursor.
        try {
          tx.abort();
        } catch {
          // ignore
        }
      }
    }
  }

  /**
   * @param {string} key
   * @returns {AsyncIterable<unknown[]>}
   */
  async *iterateRows(key) {
    for await (const batch of this.iterateBatches(key)) {
      for (const row of batch) yield row;
    }
  }

  /**
   * @param {string} key
   */
  async clear(key) {
    await this.clearPrefix(`${key}`);
  }

  /**
   * @param {string} prefix
   */
  async clearPrefix(prefix) {
    const db = await this.open();
    const tx = db.transaction(this.storeName, "readwrite");
    const store = tx.objectStore(this.storeName);
    const index = store.index("byKey");
    const req = index.openCursor();

    await new Promise((resolve, reject) => {
      tx.oncomplete = () => resolve(undefined);
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
      req.onerror = () => reject(req.error ?? new Error("IndexedDB cursor failed"));
      req.onsuccess = () => {
        const cursor = req.result;
        if (!cursor) return;
        const recordKey = cursor.key;
        if (typeof recordKey === "string" && recordKey.startsWith(prefix)) {
          try {
            cursor.delete();
          } catch {
            // ignore delete failures
          }
        }
        cursor.continue();
      };
    });
  }
}

/**
 * Create a spill store instance based on user options + environment capabilities.
 *
 * @param {SpillStoreOptions | undefined} options
 * @returns {SpillStore}
 */
export function createSpillStore(options) {
  const kind = options?.kind;
  if (kind === "memory") return new MemorySpillStore();
  if (kind === "indexeddb") return new IndexedDbSpillStore({ dbName: options?.dbName, storeName: options?.storeName });
  if (kind === "opfs") {
    // OPFS is not implemented in this environment yet; fall back to IndexedDB when available.
  }

  if (typeof indexedDB !== "undefined") {
    return new IndexedDbSpillStore({ dbName: options?.dbName, storeName: options?.storeName });
  }

  return new MemorySpillStore();
}

/**
 * Generate a deterministic-ish temporary key prefix.
 *
 * @param {string} prefix
 * @returns {string}
 */
export function makeSpillKeyPrefix(prefix) {
  const nonce = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  return `${prefix}:${fnv1a64(nonce)}`;
}
