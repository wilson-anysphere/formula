/**
 * @typedef {Object} BinaryStorage
 * @property {() => Promise<Uint8Array | null>} load
 * @property {(data: Uint8Array) => Promise<void>} save
 */

export class InMemoryBinaryStorage {
  constructor() {
    /** @type {Uint8Array | null} */
    this._data = null;
  }

  async load() {
    return this._data ? new Uint8Array(this._data) : null;
  }

  async save(data) {
    this._data = new Uint8Array(data);
  }
}

export class LocalStorageBinaryStorage {
  /**
   * @param {{ workbookId: string, namespace?: string }} opts
   */
  constructor(opts) {
    if (!opts?.workbookId) throw new Error("LocalStorageBinaryStorage requires workbookId");
    const namespace = opts.namespace ?? "formula.ai-rag.sqlite";
    this.key = `${namespace}:${opts.workbookId}`;
  }

  async load() {
    const storage = getLocalStorageOrNull();
    if (!storage) return null;
    const encoded = storage.getItem(this.key);
    if (!encoded) return null;
    return fromBase64(encoded);
  }

  async save(data) {
    const storage = getLocalStorageOrNull();
    if (!storage) return;
    storage.setItem(this.key, toBase64(data));
  }
}

export class ChunkedLocalStorageBinaryStorage {
  /**
   * @param {{ workbookId: string, namespace?: string, chunkSizeChars?: number }} opts
   */
  constructor(opts) {
    if (!opts?.workbookId) throw new Error("ChunkedLocalStorageBinaryStorage requires workbookId");
    const namespace = opts.namespace ?? "formula.ai-rag.sqlite";
    this.key = `${namespace}:${opts.workbookId}`;
    this.metaKey = `${this.key}:meta`;
    this.chunkSizeChars = Math.max(1, Math.floor(opts.chunkSizeChars ?? 1_000_000));
  }

  async load() {
    const storage = getLocalStorageOrNull();
    if (!storage) return null;

    const metaRaw = storage.getItem(this.metaKey);
    if (!metaRaw) return null;

    /** @type {{ chunks?: number } | null} */
    let meta = null;
    try {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
      meta = JSON.parse(metaRaw);
    } catch {
      return null;
    }

    const chunks = meta?.chunks;
    if (!Number.isInteger(chunks) || chunks < 0) return null;

    /** @type {string[]} */
    const parts = [];
    for (let i = 0; i < chunks; i += 1) {
      const part = storage.getItem(`${this.key}:${i}`);
      if (typeof part !== "string") return null;
      parts.push(part);
    }

    return fromBase64(parts.join(""));
  }

  /**
   * @param {Uint8Array} data
   */
  async save(data) {
    const storage = getLocalStorageOrNull();
    if (!storage) return;

    const encoded = toBase64(data);
    const chunks = encoded.length === 0 ? 0 : Math.ceil(encoded.length / this.chunkSizeChars);

    for (let i = 0; i < chunks; i += 1) {
      const start = i * this.chunkSizeChars;
      storage.setItem(`${this.key}:${i}`, encoded.slice(start, start + this.chunkSizeChars));
    }

    storage.setItem(this.metaKey, JSON.stringify({ chunks }));

    // Remove leftover chunks from a prior (larger) save. We avoid relying solely on
    // stored metadata so we can clean up even if `:meta` was missing/corrupted.
    const prefix = `${this.key}:`;
    /** @type {string[]} */
    const keysToRemove = [];
    for (let i = 0; i < storage.length; i += 1) {
      const key = storage.key(i);
      if (!key || !key.startsWith(prefix) || key === this.metaKey) continue;
      const suffix = key.slice(prefix.length);
      if (!/^\d+$/.test(suffix)) continue;
      const index = Number(suffix);
      if (Number.isInteger(index) && index >= chunks) keysToRemove.push(key);
    }
    for (const key of keysToRemove) storage.removeItem(key);
  }
}

export class IndexedDBBinaryStorage {
  /**
   * IndexedDB-backed persistence for binary payloads (e.g. sql.js exports).
   *
   * Notes:
   * - Uses a single object store keyed by `${namespace}:${workbookId}`.
   * - Falls back gracefully when `indexedDB` is unavailable (e.g. Node, restricted
   *   browser contexts):
   *   - `load()` returns null
   *   - `save()` is a no-op
   *
   * @param {{ workbookId: string, namespace?: string, dbName?: string }} opts
   */
  constructor(opts) {
    if (!opts?.workbookId) throw new Error("IndexedDBBinaryStorage requires workbookId");
    this.namespace = opts.namespace ?? "formula.ai-rag.sqlite";
    this.dbName = opts.dbName ?? "formula.ai-rag.binary-storage";
    this.key = `${this.namespace}:${opts.workbookId}`;
    this.storeName = "binary";
    /** @type {Promise<IDBDatabase> | null} */
    this._dbPromise = null;
  }

  async load() {
    const db = await this._openOrNull();
    if (!db) return null;
    try {
      const value = await idbTransaction(db, this.storeName, "readonly", (store) => store.get(this.key));
      if (value == null) return null;
      return await normalizeBinary(value);
    } catch {
      return null;
    }
  }

  /**
   * @param {Uint8Array} data
   */
  async save(data) {
    const db = await this._openOrNull();
    if (!db) return;
    try {
      // Store a tight ArrayBuffer slice to avoid persisting unrelated bytes when `data` is a view.
      const buffer = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
      await idbWrite(db, this.storeName, (store) => store.put(buffer, this.key));
    } catch {
      // Best-effort persistence. If IndexedDB writes fail (quota / permissions),
      // we intentionally do not throw to keep callers usable in restricted contexts.
    }
  }

  async _openOrNull() {
    const idb = getIndexedDBOrNull();
    if (!idb) return null;
    if (this._dbPromise) return this._dbPromise.catch(() => null);

    this._dbPromise = new Promise((resolve, reject) => {
      const request = idb.open(this.dbName, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        if (!db.objectStoreNames.contains(this.storeName)) {
          db.createObjectStore(this.storeName);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error("IndexedDB open failed"));
      request.onblocked = () => reject(new Error("IndexedDB open blocked"));
    });

    try {
      return await this._dbPromise;
    } catch {
      return null;
    }
  }
}

/**
 * localStorage is not always available:
 * - Node >=25 exposes an experimental `globalThis.localStorage` that throws unless
 *   started with `--localstorage-file`.
 * - Vitest's jsdom environment stores the real DOM window on `globalThis.jsdom.window`.
 */
function getLocalStorageOrNull() {
  const isStorage = (value) => value && typeof value.getItem === "function" && typeof value.setItem === "function";

  try {
    // Prefer Vitest's jsdom window when present.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const jsdomStorage = globalThis?.jsdom?.window?.localStorage;
    if (isStorage(jsdomStorage)) return jsdomStorage;
  } catch {
    // ignore
  }

  try {
    // eslint-disable-next-line no-undef
    const windowStorage = typeof window !== "undefined" ? window.localStorage : undefined;
    if (isStorage(windowStorage)) return windowStorage;
  } catch {
    // ignore
  }

  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const storage = globalThis?.localStorage;
    return isStorage(storage) ? storage : null;
  } catch {
    return null;
  }
}

/**
 * @param {Uint8Array} data
 */
export function toBase64(data) {
  // Prefer Node's Buffer when available.
  if (typeof Buffer !== "undefined") {
    return Buffer.from(data).toString("base64");
  }

  // Browser fallback.
  // Avoid byte-by-byte string concatenation (O(n^2) in many JS engines) by
  // building the binary string in reasonably sized chunks.
  const chunkSize = 0x8000;
  /** @type {string[]} */
  const chunks = [];
  for (let i = 0; i < data.length; i += chunkSize) {
    const chunk = data.subarray(i, i + chunkSize);
    // `String.fromCharCode` expects a list of numbers. Passing a TypedArray via
    // `apply` keeps this browser-safe while avoiding large argument lists.
    // eslint-disable-next-line prefer-spread
    chunks.push(String.fromCharCode.apply(null, chunk));
  }
  // eslint-disable-next-line no-undef
  return btoa(chunks.join(""));
}

/**
 * @param {string} encoded
 */
export function fromBase64(encoded) {
  if (typeof Buffer !== "undefined") {
    return new Uint8Array(Buffer.from(encoded, "base64"));
  }
  // eslint-disable-next-line no-undef
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  // Fill in chunks to keep the hot loop simple for large payloads.
  const chunkSize = 0x8000;
  for (let i = 0; i < binary.length; i += chunkSize) {
    const end = Math.min(i + chunkSize, binary.length);
    for (let j = i; j < end; j += 1) {
      bytes[j] = binary.charCodeAt(j);
    }
  }
  return bytes;
}

function getIndexedDBOrNull() {
  try {
    // eslint-disable-next-line no-undef
    const idb = typeof indexedDB !== "undefined" ? indexedDB : undefined;
    if (idb && typeof idb.open === "function") return idb;
  } catch {
    // ignore
  }
  try {
    const idb = globalThis?.indexedDB;
    if (idb && typeof idb.open === "function") return idb;
  } catch {
    // ignore
  }
  return null;
}

/**
 * @template T
 * @param {IDBDatabase} db
 * @param {string} storeName
 * @param {"readonly" | "readwrite"} mode
 * @param {(store: IDBObjectStore) => IDBRequest<T>} requestFn
 * @returns {Promise<T>}
 */
function idbTransaction(db, storeName, mode, requestFn) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const store = tx.objectStore(storeName);
    /** @type {T} */
    let value;
    const req = requestFn(store);

    req.onsuccess = () => {
      value = req.result;
    };
    req.onerror = () => reject(req.error ?? new Error("IndexedDB request failed"));

    tx.oncomplete = () => resolve(value);
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
  });
}

/**
 * @param {IDBDatabase} db
 * @param {string} storeName
 * @param {(store: IDBObjectStore) => void} fn
 * @returns {Promise<void>}
 */
function idbWrite(db, storeName, fn) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, "readwrite");
    tx.oncomplete = () => resolve();
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed"));
    fn(tx.objectStore(storeName));
  });
}

/**
 * @param {any} value
 * @returns {Promise<Uint8Array>}
 */
async function normalizeBinary(value) {
  if (value instanceof Uint8Array) return new Uint8Array(value);
  if (value instanceof ArrayBuffer) return new Uint8Array(value.slice(0));
  if (value && typeof value === "object" && value.buffer instanceof ArrayBuffer) {
    // Structured clone can yield `{ buffer, byteOffset, byteLength }` shapes for typed arrays.
    const offset = typeof value.byteOffset === "number" ? value.byteOffset : 0;
    const length =
      typeof value.byteLength === "number"
        ? value.byteLength
        : typeof value.length === "number"
          ? value.length
          : value.buffer.byteLength;
    const view = new Uint8Array(value.buffer, offset, length);
    return new Uint8Array(view);
  }
  if (typeof Blob !== "undefined" && value instanceof Blob) {
    const buffer = await value.arrayBuffer();
    return new Uint8Array(buffer);
  }
  throw new Error("Invalid binary blob in IndexedDB record");
}
