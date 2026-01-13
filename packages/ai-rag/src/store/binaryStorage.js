/**
 * @typedef {Object} BinaryStorage
 * @property {() => Promise<Uint8Array | null>} load
 * @property {(data: Uint8Array) => Promise<void>} save
 * @property {() => Promise<void>} [remove]
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

  async remove() {
    this._data = null;
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
    try {
      return fromBase64(encoded);
    } catch {
      // Corrupted base64 payload; clear it so future loads can recover.
      try {
        storage.removeItem?.(this.key);
      } catch {
        // ignore
      }
      return null;
    }
  }

  async save(data) {
    const storage = getLocalStorageOrNull();
    if (!storage) return;
    storage.setItem(this.key, toBase64(data));
  }

  async remove() {
    const storage = getLocalStorageOrNull();
    if (!storage) return;
    storage.removeItem?.(this.key);
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
    // Base64 is encoded in 4-character quanta. Normalize to a multiple of 4 so
    // each stored chunk is independently decodable.
    const rawChunkSizeChars = Math.max(1, Math.floor(opts.chunkSizeChars ?? 1_000_000));
    const aligned = rawChunkSizeChars - (rawChunkSizeChars % 4);
    this.chunkSizeChars = aligned >= 4 ? aligned : 4;
  }

  async load() {
    const storage = getLocalStorageOrNull();
    if (!storage) return null;

    const metaRaw = storage.getItem(this.metaKey);
    if (!metaRaw) {
      // Backwards compatibility: earlier versions stored the full base64 blob in a single
      // `${namespace}:${workbookId}` key (LocalStorageBinaryStorage). If we find a legacy
      // value, decode it and opportunistically migrate it into the chunked format.
      const legacy = storage.getItem(this.key);
      if (!legacy) {
        // If a previous write partially completed (chunks written but meta missing),
        // clear the orphaned chunks so localStorage doesn't slowly fill up.
        // We probe for chunk 0 first to avoid scanning `storage.length` in the common case.
        const orphanChunk0 = storage.getItem(`${this.key}:0`);
        if (typeof orphanChunk0 === "string") {
          try {
            await this.remove();
          } catch {
            // ignore
          }
        }
        return null;
      }
      try {
        const decoded = fromBase64(legacy);
        // Best-effort migration: write in chunked format and remove the legacy key to
        // free space and avoid stale duplicates.
        try {
          await this.save(decoded);
        } catch {
          // ignore
        }
        return decoded;
      } catch {
        // Corrupted legacy base64 payload; clear it so future loads don't keep
        // failing (and to free storage space).
        try {
          storage.removeItem?.(this.key);
        } catch {
          // ignore
        }
        return null;
      }
    }

    /** @type {{ chunks?: number } | null} */
    let meta = null;
    try {
      // eslint-disable-next-line @typescript-eslint/no-unsafe-assignment
      meta = JSON.parse(metaRaw);
    } catch {
      // Corrupted meta; clear all persisted chunks.
      try {
        await this.remove();
      } catch {
        // ignore
      }
      return null;
    }

    const chunks = meta?.chunks;
    if (!Number.isInteger(chunks) || chunks < 0) {
      try {
        await this.remove();
      } catch {
        // ignore
      }
      return null;
    }
    // Guard against corrupted metadata that could otherwise hang the main thread
    // by attempting to load an absurd number of chunks.
    const MAX_CHUNKS = 10_000;
    if (chunks > MAX_CHUNKS) {
      try {
        await this.remove();
      } catch {
        // ignore
      }
      return null;
    }

    // Decode chunk-by-chunk when possible (avoids joining a potentially huge
    // base64 string in memory). Fall back to concatenating and decoding when
    // existing chunk boundaries are not base64-aligned (older callers may have
    // provided a chunkSizeChars that wasn't a multiple of 4).
    const chunkIsDecodable = (part, index) => {
      if (part.length % 4 !== 0) return false;
      // Intermediate chunks should never contain padding. If they do, treat as
      // legacy/untrusted and fall back to full join+decode.
      if (index < chunks - 1 && part.includes("=")) return false;
      return true;
    };

    const decodedLength = (encoded) => {
      const padding = encoded.endsWith("==") ? 2 : encoded.endsWith("=") ? 1 : 0;
      return (encoded.length * 3) / 4 - padding;
    };

    let canDecodeIndividually = true;
    let totalBytes = 0;
    for (let i = 0; i < chunks; i += 1) {
      const part = storage.getItem(`${this.key}:${i}`);
      if (typeof part !== "string") {
        // Missing/truncated chunk; clear persisted state so future opens can recover.
        try {
          await this.remove();
        } catch {
          // ignore
        }
        return null;
      }
      if (!chunkIsDecodable(part, i)) {
        canDecodeIndividually = false;
        break;
      }
      totalBytes += decodedLength(part);
    }

    if (!canDecodeIndividually) {
      /** @type {string[]} */
      const parts = [];
      for (let i = 0; i < chunks; i += 1) {
        const part = storage.getItem(`${this.key}:${i}`);
        if (typeof part !== "string") {
          try {
            await this.remove();
          } catch {
            // ignore
          }
          return null;
        }
        parts.push(part);
      }

      try {
        return fromBase64(parts.join(""));
      } catch {
        // Corrupted base64 payload; clear persisted chunks.
        try {
          await this.remove();
        } catch {
          // ignore
        }
        return null;
      }
    }

    const out = new Uint8Array(totalBytes);
    let offset = 0;
    for (let i = 0; i < chunks; i += 1) {
      const part = storage.getItem(`${this.key}:${i}`);
      if (typeof part !== "string") {
        try {
          await this.remove();
        } catch {
          // ignore
        }
        return null;
      }
      try {
        const bytes = fromBase64(part);
        out.set(bytes, offset);
        offset += bytes.byteLength;
      } catch {
        try {
          await this.remove();
        } catch {
          // ignore
        }
        return null;
      }
    }
    return out;
  }

  /**
   * @param {Uint8Array} data
   */
  async save(data) {
    const storage = getLocalStorageOrNull();
    if (!storage) return;

    // Rollback protection: `setItem` can throw (quota / permissions). Since we write chunks to
    // stable keys (`${key}:0`, `${key}:1`, ...), a mid-write failure could otherwise corrupt
    // previously-saved data (meta would still point at the old chunk count, but chunk 0 might
    // already be overwritten). Keep a small in-memory backup of any chunk keys we overwrite so
    // we can restore them if something throws.
    /** @type {Array<{ key: string, prev: string | null }>} */
    const backups = [];
    const prevMeta = storage.getItem(this.metaKey);

    /** @type {number} */
    let chunks = 0;

    try {
      // Prefer Node's Buffer when available.
      if (typeof Buffer !== "undefined") {
        const encoded = toBase64(data);
        chunks = encoded.length === 0 ? 0 : Math.ceil(encoded.length / this.chunkSizeChars);
        for (let i = 0; i < chunks; i += 1) {
          const chunkKey = `${this.key}:${i}`;
          backups.push({ key: chunkKey, prev: storage.getItem(chunkKey) });
          const start = i * this.chunkSizeChars;
          storage.setItem(chunkKey, encoded.slice(start, start + this.chunkSizeChars));
        }
      } else {
        // Browser-safe streaming base64 encode + chunking (avoids building a full
        // multi-MB base64 string in memory).
        const MAX_FROM_CHAR_CODE_ARGS = 0x8000;
        const BYTE_CHUNK_SIZE = MAX_FROM_CHAR_CODE_ARGS - (MAX_FROM_CHAR_CODE_ARGS % 3); // keep base64 alignment
        let pending = "";
        for (let i = 0; i < data.length; i += BYTE_CHUNK_SIZE) {
          const chunk = data.subarray(i, i + BYTE_CHUNK_SIZE);
          // eslint-disable-next-line prefer-spread
          const binary = String.fromCharCode.apply(null, chunk);
          // eslint-disable-next-line no-undef
          pending += btoa(binary);

          while (pending.length >= this.chunkSizeChars) {
            const slice = pending.slice(0, this.chunkSizeChars);
            pending = pending.slice(this.chunkSizeChars);
            const chunkKey = `${this.key}:${chunks}`;
            backups.push({ key: chunkKey, prev: storage.getItem(chunkKey) });
            storage.setItem(chunkKey, slice);
            chunks += 1;
          }
        }

        if (pending.length > 0) {
          const chunkKey = `${this.key}:${chunks}`;
          backups.push({ key: chunkKey, prev: storage.getItem(chunkKey) });
          storage.setItem(chunkKey, pending);
          chunks += 1;
        }
      }

      storage.setItem(this.metaKey, JSON.stringify({ chunks }));
    } catch (err) {
      // Best-effort rollback to preserve the previously persisted value.
      try {
        if (prevMeta == null) storage.removeItem?.(this.metaKey);
        else storage.setItem(this.metaKey, prevMeta);
      } catch {
        // ignore
      }

      for (const { key, prev } of backups) {
        try {
          if (prev == null) storage.removeItem?.(key);
          else storage.setItem(key, prev);
        } catch {
          // ignore
        }
      }

      throw err;
    }

    // Remove leftover chunks from a prior (larger) save. Best-effort: successful chunk + meta
    // writes should not be treated as failed persistence because cleanup encountered a storage error.
    try {
      // We avoid relying solely on stored metadata so we can clean up even if `:meta` was missing/corrupted.
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
      for (const key of keysToRemove) storage.removeItem?.(key);

      // Clean up the legacy single-key storage entry if it exists.
      storage.removeItem?.(this.key);
    } catch {
      // ignore
    }
  }

  async remove() {
    const storage = getLocalStorageOrNull();
    if (!storage) return;

    const prefix = `${this.key}:`;
    /** @type {string[]} */
    const keysToRemove = [];
    for (let i = 0; i < storage.length; i += 1) {
      const key = storage.key(i);
      if (!key) continue;
      if (key === this.metaKey) {
        keysToRemove.push(key);
        continue;
      }
      if (!key.startsWith(prefix)) continue;
      const suffix = key.slice(prefix.length);
      if (!/^\d+$/.test(suffix)) continue;
      keysToRemove.push(key);
    }

    for (const key of keysToRemove) storage.removeItem(key);

    // Also remove the legacy single-key entry (if it exists) so `load()` doesn't
    // fall back to stale data after a caller explicitly clears the store.
    storage.removeItem(this.key);
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
      try {
        return await normalizeBinary(value);
      } catch {
        // Corrupted IndexedDB record (unexpected type / shape). Clear it so future
        // loads don't repeatedly fail and callers can recover by re-indexing.
        try {
          await idbWrite(db, this.storeName, (store) => store.delete(this.key));
        } catch {
          // ignore
        }
        return null;
      }
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
      // Avoid copying when the view already spans the full buffer (common for sql.js exports).
      const buffer =
        data.byteOffset === 0 && data.byteLength === data.buffer.byteLength
          ? data.buffer
          : data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
      await idbWrite(db, this.storeName, (store) => store.put(buffer, this.key));
    } catch {
      // Best-effort persistence. If IndexedDB writes fail (quota / permissions),
      // we intentionally do not throw to keep callers usable in restricted contexts.
    }
  }

  async remove() {
    const db = await this._openOrNull();
    if (!db) return;
    try {
      await idbWrite(db, this.storeName, (store) => store.delete(this.key));
    } catch {
      // Best-effort removal. If IndexedDB writes fail (quota / permissions),
      // callers can proceed without a hard failure.
    }
  }

  async _openOrNull() {
    const idb = getIndexedDBOrNull();
    if (!idb) return null;
    if (this._dbPromise) {
      try {
        return await this._dbPromise;
      } catch {
        // If opening failed (quota / permission / blocked), allow future attempts.
        this._dbPromise = null;
        return null;
      }
    }

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
      this._dbPromise = null;
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
  // Avoid extra copies: IndexedDB already returns structured-cloned values, so returning
  // a view into the returned buffer is safe and significantly reduces memory churn for
  // large sqlite exports.
  if (value instanceof Uint8Array) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (value && typeof value === "object" && value.buffer instanceof ArrayBuffer) {
    // Structured clone can yield `{ buffer, byteOffset, byteLength }` shapes for typed arrays.
    const offset = typeof value.byteOffset === "number" ? value.byteOffset : 0;
    const length =
      typeof value.byteLength === "number"
        ? value.byteLength
        : typeof value.length === "number"
          ? value.length
          : value.buffer.byteLength;
    return new Uint8Array(value.buffer, offset, length);
  }
  if (typeof Blob !== "undefined" && value instanceof Blob) {
    const buffer = await value.arrayBuffer();
    return new Uint8Array(buffer);
  }
  throw new Error("Invalid binary blob in IndexedDB record");
}
