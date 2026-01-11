/**
 * @typedef {"snapshot" | "checkpoint" | "restore"} VersionKind
 *
 * @typedef {{
 *   id: string;
 *   kind: VersionKind;
 *   timestampMs: number;
 *   userId: string | null;
 *   userName: string | null;
 *   description: string | null;
 *   checkpointName: string | null;
 *   checkpointLocked: boolean | null;
 *   checkpointAnnotations: string | null;
 *   snapshot: Uint8Array;
 * }} VersionRecord
 */

function requireIndexedDB() {
  const idb = globalThis.indexedDB;
  if (!idb) {
    throw new Error("IndexedDB is not available in this environment");
  }
  return idb;
}

/**
 * @template T
 * @param {IDBRequest<T>} request
 * @returns {Promise<T>}
 */
function requestToPromise(request) {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error ?? new Error("IndexedDB request failed"));
  });
}

/**
 * @param {IDBTransaction} tx
 * @returns {Promise<void>}
 */
function transactionDone(tx) {
  return new Promise((resolve, reject) => {
    tx.oncomplete = () => resolve();
    tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction aborted"));
    tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction error"));
  });
}

function normalizeSnapshot(snapshot) {
  if (snapshot instanceof Uint8Array) return snapshot;
  if (snapshot instanceof ArrayBuffer) return new Uint8Array(snapshot);
  // Older structured clones can return `{ buffer, byteOffset, byteLength }` shapes.
  if (snapshot && typeof snapshot === "object" && snapshot.buffer instanceof ArrayBuffer) {
    return new Uint8Array(snapshot.buffer, snapshot.byteOffset ?? 0, snapshot.byteLength);
  }
  throw new Error("Invalid snapshot blob in IndexedDB record");
}

/**
 * IndexedDB-backed version store (browser persistence).
 *
 * This is the web counterpart to `SQLiteVersionStore` used by desktop builds.
 */
export class IndexedDBVersionStore {
  /**
   * @param {{ dbName: string }} opts
   */
  constructor(opts) {
    if (!opts?.dbName) throw new Error("dbName is required");
    this.dbName = opts.dbName;
    /** @type {Promise<IDBDatabase> | null} */
    this._dbPromise = null;
  }

  async _open() {
    if (this._dbPromise) return this._dbPromise;
    const idb = requireIndexedDB();
    this._dbPromise = new Promise((resolve, reject) => {
      const request = idb.open(this.dbName, 1);
      request.onupgradeneeded = () => {
        const db = request.result;
        const store = db.createObjectStore("versions", { keyPath: "id" });
        store.createIndex("timestampMs", "timestampMs");
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error ?? new Error("Failed to open IndexedDB"));
    });
    return this._dbPromise;
  }

  /**
   * @param {VersionRecord} version
   */
  async saveVersion(version) {
    const db = await this._open();
    const tx = db.transaction("versions", "readwrite");
    const store = tx.objectStore("versions");
    store.put({
      ...version,
      snapshot: version.snapshot,
    });
    await transactionDone(tx);
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    const db = await this._open();
    const tx = db.transaction("versions", "readonly");
    const store = tx.objectStore("versions");
    const row = await requestToPromise(store.get(versionId));
    await transactionDone(tx);
    if (!row) return null;
    return {
      ...row,
      snapshot: normalizeSnapshot(row.snapshot),
    };
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    const db = await this._open();
    const tx = db.transaction("versions", "readonly");
    const store = tx.objectStore("versions");
    const index = store.index("timestampMs");
    /** @type {VersionRecord[]} */
    const out = [];
    const cursorRequest = index.openCursor(null, "prev");
    await new Promise((resolve, reject) => {
      cursorRequest.onsuccess = () => {
        const cursor = cursorRequest.result;
        if (!cursor) return resolve();
        const row = cursor.value;
        out.push({ ...row, snapshot: normalizeSnapshot(row.snapshot) });
        cursor.continue();
      };
      cursorRequest.onerror = () =>
        reject(cursorRequest.error ?? new Error("Failed to iterate IndexedDB cursor"));
    });
    await transactionDone(tx);
    return out;
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   */
  async updateVersion(versionId, patch) {
    if (patch.checkpointLocked === undefined) return;
    const db = await this._open();
    const tx = db.transaction("versions", "readwrite");
    const store = tx.objectStore("versions");
    const row = await requestToPromise(store.get(versionId));
    if (!row) {
      tx.abort();
      throw new Error(`Version not found: ${versionId}`);
    }
    row.checkpointLocked = patch.checkpointLocked;
    store.put(row);
    await transactionDone(tx);
  }

  /**
   * @param {string} versionId
   * @returns {Promise<void>}
   */
  async deleteVersion(versionId) {
    const db = await this._open();
    const tx = db.transaction("versions", "readwrite");
    const store = tx.objectStore("versions");
    store.delete(versionId);
    await transactionDone(tx);
  }

  close() {
    if (!this._dbPromise) return;
    void this._dbPromise.then((db) => db.close()).catch(() => {});
    this._dbPromise = null;
  }
}
