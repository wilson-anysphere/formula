import * as Y from "yjs";

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
 *
 * @typedef {"none" | "gzip"} SnapshotCompression
 * @typedef {"chunks" | "base64"} SnapshotEncoding
 */

function bytesToBase64(bytes) {
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return Buffer.from(bytes).toString("base64");
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) binary += String.fromCharCode(bytes[i]);
  // eslint-disable-next-line no-undef
  return btoa(binary);
}

function base64ToBytes(base64) {
  // eslint-disable-next-line no-undef
  if (typeof Buffer !== "undefined") return new Uint8Array(Buffer.from(base64, "base64"));
  // eslint-disable-next-line no-undef
  const binary = atob(base64);
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
  return out;
}

function isNodeRuntime() {
  const proc = /** @type {any} */ (globalThis.process);
  return Boolean(proc?.versions?.node);
}

/**
 * @param {unknown} value
 * @returns {value is Y.Map<any>}
 */
function isYMap(value) {
  if (value instanceof Y.Map) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YMap") return false;
  return typeof maybe.get === "function" && typeof maybe.set === "function" && typeof maybe.delete === "function";
}

/**
 * @param {unknown} value
 * @returns {value is Y.Array<any>}
 */
function isYArray(value) {
  if (value instanceof Y.Array) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (maybe.constructor?.name !== "YArray") return false;
  return (
    typeof maybe.get === "function" &&
    typeof maybe.toArray === "function" &&
    typeof maybe.push === "function" &&
    typeof maybe.delete === "function"
  );
}

/**
 * @param {import("yjs").Doc} doc
 * @param {string} name
 * @returns {any}
 */
function getMapRoot(doc, name) {
  const existing = doc.share.get(name);
  if (isYMap(existing)) return existing;
  // Root is missing or still a placeholder; instantiate via Yjs' constructor.
  return doc.getMap(name);
}

/**
 * @param {unknown} snapshot
 * @returns {Uint8Array}
 */
function normalizeSnapshotBytes(snapshot) {
  if (snapshot instanceof Uint8Array) return snapshot;
  if (snapshot instanceof ArrayBuffer) return new Uint8Array(snapshot);
  // Older structured clones can return `{ buffer, byteOffset, byteLength }` shapes.
  if (snapshot && typeof snapshot === "object" && /** @type {any} */ (snapshot).buffer instanceof ArrayBuffer) {
    const s = /** @type {any} */ (snapshot);
    return new Uint8Array(s.buffer, s.byteOffset ?? 0, s.byteLength);
  }
  throw new Error("YjsVersionStore: invalid snapshot bytes");
}

/**
 * @param {Uint8Array[]} chunks
 */
function concatChunks(chunks) {
  let total = 0;
  for (const chunk of chunks) total += chunk.length;
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.length;
  }
  return out;
}

/**
 * @param {Uint8Array} bytes
 * @param {number} chunkSize
 */
function splitIntoChunks(bytes, chunkSize) {
  if (bytes.length === 0) return [new Uint8Array()];
  /** @type {Uint8Array[]} */
  const chunks = [];
  for (let i = 0; i < bytes.length; i += chunkSize) {
    chunks.push(bytes.slice(i, Math.min(i + chunkSize, bytes.length)));
  }
  return chunks;
}

/**
 * @param {Uint8Array} bytes
 * @returns {Promise<Uint8Array>}
 */
async function gzipBytes(bytes) {
  // Prefer Node's zlib for stability/perf (Node also exposes CompressionStream).
  if (isNodeRuntime()) {
    const zlib = await import(/* @vite-ignore */ "node:zlib");
    return new Uint8Array(zlib.gzipSync(bytes));
  }

  // Browser (and other web runtimes).
  // eslint-disable-next-line no-undef
  if (typeof CompressionStream !== "undefined") {
    // eslint-disable-next-line no-undef
    const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("gzip"));
    // eslint-disable-next-line no-undef
    const compressed = await new Response(stream).arrayBuffer();
    return new Uint8Array(compressed);
  }

  throw new Error("YjsVersionStore: gzip compression is not supported in this environment");
}

/**
 * @param {Uint8Array} bytes
 * @returns {Promise<Uint8Array>}
 */
async function gunzipBytes(bytes) {
  if (isNodeRuntime()) {
    const zlib = await import(/* @vite-ignore */ "node:zlib");
    return new Uint8Array(zlib.gunzipSync(bytes));
  }

  // eslint-disable-next-line no-undef
  if (typeof DecompressionStream !== "undefined") {
    // eslint-disable-next-line no-undef
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
    // eslint-disable-next-line no-undef
    const decompressed = await new Response(stream).arrayBuffer();
    return new Uint8Array(decompressed);
  }

  throw new Error("YjsVersionStore: gzip decompression is not supported in this environment");
}

/**
 * @param {Uint8Array} bytes
 * @param {SnapshotCompression} compression
 * @returns {Promise<Uint8Array>}
 */
async function compressSnapshot(bytes, compression) {
  if (compression === "none") return bytes;
  if (compression === "gzip") return gzipBytes(bytes);
  throw new Error(`YjsVersionStore: unsupported compression: ${compression}`);
}

/**
 * @param {Uint8Array} bytes
 * @param {SnapshotCompression} compression
 * @returns {Promise<Uint8Array>}
 */
async function decompressSnapshot(bytes, compression) {
  if (compression === "none") return bytes;
  if (compression === "gzip") return gunzipBytes(bytes);
  throw new Error(`YjsVersionStore: unsupported compression: ${compression}`);
}

export class YjsVersionStore {
  /**
   * @param {{
   *   doc: import("yjs").Doc;
   *   chunkSize?: number;
   *   compression?: SnapshotCompression;
   *   snapshotEncoding?: SnapshotEncoding;
   * }} opts
   */
  constructor(opts) {
    if (!opts?.doc) throw new Error("YjsVersionStore: doc is required");

    this.doc = opts.doc;
    this.chunkSize = opts.chunkSize ?? 64 * 1024;
    /** @type {SnapshotCompression} */
    this.compression = opts.compression ?? "none";
    /** @type {SnapshotEncoding} */
    this.snapshotEncoding = opts.snapshotEncoding ?? "chunks";

    if (this.chunkSize <= 0) throw new Error("YjsVersionStore: chunkSize must be > 0");
    if (this.compression !== "none" && this.compression !== "gzip") {
      throw new Error(`YjsVersionStore: invalid compression: ${this.compression}`);
    }
    if (this.snapshotEncoding !== "chunks" && this.snapshotEncoding !== "base64") {
      throw new Error(`YjsVersionStore: invalid snapshotEncoding: ${this.snapshotEncoding}`);
    }

    /** @type {Y.Map<any>} */
    this.versions = getMapRoot(this.doc, "versions");
    /** @type {Y.Map<any>} */
    this.meta = getMapRoot(this.doc, "versionsMeta");
  }

  /**
   * @returns {Y.Array<any>}
   */
  _ensureOrderArray() {
    const existing = this.meta.get("order");
    if (isYArray(existing)) return existing;
    const created = new Y.Array();
    this.meta.set("order", created);
    return created;
  }

  /**
   * @returns {Map<string, number>}
   */
  _orderIndex() {
    const order = this.meta.get("order");
    if (!isYArray(order)) return new Map();
    const ids = order.toArray();
    /** @type {Map<string, number>} */
    const out = new Map();
    for (let i = 0; i < ids.length; i += 1) {
      const id = ids[i];
      if (typeof id === "string") out.set(id, i);
    }
    return out;
  }

  /**
   * @param {VersionRecord} version
   */
  async saveVersion(version) {
    const snapshot = normalizeSnapshotBytes(version.snapshot);
    const compression = this.compression;
    const snapshotBytes = await compressSnapshot(snapshot, compression);
    const snapshotEncoding = this.snapshotEncoding;

    this.doc.transact(() => {
      /** @type {Y.Map<any>} */
      const record = new Y.Map();
      record.set("schemaVersion", 1);
      record.set("id", version.id);
      record.set("kind", version.kind);
      record.set("timestampMs", version.timestampMs);
      record.set("userId", version.userId ?? null);
      record.set("userName", version.userName ?? null);
      record.set("description", version.description ?? null);
      record.set("checkpointName", version.checkpointName ?? null);
      record.set("checkpointLocked", version.checkpointLocked ?? null);
      record.set("checkpointAnnotations", version.checkpointAnnotations ?? null);

      record.set("compression", compression);
      record.set("snapshotEncoding", snapshotEncoding);

      if (snapshotEncoding === "base64") {
        record.set("snapshotBase64", bytesToBase64(snapshotBytes));
      } else {
        const chunks = splitIntoChunks(snapshotBytes, this.chunkSize);
        const arr = new Y.Array();
        arr.push(chunks);
        record.set("snapshotChunks", arr);
      }

      this.versions.set(version.id, record);

      const order = this._ensureOrderArray();
      order.push([version.id]);
    }, "versioning-store");
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    const raw = this.versions.get(versionId);
    if (!isYMap(raw)) return null;

    const schemaVersion = raw.get("schemaVersion") ?? 1;
    if (schemaVersion !== 1) {
      throw new Error(
        `YjsVersionStore: unsupported schemaVersion for ${versionId}: ${String(schemaVersion)}`,
      );
    }

    const kind = raw.get("kind");
    if (kind !== "snapshot" && kind !== "checkpoint" && kind !== "restore") {
      throw new Error(`YjsVersionStore: invalid kind for ${versionId}: ${String(kind)}`);
    }

    const timestampMs = raw.get("timestampMs");
    if (typeof timestampMs !== "number") {
      throw new Error(`YjsVersionStore: invalid timestampMs for ${versionId}`);
    }

    /** @type {SnapshotCompression} */
    const compression = raw.get("compression") ?? "none";
    /** @type {SnapshotEncoding} */
    const snapshotEncoding =
      raw.get("snapshotEncoding") ??
      (raw.get("snapshotChunks") ? "chunks" : raw.get("snapshotBase64") ? "base64" : null);

    /** @type {Uint8Array} */
    let stored;
    if (snapshotEncoding === "base64") {
      const base64 = raw.get("snapshotBase64");
      if (typeof base64 !== "string") {
        throw new Error(`YjsVersionStore: missing snapshotBase64 for ${versionId}`);
      }
      stored = base64ToBytes(base64);
    } else if (snapshotEncoding === "chunks") {
      const chunksArr = raw.get("snapshotChunks");
      if (!isYArray(chunksArr)) {
        throw new Error(`YjsVersionStore: missing snapshotChunks for ${versionId}`);
      }
      const chunksRaw = chunksArr.toArray();
      /** @type {Uint8Array[]} */
      const chunks = [];
      for (const chunk of chunksRaw) {
        chunks.push(normalizeSnapshotBytes(chunk));
      }
      stored = concatChunks(chunks);
    } else {
      throw new Error(`YjsVersionStore: missing snapshot encoding for ${versionId}`);
    }

    const snapshot = await decompressSnapshot(stored, compression);

    return {
      id: versionId,
      kind,
      timestampMs,
      userId: raw.get("userId") ?? null,
      userName: raw.get("userName") ?? null,
      description: raw.get("description") ?? null,
      checkpointName: raw.get("checkpointName") ?? null,
      checkpointLocked: raw.get("checkpointLocked") ?? null,
      checkpointAnnotations: raw.get("checkpointAnnotations") ?? null,
      snapshot,
    };
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    /** @type {string[]} */
    const ids = [];
    this.versions.forEach((_value, key) => {
      if (typeof key === "string") ids.push(key);
    });

    const records = await Promise.all(ids.map((id) => this.getVersion(id)));
    /** @type {VersionRecord[]} */
    const out = [];
    for (const r of records) {
      if (r) out.push(r);
    }

    const orderIndex = this._orderIndex();
    out.sort((a, b) => {
      const dt = b.timestampMs - a.timestampMs;
      if (dt !== 0) return dt;
      const ai = orderIndex.get(a.id) ?? 0;
      const bi = orderIndex.get(b.id) ?? 0;
      if (ai !== bi) return bi - ai;
      return a.id < b.id ? 1 : a.id > b.id ? -1 : 0;
    });

    return out;
  }

  /**
   * @param {string} versionId
   * @param {{ checkpointLocked?: boolean }} patch
   */
  async updateVersion(versionId, patch) {
    if (patch.checkpointLocked === undefined) return;
    this.doc.transact(() => {
      const raw = this.versions.get(versionId);
      if (!isYMap(raw)) {
        throw new Error(`Version not found: ${versionId}`);
      }
      raw.set("checkpointLocked", patch.checkpointLocked);
    }, "versioning-store");
  }

  /**
   * @param {string} versionId
   */
  async deleteVersion(versionId) {
    this.doc.transact(() => {
      this.versions.delete(versionId);

      const order = this.meta.get("order");
      if (!isYArray(order)) return;
      for (let i = order.length - 1; i >= 0; i -= 1) {
        if (order.get(i) === versionId) {
          order.delete(i, 1);
        }
      }
    }, "versioning-store");
  }
}
