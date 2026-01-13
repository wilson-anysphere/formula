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
 * @typedef {"single" | "stream"} WriteMode
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
  // Bundlers can rename constructors and pnpm workspaces can load multiple `yjs`
  // module instances (ESM + CJS). Avoid relying on `constructor.name`; prefer a
  // structural check instead.
  if (typeof maybe.get !== "function") return false;
  if (typeof maybe.set !== "function") return false;
  if (typeof maybe.delete !== "function") return false;
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return true;
}

/**
 * @param {unknown} value
 * @returns {value is Y.Array<any>}
 */
function isYArray(value) {
  if (value instanceof Y.Array) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  return (
    typeof maybe.get === "function" &&
    typeof maybe.toArray === "function" &&
    typeof maybe.push === "function" &&
    typeof maybe.delete === "function" &&
    typeof maybe.observeDeep === "function" &&
    typeof maybe.unobserveDeep === "function"
  );
}

function isYAbstractType(value) {
  if (value instanceof Y.AbstractType) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return Boolean(maybe._map instanceof Map || maybe._start || maybe._item || maybe._length != null);
}

function replaceForeignRootType(params) {
  const { doc, name, existing, create } = params;
  const t = create();

  // Mirror Yjs' own Doc.get conversion logic for AbstractType placeholders, but
  // also support roots instantiated by a different Yjs module instance (e.g.
  // CJS `require("yjs")`).
  //
  // We intentionally only do this replacement when `doc` is from this module's
  // Yjs instance (i.e. `doc instanceof Y.Doc`). If the entire doc was created by
  // a foreign Yjs build, inserting local types into it can cause the same
  // cross-instance integration errors we're trying to avoid.
  (t)._map = existing?._map;
  (t)._start = existing?._start;
  (t)._length = existing?._length;

  const map = existing?._map;
  if (map instanceof Map) {
    map.forEach((item) => {
      for (let n = item; n !== null; n = n.left) {
        n.parent = t;
      }
    });
  }

  for (let n = existing?._start ?? null; n !== null; n = n.right) {
    n.parent = t;
  }

  doc.share.set(name, t);
  t._integrate(doc, null);
  return t;
}

/**
 * @param {import("yjs").Doc} doc
 * @param {string} name
 * @returns {any}
 */
function getMapRoot(doc, name) {
  const existing = doc.share.get(name);
  if (existing == null) return doc.getMap(name);

  if (isYMap(existing)) {
    // If the map root was created by a different Yjs module instance (ESM vs CJS),
    // `instanceof` checks fail and inserting local nested types can throw
    // ("Unexpected content type"). Normalize the root to this module instance.
    if (!(existing instanceof Y.Map) && doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
    }
    return existing;
  }

  // Placeholder root types should be coerced via Yjs' own constructors.
  //
  // Note: other parts of the system (e.g. collaborative undo) patch foreign
  // prototype chains so foreign types can pass `instanceof Y.AbstractType`
  // checks. Use constructor identity to detect placeholders created by *this*
  // Yjs module instance.
  if (existing instanceof Y.AbstractType && existing.constructor === Y.AbstractType) return doc.getMap(name);
  if (isYAbstractType(existing) && doc instanceof Y.Doc) {
    return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
  }
  if (isYAbstractType(existing)) return doc.getMap(name);

  // Root is missing or unsupported; instantiate via Yjs' constructor (will throw
  // if the root exists but has a non-map schema).
  return doc.getMap(name);
}

/**
 * @param {any} map
 * @returns {any}
 */
function mapConstructor(map) {
  const ctor = map?.constructor;
  return typeof ctor === "function" ? ctor : Y.Map;
}

/**
 * @param {any} array
 * @returns {any}
 */
function arrayConstructor(array) {
  const ctor = array?.constructor;
  return typeof ctor === "function" ? ctor : Y.Array;
}

/**
 * @param {any} ctor
 * @returns {any | null}
 */
function abstractTypeSuperclass(ctor) {
  if (typeof ctor !== "function") return null;
  const parent = Object.getPrototypeOf(ctor);
  return typeof parent === "function" ? parent : null;
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
    // Keep the import specifier non-literal so bundlers targeting the browser
    // (e.g. esbuild) don't try to resolve `node:zlib` at build time.
    const zlibName = "zlib";
    const zlib = await import(/* @vite-ignore */ `node:${zlibName}`);
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
    // Keep the import specifier non-literal so bundlers targeting the browser
    // (e.g. esbuild) don't try to resolve `node:zlib` at build time.
    const zlibName = "zlib";
    const zlib = await import(/* @vite-ignore */ `node:${zlibName}`);
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
   *   writeMode?: WriteMode;
   *   maxChunksPerTransaction?: number;
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
    /** @type {WriteMode} */
    this.writeMode = opts.writeMode ?? "single";
    this.maxChunksPerTransaction = opts.maxChunksPerTransaction ?? null;

    if (this.chunkSize <= 0) throw new Error("YjsVersionStore: chunkSize must be > 0");
    if (this.compression !== "none" && this.compression !== "gzip") {
      throw new Error(`YjsVersionStore: invalid compression: ${this.compression}`);
    }
    if (this.snapshotEncoding !== "chunks" && this.snapshotEncoding !== "base64") {
      throw new Error(`YjsVersionStore: invalid snapshotEncoding: ${this.snapshotEncoding}`);
    }
    if (this.writeMode !== "single" && this.writeMode !== "stream") {
      throw new Error(`YjsVersionStore: invalid writeMode: ${this.writeMode}`);
    }
    if (this.maxChunksPerTransaction != null) {
      if (
        typeof this.maxChunksPerTransaction !== "number" ||
        !Number.isFinite(this.maxChunksPerTransaction) ||
        this.maxChunksPerTransaction <= 0
      ) {
        throw new Error("YjsVersionStore: maxChunksPerTransaction must be a positive number");
      }
    }

    /** @type {Y.Map<any>} */
    this.versions = getMapRoot(this.doc, "versions");
    /** @type {Y.Map<any>} */
    this.meta = getMapRoot(this.doc, "versionsMeta");
  }

  /**
   * @returns {Y.Array<any> | null}
   */
  _ensureOrderArray() {
    const existing = this.meta.get("order");
    if (isYArray(existing)) return existing;

    const metaCtor = mapConstructor(this.meta);
    const arrayCtor = this._inferArrayCtorForMapCtor(metaCtor);
    if (!arrayCtor) return null;
    const created = new arrayCtor();
    this.meta.set("order", created);
    return created;
  }

  /**
   * Try to infer a `Y.Array` constructor compatible with the given map constructor.
   *
   * This is needed when multiple `yjs` module instances are loaded (ESM vs CJS).
   * Yjs does strict `instanceof AbstractType` checks when integrating nested types.
   *
   * @param {any} mapCtor
   * @returns {any | null}
   */
  _inferArrayCtorForMapCtor(mapCtor) {
    const moduleId = abstractTypeSuperclass(mapCtor);
    if (!moduleId) return null;

    const localModuleId = abstractTypeSuperclass(Y.Map);
    if (moduleId === localModuleId) return Y.Array;

    // Fast path: probe the Yjs constructors used to create `doc` so we can
    // allocate nested arrays even when different Yjs module instances are loaded
    // (e.g. pnpm workspaces where y-websocket uses CJS `require("yjs")` and the
    // app uses ESM `import "yjs"`).
    const DocCtor = /** @type {any} */ (this.doc)?.constructor;
    if (typeof DocCtor === "function") {
      try {
        const probe = new DocCtor();
        const probeMapCtor = probe.getMap("__versioning_store_ctor_probe_map").constructor;
        if (abstractTypeSuperclass(probeMapCtor) === moduleId) {
          const probeArrayCtor = probe.getArray("__versioning_store_ctor_probe_array").constructor;
          if (abstractTypeSuperclass(probeArrayCtor) === moduleId) return probeArrayCtor;
        }
      } catch {
        // ignore
      }
    }

    const existingOrder = this.meta.get("order");
    if (isYArray(existingOrder)) {
      const orderCtor = arrayConstructor(existingOrder);
      if (abstractTypeSuperclass(orderCtor) === moduleId) return orderCtor;
    }

    /** @type {any | null} */
    let ctor = null;
    this.versions.forEach((value) => {
      if (ctor) return;
      if (!isYMap(value)) return;
      if (abstractTypeSuperclass(value.constructor) !== moduleId) return;
      const chunks = value.get("snapshotChunks");
      if (isYArray(chunks)) {
        const chunksCtor = arrayConstructor(chunks);
        if (abstractTypeSuperclass(chunksCtor) === moduleId) ctor = chunksCtor;
      }
    });
    if (ctor) return ctor;

    return null;
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
    const desiredSnapshotEncoding = this.writeMode === "stream" ? "chunks" : this.snapshotEncoding;
    const createdAtMs = Date.now();

    if (this.writeMode !== "stream") {
      this.doc.transact(() => {
        const MapCtor = mapConstructor(this.versions);
        const arrayCtor =
          desiredSnapshotEncoding === "chunks" ? this._inferArrayCtorForMapCtor(MapCtor) : null;
        const actualSnapshotEncoding =
          desiredSnapshotEncoding === "chunks" && arrayCtor ? "chunks" : "base64";
        /** @type {Y.Map<any>} */
        const record = new MapCtor();
        record.set("schemaVersion", 1);
        record.set("id", version.id);
        record.set("kind", version.kind);
        record.set("timestampMs", version.timestampMs);
        record.set("createdAtMs", createdAtMs);
        record.set("userId", version.userId ?? null);
        record.set("userName", version.userName ?? null);
        record.set("description", version.description ?? null);
        record.set("checkpointName", version.checkpointName ?? null);
        record.set("checkpointLocked", version.checkpointLocked ?? null);
        record.set("checkpointAnnotations", version.checkpointAnnotations ?? null);

        record.set("compression", compression);
        record.set("snapshotEncoding", actualSnapshotEncoding);

        if (actualSnapshotEncoding === "base64") {
          record.set("snapshotBase64", bytesToBase64(snapshotBytes));
        } else {
          const chunks = splitIntoChunks(snapshotBytes, this.chunkSize);
          const arr = new arrayCtor();
          arr.push(chunks);
          record.set("snapshotChunks", arr);
        }

        this.versions.set(version.id, record);

        const order = this._ensureOrderArray();
        if (order) order.push([version.id]);
      }, "versioning-store");
      return;
    }

    // Streaming mode: write the record metadata first, then append snapshot chunks
    // across multiple transactions so each Yjs update stays small enough for the
    // sync-server's `SYNC_SERVER_MAX_MESSAGE_BYTES` limit.
    const MapCtor = mapConstructor(this.versions);
    const arrayCtor = this._inferArrayCtorForMapCtor(MapCtor);
    if (!arrayCtor) {
      // Fallback: if we can't create a compatible Y.Array constructor for nested
      // chunk arrays (multiple yjs instances loaded), store as base64 in a single
      // update. This matches legacy behavior (and may still exceed message limits
      // for very large snapshots).
      this.doc.transact(() => {
        /** @type {Y.Map<any>} */
        const record = new MapCtor();
        record.set("schemaVersion", 1);
        record.set("id", version.id);
        record.set("kind", version.kind);
        record.set("timestampMs", version.timestampMs);
        record.set("createdAtMs", createdAtMs);
        record.set("userId", version.userId ?? null);
        record.set("userName", version.userName ?? null);
        record.set("description", version.description ?? null);
        record.set("checkpointName", version.checkpointName ?? null);
        record.set("checkpointLocked", version.checkpointLocked ?? null);
        record.set("checkpointAnnotations", version.checkpointAnnotations ?? null);

        record.set("compression", compression);
        record.set("snapshotEncoding", "base64");
        record.set("snapshotBase64", bytesToBase64(snapshotBytes));

        this.versions.set(version.id, record);
        const order = this._ensureOrderArray();
        if (order) order.push([version.id]);
      }, "versioning-store");
      return;
    }

    const chunks = splitIntoChunks(snapshotBytes, this.chunkSize);
    const maxChunksPerTransaction =
      this.maxChunksPerTransaction ?? Math.max(1, Math.floor((256 * 1024) / this.chunkSize));

    /** @type {any} */
    let record = null;
    this.doc.transact(() => {
      /** @type {Y.Map<any>} */
      const r = new MapCtor();
      r.set("schemaVersion", 1);
      r.set("id", version.id);
      r.set("kind", version.kind);
      r.set("timestampMs", version.timestampMs);
      r.set("createdAtMs", createdAtMs);
      r.set("userId", version.userId ?? null);
      r.set("userName", version.userName ?? null);
      r.set("description", version.description ?? null);
      r.set("checkpointName", version.checkpointName ?? null);
      r.set("checkpointLocked", version.checkpointLocked ?? null);
      r.set("checkpointAnnotations", version.checkpointAnnotations ?? null);

      r.set("compression", compression);
      r.set("snapshotEncoding", "chunks");
      r.set("snapshotChunkCountExpected", chunks.length);

      const arr = new arrayCtor();
      r.set("snapshotChunks", arr);
      r.set("snapshotComplete", false);

      this.versions.set(version.id, r);

      const order = this._ensureOrderArray();
      if (order) order.push([version.id]);
      record = r;
    }, "versioning-store");

    // Append chunks in batches, each in its own Yjs transaction/update.
    for (let i = 0; i < chunks.length; i += maxChunksPerTransaction) {
      const batch = chunks.slice(i, Math.min(i + maxChunksPerTransaction, chunks.length));
      this.doc.transact(() => {
        const raw = record ?? this.versions.get(version.id);
        if (!isYMap(raw)) {
          throw new Error(`YjsVersionStore: missing streamed version record for ${version.id}`);
        }
        const chunksArr = raw.get("snapshotChunks");
        if (!isYArray(chunksArr)) {
          throw new Error(`YjsVersionStore: missing snapshotChunks for ${version.id}`);
        }
        chunksArr.push(batch);
      }, "versioning-store");
    }

    this.doc.transact(() => {
      const raw = record ?? this.versions.get(version.id);
      if (!isYMap(raw)) return;
      raw.set("snapshotComplete", true);
    }, "versioning-store");
  }

  /**
   * @param {string} versionId
   * @returns {Promise<VersionRecord | null>}
   */
  async getVersion(versionId) {
    const raw = this.versions.get(versionId);
    if (!isYMap(raw)) return null;

    // Streaming writes may produce partially-written versions. Avoid surfacing
    // these until the snapshot blob has been fully appended. Check this early so
    // malformed/incomplete records don't throw during field validation.
    const snapshotCompleteRaw = raw.get("snapshotComplete");
    if (snapshotCompleteRaw === false) return null;

    // Additional incompleteness checks: a record can be unreadable even if
    // `snapshotComplete` is unset/incorrect (e.g. legacy writers or corrupted
    // docs). If we can prove the snapshot is not fully present, treat it as
    // incomplete and return null rather than throwing on missing metadata.
    const snapshotEncodingRaw = raw.get("snapshotEncoding");
    const chunksArr = raw.get("snapshotChunks");
    if ((snapshotEncodingRaw === "chunks" || chunksArr !== undefined) && !isYArray(chunksArr)) {
      // Chunks mode must have a chunk array.
      return null;
    }
    if (isYArray(chunksArr)) {
      const expectedChunks = raw.get("snapshotChunkCountExpected");
      if (typeof expectedChunks === "number" && chunksArr.length < expectedChunks) return null;
    }
    if (snapshotEncodingRaw === "base64") {
      const base64 = raw.get("snapshotBase64");
      if (typeof base64 !== "string") return null;
    }

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
      if (typeof base64 !== "string") return null;
      stored = base64ToBytes(base64);
    } else if (snapshotEncoding === "chunks") {
      const chunksArr = raw.get("snapshotChunks");
      if (!isYArray(chunksArr)) return null;
      const expectedChunks = raw.get("snapshotChunkCountExpected");
      if (typeof expectedChunks === "number" && chunksArr.length < expectedChunks) return null;
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
   * Best-effort cleanup for partially-written streamed version records.
   *
   * Streaming `saveVersion()` writes create a record with `snapshotComplete=false`
   * and then append snapshot chunks across multiple Yjs transactions. If the
   * writer crashes, these incomplete records remain in the Y.Doc but are hidden
   * from callers (`getVersion()` returns null, `listVersions()` drops them).
   *
   * Some older/corrupted docs may also contain version records that are unreadable
   * even when `snapshotComplete` is unset/incorrect (e.g. chunk payloads that are
   * missing expected chunks).
   *
   * Because callers don't see them, retention policies may never delete them.
   * This method deletes *stale* incomplete records so they don't accumulate
   * indefinitely.
   *
   * @param {{ olderThanMs?: number }} [opts]
   * @returns {Promise<{ prunedIds: string[] }>}
   */
  async pruneIncompleteVersions(opts) {
    const olderThanMs = opts?.olderThanMs ?? 10 * 60 * 1000;
    const nowMs = Date.now();

    /** @type {Set<string>} */
    const staleIds = new Set();
    /** @type {Set<string>} */
    const finalizeIds = new Set();

    this.versions.forEach((value, key) => {
      if (typeof key !== "string") return;
      if (!isYMap(value)) return;

      const snapshotComplete = value.get("snapshotComplete");
      const snapshotEncoding = value.get("snapshotEncoding");
      const chunksArr = value.get("snapshotChunks");
      const expectedChunks = value.get("snapshotChunkCountExpected");
      const base64 = value.get("snapshotBase64");

      const hasChunks = chunksArr !== undefined;
      const chunksExpected = typeof expectedChunks === "number" && expectedChunks >= 0;
      const isChunkEncoded = snapshotEncoding === "chunks" || (snapshotEncoding == null && hasChunks);
      const hasChunkArray = isYArray(chunksArr);

      const isMissingChunks = isChunkEncoded && chunksExpected && hasChunkArray && chunksArr.length < expectedChunks;
      const isMissingChunkArray = isChunkEncoded && hasChunks && !hasChunkArray;
      const isMissingBase64 = snapshotEncoding === "base64" && typeof base64 !== "string";

      const isIncomplete = snapshotComplete === false || isMissingChunks || isMissingChunkArray || isMissingBase64;
      if (!isIncomplete) return;

      // If all chunks are present but the writer crashed before flipping the
      // `snapshotComplete` flag, we can safely finalize the record instead of
      // deleting it (keeps the snapshot recoverable and allows normal retention
      // to apply).
      if (snapshotComplete === false && chunksExpected && hasChunkArray && chunksArr.length >= expectedChunks) {
        {
          // Only finalize records that are likely to be readable via `getVersion`.
          // If required metadata is missing/corrupted, keep it incomplete so it
          // can be deleted via the staleness policy below.
          const schemaVersion = value.get("schemaVersion") ?? 1;
          const kind = value.get("kind");
          const timestampMs = value.get("timestampMs");
          const kindValid = kind === "snapshot" || kind === "checkpoint" || kind === "restore";
          const compression = value.get("compression");
          const compressionValid =
            compression == null || compression === "none" || compression === "gzip";
          const snapshotEncodingValid = snapshotEncoding == null || snapshotEncoding === "chunks";
          if (
            schemaVersion === 1 &&
            kindValid &&
            typeof timestampMs === "number" &&
            Number.isFinite(timestampMs) &&
            compressionValid &&
            snapshotEncodingValid
          ) {
            finalizeIds.add(key);
            return;
          }
        }
      }

      // Prefer `createdAtMs` if present (newer schema), falling back to the
      // version timestamp for backwards compatibility.
      const createdAtMs = value.get("createdAtMs");
      let ts =
        typeof createdAtMs === "number"
          ? createdAtMs
          : typeof value.get("timestampMs") === "number"
            ? value.get("timestampMs")
            : 0;

      // Defensive: clamp timestamps to avoid "future" records never becoming stale
      // (can happen with clock skew or older clients without createdAtMs).
      if (!Number.isFinite(ts) || ts < 0) ts = 0;
      if (ts > nowMs) ts = nowMs;

      if (nowMs - ts >= olderThanMs) staleIds.add(key);
    });

    if (staleIds.size === 0 && finalizeIds.size === 0) return { prunedIds: [] };

    const prunedIds = Array.from(staleIds);
    this.doc.transact(() => {
      for (const id of finalizeIds) {
        const raw = this.versions.get(id);
        if (!isYMap(raw)) continue;
        if (raw.get("snapshotComplete") !== false) continue;
        const expectedChunks = raw.get("snapshotChunkCountExpected");
        const chunksArr = raw.get("snapshotChunks");
        if (typeof expectedChunks !== "number" || expectedChunks < 0 || !isYArray(chunksArr)) continue;
        if (chunksArr.length < expectedChunks) continue;
        const schemaVersion = raw.get("schemaVersion") ?? 1;
        const kind = raw.get("kind");
        const timestampMs = raw.get("timestampMs");
        const kindValid = kind === "snapshot" || kind === "checkpoint" || kind === "restore";
        const compression = raw.get("compression");
        const compressionValid = compression == null || compression === "none" || compression === "gzip";
        const snapshotEncoding = raw.get("snapshotEncoding");
        const snapshotEncodingValid = snapshotEncoding == null || snapshotEncoding === "chunks";
        if (
          schemaVersion !== 1 ||
          !kindValid ||
          typeof timestampMs !== "number" ||
          !Number.isFinite(timestampMs) ||
          !compressionValid ||
          !snapshotEncodingValid
        ) {
          continue;
        }
        raw.set("snapshotComplete", true);
      }

      for (const id of staleIds) this.versions.delete(id);

      const order = this.meta.get("order");
      if (!isYArray(order)) return;
      for (let i = order.length - 1; i >= 0; i -= 1) {
        const id = order.get(i);
        if (typeof id === "string" && staleIds.has(id)) {
          order.delete(i, 1);
        }
      }
    }, "versioning-store");

    return { prunedIds };
  }

  /**
   * @returns {Promise<VersionRecord[]>}
   */
  async listVersions() {
    // Opportunistically clean up stale partial streamed versions. This is
    // best-effort; `listVersions()` should still work even if the cleanup hits
    // unexpected schema/cross-instance issues.
    try {
      await this.pruneIncompleteVersions();
    } catch {
      // ignore
    }

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
