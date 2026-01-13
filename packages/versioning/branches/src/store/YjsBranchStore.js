import * as Y from "yjs";
import { applyPatch, diffDocumentStates } from "../patch.js";
import { emptyDocumentState, normalizeDocumentState } from "../state.js";
import { randomUUID } from "../uuid.js";

/**
 * @typedef {import("../types.js").Actor} Actor
 * @typedef {import("../types.js").Branch} Branch
 * @typedef {import("../types.js").Commit} Commit
 * @typedef {import("../types.js").DocumentState} DocumentState
 * @typedef {import("../patch.js").Patch} Patch
 */

const UTF8_ENCODER = new TextEncoder();
const UTF8_DECODER = new TextDecoder();

function isNodeRuntime() {
  const proc = /** @type {any} */ (globalThis.process);
  // Require `process.release.name === "node"` to avoid false positives from lightweight
  // `process` polyfills some bundlers inject into browser environments.
  return Boolean(proc?.versions?.node) && proc?.release?.name === "node";
}

async function importNodeZlib() {
  // Important: keep the specifier non-literal so browser bundlers (esbuild, Rollup, etc)
  // don't try to resolve Node built-ins when producing a browser bundle. This code path
  // is only executed in Node runtimes (guarded by `isNodeRuntime()`).
  const specifier = "zlib";
  return import(/* @vite-ignore */ specifier);
}

/**
 * @param {Uint8Array} bytes
 * @returns {Promise<Uint8Array>}
 */
async function gzipBytes(bytes) {
  if (isNodeRuntime()) {
    const zlib = await importNodeZlib();
    return new Uint8Array(zlib.gzipSync(bytes));
  }

  // eslint-disable-next-line no-undef
  if (typeof CompressionStream !== "undefined") {
    // eslint-disable-next-line no-undef
    const stream = new Blob([bytes]).stream().pipeThrough(new CompressionStream("gzip"));
    // eslint-disable-next-line no-undef
    const compressed = await new Response(stream).arrayBuffer();
    return new Uint8Array(compressed);
  }

  throw new Error("YjsBranchStore: gzip compression is not supported in this environment");
}

/**
 * @param {Uint8Array} bytes
 * @returns {Promise<Uint8Array>}
 */
async function gunzipBytes(bytes) {
  if (isNodeRuntime()) {
    const zlib = await importNodeZlib();
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

  throw new Error("YjsBranchStore: gzip decompression is not supported in this environment");
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
 * @param {unknown} bytes
 * @returns {Uint8Array}
 */
function normalizeBytes(bytes) {
  if (bytes instanceof Uint8Array) return bytes;
  if (bytes instanceof ArrayBuffer) return new Uint8Array(bytes);
  // Older structured clones can return `{ buffer, byteOffset, byteLength }` shapes.
  if (bytes && typeof bytes === "object" && /** @type {any} */ (bytes).buffer instanceof ArrayBuffer) {
    const s = /** @type {any} */ (bytes);
    return new Uint8Array(s.buffer, s.byteOffset ?? 0, s.byteLength);
  }
  throw new Error("YjsBranchStore: invalid bytes");
}

/**
 * @param {unknown} value
 * @returns {Y.Map<any> | null}
 */
function getYMap(value) {
  if (value instanceof Y.Map) return value;

  // See CollabSession#getYMapCell for why we can't rely solely on instanceof.
  if (!value || typeof value !== "object") return null;
  const maybe = /** @type {any} */ (value);
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.set !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  // Plain JS Maps also have get/set/delete, so additionally require Yjs' deep
  // observer APIs.
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return /** @type {Y.Map<any>} */ (maybe);
}

/**
 * @param {unknown} value
 * @returns {Y.Array<any> | null}
 */
function getYArray(value) {
  if (value instanceof Y.Array) return value;
  if (!value || typeof value !== "object") return null;
  const maybe = /** @type {any} */ (value);
  if (typeof maybe.get !== "function") return null;
  if (typeof maybe.toArray !== "function") return null;
  if (typeof maybe.push !== "function") return null;
  if (typeof maybe.delete !== "function") return null;
  if (typeof maybe.observeDeep !== "function") return null;
  if (typeof maybe.unobserveDeep !== "function") return null;
  return /** @type {Y.Array<any>} */ (maybe);
}

function isYAbstractType(value) {
  if (value instanceof Y.AbstractType) return true;
  if (!value || typeof value !== "object") return false;
  const maybe = /** @type {any} */ (value);
  if (typeof maybe.observeDeep !== "function") return false;
  if (typeof maybe.unobserveDeep !== "function") return false;
  return Boolean(maybe._map instanceof Map || maybe._start || maybe._item || maybe._length != null);
}

function replaceForeignRootType({ doc, name, existing, create }) {
  const t = create();

  // Mirror Yjs' own Doc.get conversion logic for AbstractType placeholders, but
  // also support roots instantiated by a different Yjs module instance (e.g.
  // CJS `require("yjs")`).
  t._map = existing?._map;
  t._start = existing?._start;
  t._length = existing?._length;

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
 * Safely access a Map root without relying on `doc.getMap`, which can throw when
 * the root was instantiated by a different Yjs module instance (ESM vs CJS).
 *
 * @param {any} doc
 * @param {string} name
 * @returns {Y.Map<any>}
 */
function getMapRoot(doc, name) {
  const existing = doc?.share?.get?.(name);
  if (!existing) return doc.getMap(name);

  const map = getYMap(existing);
  if (map) {
    if (map instanceof Y.Map) return map;
    if (doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing: map, create: () => new Y.Map() });
    }
    return map;
  }

  if (isYAbstractType(existing)) {
    if (doc instanceof Y.Doc) {
      return replaceForeignRootType({ doc, name, existing, create: () => new Y.Map() });
    }
    return doc.getMap(name);
  }

  throw new Error(`Unsupported Yjs root type for "${name}"`);
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
 * Yjs-backed implementation of the BranchStore interface.
 *
 * Stores the branch + commit graph inside a shared Y.Doc so history syncs and
 * persists automatically via the collaboration layer.
 */
export class YjsBranchStore {
  /** @type {Y.Doc} */
  #ydoc;
  /** @type {string} */
  #rootName;
  /** @type {Y.Map<any>} */
  #branches;
  /** @type {Y.Map<any>} */
  #commits;
  /** @type {Y.Map<any>} */
  #meta;
  /** @type {number | null} */
  #snapshotEveryNCommits;
  /** @type {number | null} */
  #snapshotWhenPatchExceedsBytes;
  /** @type {"json" | "gzip-chunks"} */
  #payloadEncoding;
  /** @type {number} */
  #chunkSize;
  /** @type {number} */
  #maxChunksPerTransaction;

  /**
   * @param {{
   *   ydoc: Y.Doc,
   *   rootName?: string,
   *   snapshotEveryNCommits?: number,
   *   snapshotWhenPatchExceedsBytes?: number,
   *   payloadEncoding?: "json" | "gzip-chunks",
   *   chunkSize?: number,
   *   maxChunksPerTransaction?: number
   * }} input
   */
  constructor({
    ydoc,
    rootName,
    snapshotEveryNCommits,
    snapshotWhenPatchExceedsBytes,
    payloadEncoding,
    chunkSize,
    maxChunksPerTransaction,
  }) {
    if (!ydoc) throw new Error("YjsBranchStore requires { ydoc }");
    this.#ydoc = ydoc;
    this.#rootName = rootName ?? "branching";
    this.#branches = getMapRoot(ydoc, `${this.#rootName}:branches`);
    this.#commits = getMapRoot(ydoc, `${this.#rootName}:commits`);
    this.#meta = getMapRoot(ydoc, `${this.#rootName}:meta`);
    this.#snapshotEveryNCommits =
      snapshotEveryNCommits == null ? 50 : snapshotEveryNCommits;
    this.#snapshotWhenPatchExceedsBytes =
      snapshotWhenPatchExceedsBytes == null ? null : snapshotWhenPatchExceedsBytes;

    /** @type {"json" | "gzip-chunks"} */
    this.#payloadEncoding = payloadEncoding ?? "json";
    this.#chunkSize = chunkSize ?? 64 * 1024;
    this.#maxChunksPerTransaction = maxChunksPerTransaction ?? 16;

    if (this.#payloadEncoding !== "json" && this.#payloadEncoding !== "gzip-chunks") {
      throw new Error(`YjsBranchStore: invalid payloadEncoding: ${String(this.#payloadEncoding)}`);
    }
    if (!Number.isFinite(this.#chunkSize) || !Number.isSafeInteger(this.#chunkSize) || this.#chunkSize <= 0) {
      throw new Error("YjsBranchStore: chunkSize must be a positive integer");
    }
    if (
      !Number.isFinite(this.#maxChunksPerTransaction) ||
      !Number.isSafeInteger(this.#maxChunksPerTransaction) ||
      this.#maxChunksPerTransaction <= 0
    ) {
      throw new Error("YjsBranchStore: maxChunksPerTransaction must be a positive integer");
    }
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
  #inferArrayCtorForMapCtor(mapCtor) {
    const moduleId = abstractTypeSuperclass(mapCtor);
    if (!moduleId) return null;
    const localModuleId = abstractTypeSuperclass(Y.Map);

    // Common case: same module instance.
    if (moduleId === localModuleId) return Y.Array;

    // Try to infer from existing commit chunk arrays (if any).
    /** @type {any | null} */
    let ctor = null;
    this.#commits.forEach((value) => {
      if (ctor) return;
      const commitMap = getYMap(value);
      if (!commitMap) return;
      if (abstractTypeSuperclass(commitMap.constructor) !== moduleId) return;
      const candidates = [commitMap.get("patchChunks"), commitMap.get("snapshotChunks")];
      for (const c of candidates) {
        const arr = getYArray(c);
        if (!arr) continue;
        const arrCtor = arrayConstructor(arr);
        if (abstractTypeSuperclass(arrCtor) === moduleId) {
          ctor = arrCtor;
          return;
        }
      }
    });
    if (ctor) return ctor;

    // Last resort: create a root array to obtain the correct constructor. This
    // remains empty (minimal overhead), but ensures we can always create nested
    // chunk arrays even when this package's `yjs` import is a different module
    // instance than the one that created `ydoc`.
    const probe = this.#ydoc.getArray(`${this.#rootName}:__chunkCtor`);
    const probeCtor = arrayConstructor(probe);
    if (abstractTypeSuperclass(probeCtor) === moduleId) return probeCtor;
    return probeCtor;
  }

  /**
   * @param {any} commitMap
   */
  #isCommitComplete(commitMap) {
    return commitMap.get("commitComplete") !== false;
  }

  /**
   * @param {any} commitMap
   */
  #commitHasPatch(commitMap) {
    const patchEncoding = commitMap.get("patchEncoding");
    if (patchEncoding === "gzip-chunks" || commitMap.get("patchChunks") !== undefined) {
      if (!this.#isCommitComplete(commitMap)) return false;
      const arr = getYArray(commitMap.get("patchChunks"));
      return arr !== null && arr.length > 0;
    }
    return commitMap.get("patch") !== undefined;
  }

  /**
   * @param {any} commitMap
   */
  #commitHasSnapshot(commitMap) {
    if (commitMap.get("snapshot") !== undefined) return true;
    if (!this.#isCommitComplete(commitMap)) return false;
    const arr = getYArray(commitMap.get("snapshotChunks"));
    return arr !== null && arr.length > 0;
  }

  /**
   * Best-effort cleanup for interrupted multi-transaction gzip-chunks writes.
   *
   * Incomplete commits (`commitComplete=false`) are ignored by inference, but can
   * accumulate over time. To avoid deleting legitimate history, we only delete
   * commits when they're clearly unreachable:
   *   - older than a conservative TTL
   *   - not referenced by any branch head
   *   - not referenced by rootCommitId
   *   - not referenced as a parent/merge-parent by any commit
   *   - have a `writeStartedAt` timestamp (set by this store for gzip-chunks writes)
   *
   * @param {string} docId
   */
  #cleanupStaleIncompleteCommits(docId) {
    // Avoid relying on `createdAt` because callers can pass arbitrary commit
    // timestamps (e.g. imported history). `writeStartedAt` is written by this
    // store at commit write time and is safe to compare to wall clock time.
    const ttlMs = 60 * 60 * 1000;
    const now = Date.now();

    /** @type {Set<string>} */
    const referenced = new Set();

    const root = this.#meta.get("rootCommitId");
    if (typeof root === "string" && root.length > 0) referenced.add(root);

    this.#branches.forEach((value) => {
      const branchMap = getYMap(value);
      if (!branchMap) return;
      const head = branchMap.get("headCommitId");
      if (typeof head === "string" && head.length > 0) referenced.add(head);
    });

    // Be extra conservative: don't delete commits that are referenced by any
    // other commit's parent pointers.
    this.#commits.forEach((value) => {
      const commitMap = getYMap(value);
      if (!commitMap) return;
      const parent = commitMap.get("parentCommitId");
      if (typeof parent === "string" && parent.length > 0) referenced.add(parent);
      const mergeParent = commitMap.get("mergeParentCommitId");
      if (typeof mergeParent === "string" && mergeParent.length > 0) referenced.add(mergeParent);
    });

    /** @type {string[]} */
    const staleUnreachable = [];
    this.#commits.forEach((value, key) => {
      if (typeof key !== "string" || key.length === 0) return;
      if (referenced.has(key)) return;
      const commitMap = getYMap(value);
      if (!commitMap) return;
      const commitDocId = commitMap.get("docId");
      if (typeof commitDocId === "string" && commitDocId.length > 0 && commitDocId !== docId) return;
      if (this.#isCommitComplete(commitMap)) return;

      const writeStartedAt = Number(commitMap.get("writeStartedAt") ?? 0);
      if (!Number.isFinite(writeStartedAt) || writeStartedAt <= 0) return;
      if (now - writeStartedAt < ttlMs) return;

      staleUnreachable.push(key);
    });

    if (staleUnreachable.length === 0) return;

    this.#ydoc.transact(() => {
      /** @type {Set<string>} */
      const stillReferenced = new Set();

      const currentRoot = this.#meta.get("rootCommitId");
      if (typeof currentRoot === "string" && currentRoot.length > 0) {
        stillReferenced.add(currentRoot);
      }

      this.#branches.forEach((value) => {
        const branchMap = getYMap(value);
        if (!branchMap) return;
        const head = branchMap.get("headCommitId");
        if (typeof head === "string" && head.length > 0) stillReferenced.add(head);
      });

      this.#commits.forEach((value) => {
        const commitMap = getYMap(value);
        if (!commitMap) return;
        const parent = commitMap.get("parentCommitId");
        if (typeof parent === "string" && parent.length > 0) stillReferenced.add(parent);
        const mergeParent = commitMap.get("mergeParentCommitId");
        if (typeof mergeParent === "string" && mergeParent.length > 0) stillReferenced.add(mergeParent);
      });

      for (const commitId of staleUnreachable) {
        if (stillReferenced.has(commitId)) continue;
        const commitMap = getYMap(this.#commits.get(commitId));
        if (!commitMap) continue;
        if (this.#isCommitComplete(commitMap)) continue;
        this.#commits.delete(commitId);
      }
    }, "branching-store");
  }

  /**
   * @param {unknown} value
   * @returns {Promise<Uint8Array[]>}
   */
  async #encodeJsonToGzipChunks(value) {
    const json = JSON.stringify(value);
    const bytes = UTF8_ENCODER.encode(json);
    const compressed = await gzipBytes(bytes);
    return splitIntoChunks(compressed, this.#chunkSize);
  }

  /**
   * @param {Y.Array<any>} chunksArr
   * @param {string} label
   * @returns {Promise<any>}
   */
  async #decodeJsonFromGzipChunks(chunksArr, label) {
    const chunksRaw = chunksArr.toArray();
    /** @type {Uint8Array[]} */
    const chunks = [];
    for (const chunk of chunksRaw) chunks.push(normalizeBytes(chunk));
    const compressed = concatChunks(chunks);
    const bytes = await gunzipBytes(compressed);
    const json = UTF8_DECODER.decode(bytes);
    try {
      return JSON.parse(json);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      throw new Error(`YjsBranchStore: failed to decode ${label}: ${message}`);
    }
  }

  /**
   * @param {Y.Map<any>} commitMap
   * @param {string} commitId
   * @returns {Promise<Patch>}
   */
  async #readCommitPatch(commitMap, commitId) {
    const patchEncoding = commitMap.get("patchEncoding");
    if (patchEncoding === "gzip-chunks" || commitMap.get("patchChunks") !== undefined) {
      if (!this.#isCommitComplete(commitMap)) {
        throw new Error(`Commit not fully written yet: ${commitId}`);
      }
      const chunksArr = getYArray(commitMap.get("patchChunks"));
      if (!chunksArr) {
        throw new Error(`YjsBranchStore: missing patchChunks for ${commitId}`);
      }
      const decoded = await this.#decodeJsonFromGzipChunks(chunksArr, `patch(${commitId})`);
      return /** @type {Patch} */ (decoded);
    }

    return structuredClone(commitMap.get("patch") ?? { schemaVersion: 1 });
  }

  /**
   * @param {Y.Map<any>} commitMap
   * @param {string} commitId
   * @returns {Promise<DocumentState | null>}
   */
  async #readCommitSnapshot(commitMap, commitId) {
    const snapshot = commitMap.get("snapshot");
    if (snapshot !== undefined) {
      return normalizeDocumentState(structuredClone(snapshot));
    }

    const snapshotEncoding = commitMap.get("snapshotEncoding");
    if (snapshotEncoding === "gzip-chunks" || commitMap.get("snapshotChunks") !== undefined) {
      if (!this.#isCommitComplete(commitMap)) {
        // Snapshot payloads are an optimization; if the patch is stored inline we
        // can still reconstruct the commit state by replaying patches even while
        // an interrupted snapshot chunk write is being repaired.
        if (this.#commitHasPatch(commitMap)) return null;
        throw new Error(`Commit not fully written yet: ${commitId}`);
      }
      const chunksArr = getYArray(commitMap.get("snapshotChunks"));
      if (!chunksArr) return null;
      const decoded = await this.#decodeJsonFromGzipChunks(chunksArr, `snapshot(${commitId})`);
      return normalizeDocumentState(decoded);
    }

    return null;
  }

  /**
   * @param {string} docId
   * @param {Actor} actor
   * @param {DocumentState} initialState
   */
  async ensureDocument(docId, actor, initialState) {
    let existingRoot = this.#meta.get("rootCommitId");
    if (typeof existingRoot !== "string" || existingRoot.length === 0) {
      const inferredRoot = this.#inferRootCommitId(docId);
      if (inferredRoot) {
        this.#ydoc.transact(() => {
          const current = this.#meta.get("rootCommitId");
          if (typeof current === "string" && current.length > 0) return;
          this.#meta.set("rootCommitId", inferredRoot);
        });
        existingRoot = inferredRoot;
      }
    }

    if (typeof existingRoot === "string" && existingRoot.length > 0) {
      /** @type {string} */
      let rootCommitId = existingRoot;
      /** @type {Y.Map<any> | null} */
      let rootCommitMap = getYMap(this.#commits.get(rootCommitId));
      if (!rootCommitMap) throw new Error(`Root commit not found: ${rootCommitId}`);

      // Resilience: interrupted gzip-chunks writes can leave partially-written
      // commits around (commitComplete=false). Ensure document repair doesn't
      // select those as the root commit pointer.
      if (!this.#isCommitComplete(rootCommitMap)) {
        // Some repairs/migrations can write chunk payloads across multiple
        // transactions (e.g. restoring a missing root snapshot). If we crashed
        // mid-write, `commitComplete=false` may linger even though the commit still
        // contains an inline JSON patch/snapshot that lets us safely proceed and
        // finish the migration.
        const hasReadablePayload =
          this.#commitHasPatch(rootCommitMap) || this.#commitHasSnapshot(rootCommitMap);
        if (!hasReadablePayload) {
          const inferredRoot = this.#inferRootCommitId(docId);
          if (
            typeof inferredRoot === "string" &&
            inferredRoot.length > 0 &&
            inferredRoot !== rootCommitId
          ) {
            this.#ydoc.transact(() => {
              const current = this.#meta.get("rootCommitId");
              if (current !== rootCommitId) return;
              this.#meta.set("rootCommitId", inferredRoot);
            });
            rootCommitId = inferredRoot;
            existingRoot = inferredRoot;
            rootCommitMap = getYMap(this.#commits.get(rootCommitId));
            if (!rootCommitMap) throw new Error(`Root commit not found: ${rootCommitId}`);
          } else {
            throw new Error(
              `YjsBranchStore: root commit not fully written yet (commitComplete=false): ${rootCommitId}`
            );
          }
        }
      }

      const rootDocId = rootCommitMap.get("docId");
      if (typeof rootDocId === "string" && rootDocId.length > 0 && rootDocId !== docId) {
        throw new Error(
          `YjsBranchStore docId mismatch: requested ${docId}, but ydoc history is for ${rootDocId}`
        );
      }

      // Migration: older docs may not have stored commit snapshots. (Snapshots are
      // an optimization; we can still replay patches, but restoring at least the
      // root snapshot helps keep history reads bounded.)
      const needsSnapshot = !this.#commitHasSnapshot(rootCommitMap) && this.#commitHasPatch(rootCommitMap);
      const snapshot = needsSnapshot
        ? this._applyPatch(
            emptyDocumentState(),
            await this.#readCommitPatch(rootCommitMap, rootCommitId),
          )
        : null;

      const headForMissingMainBranch = this.#inferLatestCommitId(docId) ?? rootCommitId;

      this.#ydoc.transact(() => {
        // Backwards-compatible migration: ensure main branch exists.
        let mainBranch = getYMap(this.#branches.get("main"));
        if (!mainBranch) {
          const createdBy =
            typeof rootCommitMap.get("createdBy") === "string" ? rootCommitMap.get("createdBy") : actor.userId;
          const createdAt = Number(rootCommitMap.get("createdAt") ?? Date.now());
          const BranchCtor = mapConstructor(this.#branches);
          const main = new BranchCtor();
          main.set("id", randomUUID());
          main.set("docId", docId);
          main.set("name", "main");
          main.set("createdBy", createdBy);
          main.set("createdAt", createdAt);
          main.set("description", null);
          main.set("headCommitId", headForMissingMainBranch);
          this.#branches.set("main", main);
          mainBranch = main;
        } else if (String(mainBranch.get("docId") ?? "") !== docId) {
          throw new Error(
            `YjsBranchStore docId mismatch: requested ${docId}, but branch 'main' is for ${String(mainBranch.get("docId") ?? "")}`
          );
        }

        // Defensive repair: ensure main branch head points at a readable commit.
        // Some corrupted docs can end up with a missing headCommitId, a dangling
        // commit reference, or a partially-written gzip-chunks commit id.
        const head = mainBranch.get("headCommitId");
        let needsHeadRepair = typeof head !== "string" || head.length === 0;
        if (!needsHeadRepair) {
          const headCommit = getYMap(this.#commits.get(head));
          if (!headCommit) {
            needsHeadRepair = true;
          } else {
            const headDocId = headCommit.get("docId");
            if (typeof headDocId === "string" && headDocId.length > 0 && headDocId !== docId) {
              needsHeadRepair = true;
            } else if (!this.#isCommitComplete(headCommit)) {
              needsHeadRepair = true;
            } else {
              const hasPayload = this.#commitHasPatch(headCommit) || this.#commitHasSnapshot(headCommit);
              if (!hasPayload) needsHeadRepair = true;
            }
          }
        }
        if (needsHeadRepair) {
          mainBranch.set("headCommitId", headForMissingMainBranch);
        }

        // Backwards-compatible migration: ensure current branch name exists + is valid.
        let current = this.#meta.get("currentBranchName");
        if (typeof current !== "string" || current.length === 0) current = "main";
        if (!getYMap(this.#branches.get(current))) current = "main";
        this.#meta.set("currentBranchName", current);

        // Persist a recovered root snapshot for older histories. For gzip-chunks
        // payloads, we write in a separate pass below so we can stream chunks
        // without inflating message sizes.
      });

      if (snapshot) {
        if (this.#payloadEncoding === "gzip-chunks") {
          const snapshotChunks = await this.#encodeJsonToGzipChunks(snapshot);
          const MapCtor = mapConstructor(this.#commits);
          const arrayCtor = this.#inferArrayCtorForMapCtor(MapCtor);
          if (!arrayCtor) return;

          let snapshotArr = null;
          let snapshotIndex = 0;

          const pushMore = () => {
            if (snapshotIndex >= snapshotChunks.length) return true;
            const remaining = Math.max(1, this.#maxChunksPerTransaction);
            const slice = snapshotChunks.slice(snapshotIndex, snapshotIndex + remaining);
            snapshotIndex += slice.length;
            snapshotArr.push(slice);
            return snapshotIndex >= snapshotChunks.length;
          };

          this.#ydoc.transact(() => {
            const commit = getYMap(this.#commits.get(rootCommitId));
            if (!commit) return;
            if (this.#commitHasSnapshot(commit)) return;

            snapshotArr = new arrayCtor();
            commit.set("commitComplete", false);
            commit.set("snapshotEncoding", "gzip-chunks");
            commit.set("snapshotChunks", snapshotArr);

            const done = pushMore();
            if (done) commit.set("commitComplete", true);
          }, "branching-store");

          while (snapshotArr && snapshotIndex < snapshotChunks.length) {
            this.#ydoc.transact(() => {
              const commit = getYMap(this.#commits.get(rootCommitId));
              if (!commit) return;
              const arr = getYArray(commit.get("snapshotChunks"));
              if (!arr) return;
              snapshotArr = arr;
              commit.set("commitComplete", false);
              const done = pushMore();
              if (done) commit.set("commitComplete", true);
            }, "branching-store");
          }
        } else {
          this.#ydoc.transact(() => {
            const commit = getYMap(this.#commits.get(rootCommitId));
            if (commit && commit.get("snapshot") === undefined) {
              commit.set("snapshot", structuredClone(snapshot));
            }
          }, "branching-store");
        }
      }

      this.#cleanupStaleIncompleteCommits(docId);
      return;
    }

    const now = Date.now();
    const rootCommitId = randomUUID();
    const mainBranchId = randomUUID();

    /** @type {Patch} */
    const patch = diffDocumentStates(emptyDocumentState(), initialState);
    const snapshot = this._applyPatch(emptyDocumentState(), patch);

    if (this.#payloadEncoding === "gzip-chunks") {
      const [patchChunks, snapshotChunks] = await Promise.all([
        this.#encodeJsonToGzipChunks(patch),
        this.#encodeJsonToGzipChunks(snapshot),
      ]);

      const CommitCtor = mapConstructor(this.#commits);
      const BranchCtor = mapConstructor(this.#branches);
      const arrayCtor = this.#inferArrayCtorForMapCtor(CommitCtor) ?? Y.Array;

      let created = false;
      let patchArr = null;
      let snapshotArr = null;
      let patchIndex = 0;
      let snapshotIndex = 0;

      const pushMore = () => {
        let remaining = this.#maxChunksPerTransaction;
        if (patchArr && patchIndex < patchChunks.length && remaining > 0) {
          const slice = patchChunks.slice(patchIndex, patchIndex + remaining);
          patchIndex += slice.length;
          remaining -= slice.length;
          if (slice.length > 0) patchArr.push(slice);
        }
        if (snapshotArr && snapshotIndex < snapshotChunks.length && remaining > 0) {
          const slice = snapshotChunks.slice(snapshotIndex, snapshotIndex + remaining);
          snapshotIndex += slice.length;
          remaining -= slice.length;
          if (slice.length > 0) snapshotArr.push(slice);
        }
        return patchIndex >= patchChunks.length && snapshotIndex >= snapshotChunks.length;
      };

      this.#ydoc.transact(() => {
        const rootAfter = this.#meta.get("rootCommitId");
        if (typeof rootAfter === "string" && rootAfter.length > 0) return;

        const commit = new CommitCtor();
        commit.set("id", rootCommitId);
        commit.set("docId", docId);
        commit.set("parentCommitId", null);
        commit.set("mergeParentCommitId", null);
        commit.set("createdBy", actor.userId);
        commit.set("createdAt", now);
        commit.set("writeStartedAt", now);
        commit.set("message", "root");
        commit.set("patchEncoding", "gzip-chunks");
        commit.set("snapshotEncoding", "gzip-chunks");
        commit.set("commitComplete", false);

        patchArr = new arrayCtor();
        snapshotArr = new arrayCtor();
        commit.set("patchChunks", patchArr);
        commit.set("snapshotChunks", snapshotArr);

        const done = pushMore();
        if (done) commit.set("commitComplete", true);

        this.#commits.set(rootCommitId, commit);

        const main = new BranchCtor();
        main.set("id", mainBranchId);
        main.set("docId", docId);
        main.set("name", "main");
        main.set("createdBy", actor.userId);
        main.set("createdAt", now);
        main.set("description", null);
        main.set("headCommitId", rootCommitId);
        this.#branches.set("main", main);

        this.#meta.set("rootCommitId", rootCommitId);
        this.#meta.set("currentBranchName", "main");

        created = true;
      }, "branching-store");

      if (!created) return;

      while (patchIndex < patchChunks.length || snapshotIndex < snapshotChunks.length) {
        this.#ydoc.transact(() => {
          const commit = getYMap(this.#commits.get(rootCommitId));
          if (!commit) return;
          patchArr = getYArray(commit.get("patchChunks"));
          snapshotArr = getYArray(commit.get("snapshotChunks"));
          if (!patchArr || !snapshotArr) return;

          commit.set("commitComplete", false);
          const done = pushMore();
          if (done) commit.set("commitComplete", true);
        }, "branching-store");
      }

      return;
    }

    this.#ydoc.transact(() => {
      const rootAfter = this.#meta.get("rootCommitId");
      if (typeof rootAfter === "string" && rootAfter.length > 0) return;

      const CommitCtor = mapConstructor(this.#commits);
      const commit = new CommitCtor();
      commit.set("id", rootCommitId);
      commit.set("docId", docId);
      commit.set("parentCommitId", null);
      commit.set("mergeParentCommitId", null);
      commit.set("createdBy", actor.userId);
      commit.set("createdAt", now);
      commit.set("message", "root");
      commit.set("patch", structuredClone(patch));
      commit.set("snapshot", structuredClone(snapshot));
      this.#commits.set(rootCommitId, commit);

      const BranchCtor = mapConstructor(this.#branches);
      const main = new BranchCtor();
      main.set("id", mainBranchId);
      main.set("docId", docId);
      main.set("name", "main");
      main.set("createdBy", actor.userId);
      main.set("createdAt", now);
      main.set("description", null);
      main.set("headCommitId", rootCommitId);
      this.#branches.set("main", main);

      this.#meta.set("rootCommitId", rootCommitId);
      this.#meta.set("currentBranchName", "main");
    });
  }

  /**
   * Best-effort recovery for corrupted/migrating docs: infer the root commit id
   * even if `branching:meta.rootCommitId` is missing.
   *
   * @param {string} docId
   * @returns {string | null}
   */
  #inferRootCommitId(docId) {
    const mainBranch = getYMap(this.#branches.get("main"));
    if (mainBranch) {
      const branchDocId = mainBranch.get("docId");
      if (typeof branchDocId === "string" && branchDocId.length > 0 && branchDocId !== docId) {
        return null;
      }
      let currentId = mainBranch.get("headCommitId");
      /** @type {Set<string>} */
      const seen = new Set();
      while (typeof currentId === "string" && currentId.length > 0) {
        if (seen.has(currentId)) break;
        seen.add(currentId);
        const commitMap = getYMap(this.#commits.get(currentId));
        if (!commitMap) break;
        const commitDocId = commitMap.get("docId");
        if (typeof commitDocId === "string" && commitDocId.length > 0 && commitDocId !== docId) break;
        const parent = commitMap.get("parentCommitId");
        if (!parent) {
          // Avoid selecting partially-written gzip-chunks commits as the root.
          // However, some repairs (e.g. root snapshot migration) temporarily set
          // `commitComplete=false` even though the commit still has an inline
          // patch/snapshot payload that makes it safe to use.
          const hasPayload = this.#commitHasPatch(commitMap) || this.#commitHasSnapshot(commitMap);
          if (!this.#isCommitComplete(commitMap) && !hasPayload) break;
          if (!hasPayload) break;
          return currentId;
        }
        currentId = parent;
      }
    }

    /** @type {{ id: string, createdAt: number, hasPayload: boolean, isComplete: boolean }[]} */
    const candidates = [];
    this.#commits.forEach((value, key) => {
      const commitMap = getYMap(value);
      if (!commitMap) return;
      const parent = commitMap.get("parentCommitId");
      if (parent !== null && parent !== undefined) return;

      const commitDocId = commitMap.get("docId");
      if (typeof commitDocId === "string" && commitDocId.length > 0 && commitDocId !== docId) return;
      const hasPayload = this.#commitHasPatch(commitMap) || this.#commitHasSnapshot(commitMap);
      if (!this.#isCommitComplete(commitMap) && !hasPayload) return;
      const createdAt = Number(commitMap.get("createdAt") ?? 0);
      const isComplete = this.#isCommitComplete(commitMap);
      if (typeof key === "string" && key.length > 0) {
        candidates.push({ id: key, createdAt, hasPayload, isComplete });
      }
    });

    candidates.sort((a, b) => {
      if (a.isComplete !== b.isComplete) return a.isComplete ? -1 : 1;
      if (a.hasPayload !== b.hasPayload) return a.hasPayload ? -1 : 1;
      return a.createdAt - b.createdAt || a.id.localeCompare(b.id);
    });
    return candidates[0]?.id ?? null;
  }

  /**
   * Infer a reasonable branch head for a repaired "main" branch.
   *
   * @param {string} docId
   * @returns {string | null}
   */
  #inferLatestCommitId(docId) {
    /** @type {{ id: string, createdAt: number, hasPayload: boolean, isComplete: boolean }[]} */
    const candidates = [];
    this.#commits.forEach((value, key) => {
      const commitMap = getYMap(value);
      if (!commitMap) return;
      const commitDocId = commitMap.get("docId");
      if (typeof commitDocId === "string" && commitDocId.length > 0 && commitDocId !== docId) return;
      const hasPayload = this.#commitHasPatch(commitMap) || this.#commitHasSnapshot(commitMap);
      if (!this.#isCommitComplete(commitMap) && !hasPayload) return;
      const createdAt = Number(commitMap.get("createdAt") ?? 0);
      const isComplete = this.#isCommitComplete(commitMap);
      if (typeof key === "string" && key.length > 0) {
        candidates.push({ id: key, createdAt, hasPayload, isComplete });
      }
    });

    candidates.sort((a, b) => {
      if (a.isComplete !== b.isComplete) return a.isComplete ? -1 : 1;
      if (a.hasPayload !== b.hasPayload) return a.hasPayload ? -1 : 1;
      return b.createdAt - a.createdAt || a.id.localeCompare(b.id);
    });
    return candidates[0]?.id ?? null;
  }

  /**
   * @param {string} docId
   * @returns {Promise<boolean>}
   */
  async hasDocument(docId) {
    let root = this.#meta.get("rootCommitId");
    if (typeof root !== "string" || root.length === 0) {
      root = this.#inferRootCommitId(docId);
    }
    if (typeof root !== "string" || root.length === 0) return false;
    const commit = getYMap(this.#commits.get(root));
    if (!commit) return false;
    const commitDocId = commit.get("docId");
    if (typeof commitDocId === "string" && commitDocId.length > 0) {
      return commitDocId === docId;
    }
    const mainBranch = getYMap(this.#branches.get("main"));
    const mainDocId = mainBranch?.get("docId");
    if (typeof mainDocId === "string" && mainDocId.length > 0) {
      return mainDocId === docId;
    }
    return true;
  }

  async getCurrentBranchName(docId) {
    const raw = this.#meta.get("currentBranchName");
    const name = typeof raw === "string" && raw.length > 0 ? raw : "main";
    const branch = getYMap(this.#branches.get(name));
    if (branch && String(branch.get("docId") ?? "") === docId) return name;

    const main = getYMap(this.#branches.get("main"));
    if (main && String(main.get("docId") ?? "") === docId) {
      // Self-heal invalid pointers so other collaborators don't keep reading a
      // dangling branch name.
      if (raw !== "main") {
        this.#ydoc.transact(() => {
          const current = this.#meta.get("currentBranchName");
          if (current === "main") return;
          this.#meta.set("currentBranchName", "main");
        });
      }
      return "main";
    }

    return "main";
  }

  /**
   * @param {string} _docId
   * @param {string} name
   */
  async setCurrentBranchName(docId, name) {
    this.#ydoc.transact(() => {
      const branchMap = getYMap(this.#branches.get(name));
      if (!branchMap) {
        throw new Error(`Branch not found: ${name}`);
      }
      if (String(branchMap.get("docId") ?? "") !== docId) {
        throw new Error(`Branch not found: ${name}`);
      }
      this.#meta.set("currentBranchName", name);
    });
  }

  /**
   * @param {Y.Map<any>} branchMap
   * @returns {Branch}
   */
  #branchFromYMap(branchMap) {
    return {
      id: String(branchMap.get("id") ?? ""),
      docId: String(branchMap.get("docId") ?? ""),
      name: String(branchMap.get("name") ?? ""),
      createdBy: String(branchMap.get("createdBy") ?? ""),
      createdAt: Number(branchMap.get("createdAt") ?? 0),
      description: (branchMap.get("description") ?? null) === null ? null : String(branchMap.get("description")),
      headCommitId: String(branchMap.get("headCommitId") ?? "")
    };
  }

  /**
   * @param {Y.Map<any>} commitMap
   * @returns {Omit<Commit, "patch">}
   */
  #commitMetaFromYMap(commitMap) {
    return {
      id: String(commitMap.get("id") ?? ""),
      docId: String(commitMap.get("docId") ?? ""),
      parentCommitId: commitMap.get("parentCommitId") ?? null,
      mergeParentCommitId: commitMap.get("mergeParentCommitId") ?? null,
      createdBy: String(commitMap.get("createdBy") ?? ""),
      createdAt: Number(commitMap.get("createdAt") ?? 0),
      message: commitMap.get("message") ?? null,
    };
  }

  /**
   * @param {string} docId
   * @returns {Promise<Branch[]>}
   */
  async listBranches(docId) {
    /** @type {Branch[]} */
    const out = [];
    this.#branches.forEach((value) => {
      const branchMap = getYMap(value);
      if (!branchMap) return;
      const branch = this.#branchFromYMap(branchMap);
      if (branch.docId !== docId) return;
      out.push(branch);
    });
    out.sort((a, b) => (a.createdAt - b.createdAt === 0 ? a.name.localeCompare(b.name) : a.createdAt - b.createdAt));
    return structuredClone(out);
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @returns {Promise<Branch | null>}
   */
  async getBranch(docId, name) {
    const branchMap = getYMap(this.#branches.get(name));
    if (!branchMap) return null;
    const branch = this.#branchFromYMap(branchMap);
    if (branch.docId !== docId) return null;
    return structuredClone(branch);
  }

  /**
   * @param {{ docId: string, name: string, createdBy: string, createdAt: number, description: string | null, headCommitId: string }} input
   * @returns {Promise<Branch>}
   */
  async createBranch({ docId, name, createdBy, createdAt, description, headCommitId }) {
    if (this.#branches.has(name)) {
      throw new Error(`Branch already exists: ${name}`);
    }

    const id = randomUUID();
    this.#ydoc.transact(() => {
      if (this.#branches.has(name)) {
        throw new Error(`Branch already exists: ${name}`);
      }
      const BranchCtor = mapConstructor(this.#branches);
      const branch = new BranchCtor();
      branch.set("id", id);
      branch.set("docId", docId);
      branch.set("name", name);
      branch.set("createdBy", createdBy);
      branch.set("createdAt", createdAt);
      branch.set("description", description ?? null);
      branch.set("headCommitId", headCommitId);
      this.#branches.set(name, branch);
    });

    return {
      id,
      docId,
      name,
      createdBy,
      createdAt,
      description: description ?? null,
      headCommitId
    };
  }

  /**
   * @param {string} docId
   * @param {string} oldName
   * @param {string} newName
   */
  async renameBranch(docId, oldName, newName) {
    this.#ydoc.transact(() => {
      if (this.#branches.has(newName)) {
        throw new Error(`Branch already exists: ${newName}`);
      }

      const branchMap = getYMap(this.#branches.get(oldName));
      if (!branchMap) throw new Error(`Branch not found: ${oldName}`);
      if (String(branchMap.get("docId") ?? "") !== docId) {
        throw new Error(`Branch not found: ${oldName}`);
      }

      const BranchCtor = mapConstructor(branchMap);
      const next = new BranchCtor();
      branchMap.forEach((v, k) => {
        if (k === "name") return;
        next.set(k, v);
      });
      next.set("name", newName);

      this.#branches.delete(oldName);
      this.#branches.set(newName, next);

      // Keep the global checked-out branch pointer consistent within the same
      // transaction so other collaborators never observe a dangling
      // `currentBranchName` that points at a non-existent branch.
      if (this.#meta.get("currentBranchName") === oldName) {
        this.#meta.set("currentBranchName", newName);
      }
    });
  }

  /**
   * @param {string} docId
   * @param {string} name
   */
  async deleteBranch(docId, name) {
    this.#ydoc.transact(() => {
      const branchMap = getYMap(this.#branches.get(name));
      if (!branchMap) return;
      if (String(branchMap.get("docId") ?? "") !== docId) return;
      this.#branches.delete(name);

      // Defensive: if a caller bypasses BranchService and deletes the
      // checked-out branch, fall back to main.
      if (this.#meta.get("currentBranchName") === name) {
        this.#meta.set("currentBranchName", "main");
      }
    });
  }

  /**
   * @param {string} docId
   * @param {string} name
   * @param {string} headCommitId
   */
  async updateBranchHead(docId, name, headCommitId) {
    this.#ydoc.transact(() => {
      const branchMap = getYMap(this.#branches.get(name));
      if (!branchMap) throw new Error(`Branch not found: ${name}`);
      if (String(branchMap.get("docId") ?? "") !== docId) {
        throw new Error(`Branch not found: ${name}`);
      }
      branchMap.set("headCommitId", headCommitId);
    });
  }

  /**
   * @param {{ docId: string, parentCommitId: string | null, mergeParentCommitId: string | null, createdBy: string, createdAt: number, message: string | null, patch: Patch, nextState?: DocumentState }} input
   * @returns {Promise<Commit>}
   */
  async createCommit({
    docId,
    parentCommitId,
    mergeParentCommitId,
    createdBy,
    createdAt,
    message,
    patch,
    nextState,
  }) {
    const id = randomUUID();

    const shouldSnapshot = await this.#shouldSnapshotCommit({ parentCommitId, patch });
    const snapshotState = shouldSnapshot
      ? await this.#resolveSnapshotState({ parentCommitId, patch, nextState })
      : null;

    if (this.#payloadEncoding === "gzip-chunks") {
      const patchChunks = await this.#encodeJsonToGzipChunks(patch);
      const snapshotChunks = snapshotState ? await this.#encodeJsonToGzipChunks(snapshotState) : null;

      const CommitCtor = mapConstructor(this.#commits);
      const arrayCtor = this.#inferArrayCtorForMapCtor(CommitCtor) ?? Y.Array;

      let patchArr = null;
      let snapshotArr = null;
      let patchIndex = 0;
      let snapshotIndex = 0;

      const pushMore = () => {
        let remaining = this.#maxChunksPerTransaction;
        if (patchArr && patchIndex < patchChunks.length && remaining > 0) {
          const slice = patchChunks.slice(patchIndex, patchIndex + remaining);
          patchIndex += slice.length;
          remaining -= slice.length;
          if (slice.length > 0) patchArr.push(slice);
        }
        if (snapshotArr && snapshotChunks && snapshotIndex < snapshotChunks.length && remaining > 0) {
          const slice = snapshotChunks.slice(snapshotIndex, snapshotIndex + remaining);
          snapshotIndex += slice.length;
          remaining -= slice.length;
          if (slice.length > 0) snapshotArr.push(slice);
        }
        return patchIndex >= patchChunks.length && (!snapshotChunks || snapshotIndex >= snapshotChunks.length);
      };

      this.#ydoc.transact(() => {
        const commit = new CommitCtor();
        commit.set("id", id);
        commit.set("docId", docId);
        commit.set("parentCommitId", parentCommitId);
        commit.set("mergeParentCommitId", mergeParentCommitId);
        commit.set("createdBy", createdBy);
        commit.set("createdAt", createdAt);
        commit.set("writeStartedAt", Date.now());
        commit.set("message", message ?? null);
        commit.set("patchEncoding", "gzip-chunks");
        commit.set("commitComplete", false);

        patchArr = new arrayCtor();
        commit.set("patchChunks", patchArr);

        if (snapshotChunks) {
          snapshotArr = new arrayCtor();
          commit.set("snapshotEncoding", "gzip-chunks");
          commit.set("snapshotChunks", snapshotArr);
        }

        const done = pushMore();
        if (done) commit.set("commitComplete", true);

        this.#commits.set(id, commit);
      }, "branching-store");

      while (patchIndex < patchChunks.length || (snapshotChunks && snapshotIndex < snapshotChunks.length)) {
        this.#ydoc.transact(() => {
          const commit = getYMap(this.#commits.get(id));
          if (!commit) return;
          patchArr = getYArray(commit.get("patchChunks"));
          snapshotArr = snapshotChunks ? getYArray(commit.get("snapshotChunks")) : null;
          if (!patchArr) return;

          commit.set("commitComplete", false);
          const done = pushMore();
          if (done) commit.set("commitComplete", true);
        }, "branching-store");
      }
    } else {
      this.#ydoc.transact(() => {
        const CommitCtor = mapConstructor(this.#commits);
        const commit = new CommitCtor();
        commit.set("id", id);
        commit.set("docId", docId);
        commit.set("parentCommitId", parentCommitId);
        commit.set("mergeParentCommitId", mergeParentCommitId);
        commit.set("createdBy", createdBy);
        commit.set("createdAt", createdAt);
        commit.set("message", message ?? null);
        commit.set("patch", structuredClone(patch));
        if (snapshotState) commit.set("snapshot", structuredClone(snapshotState));
        this.#commits.set(id, commit);
      }, "branching-store");
    }

    return {
      id,
      docId,
      parentCommitId,
      mergeParentCommitId,
      createdBy,
      createdAt,
      message: message ?? null,
      patch: structuredClone(patch)
    };
  }

  /**
   * @param {string} commitId
   * @returns {Promise<Commit | null>}
   */
  async getCommit(commitId) {
    const commitMap = getYMap(this.#commits.get(commitId));
    if (!commitMap) return null;
    const meta = this.#commitMetaFromYMap(commitMap);
    const patch = await this.#readCommitPatch(commitMap, commitId);
    return {
      ...structuredClone(meta),
      patch: structuredClone(patch),
    };
  }

  /**
   * @param {string} commitId
   * @returns {Promise<DocumentState>}
   */
  async getDocumentStateAtCommit(commitId) {
    const direct = getYMap(this.#commits.get(commitId));
    if (!direct) throw new Error(`Commit not found: ${commitId}`);

    const directSnapshot = await this.#readCommitSnapshot(direct, commitId);
    if (directSnapshot) return directSnapshot;

    /** @type {Patch[]} */
    const chain = [];
    let currentId = commitId;
    /** @type {DocumentState | null} */
    let baseSnapshot = null;

    while (currentId) {
      const commitMap = getYMap(this.#commits.get(currentId));
      if (!commitMap) throw new Error(`Commit not found: ${currentId}`);

      const snapshot = await this.#readCommitSnapshot(commitMap, currentId);
      if (snapshot) {
        baseSnapshot = snapshot;
        break;
      }

      chain.push(await this.#readCommitPatch(commitMap, currentId));

      const parent = commitMap.get("parentCommitId");
      if (!parent) break;
      currentId = parent;
    }

    chain.reverse();

    /** @type {DocumentState} */
    let state = baseSnapshot ?? emptyDocumentState();
    for (const patch of chain) {
      state = this._applyPatch(state, patch);
    }
    return state;
  }

  /**
   * @param {DocumentState} state
   * @param {Patch} patch
   * @returns {DocumentState}
   */
  _applyPatch(state, patch) {
    return applyPatch(state, patch);
  }

  async #shouldSnapshotCommit({ parentCommitId, patch }) {
    if (this.#snapshotWhenPatchExceedsBytes != null && this.#snapshotWhenPatchExceedsBytes > 0) {
      const patchBytes = UTF8_ENCODER.encode(JSON.stringify(patch)).length;
      if (patchBytes > this.#snapshotWhenPatchExceedsBytes) return true;
    }

    if (this.#snapshotEveryNCommits != null && this.#snapshotEveryNCommits > 0) {
      const distance = this.#distanceFromSnapshotCommit(parentCommitId);
      if (distance + 1 >= this.#snapshotEveryNCommits) return true;
    }

    return false;
  }

  #distanceFromSnapshotCommit(startCommitId) {
    if (!startCommitId) return 0;
    let distance = 0;
    let currentId = startCommitId;
    while (currentId) {
      const commitMap = getYMap(this.#commits.get(currentId));
      if (!commitMap) throw new Error(`Commit not found: ${currentId}`);
      if (this.#commitHasSnapshot(commitMap)) return distance;
      const parentId = commitMap.get("parentCommitId");
      if (!parentId) return distance;
      distance += 1;
      currentId = parentId;
    }
    return distance;
  }

  async #resolveSnapshotState({ parentCommitId, patch, nextState }) {
    if (nextState) return normalizeDocumentState(nextState);
    const base = parentCommitId ? await this.getDocumentStateAtCommit(parentCommitId) : emptyDocumentState();
    return this._applyPatch(base, patch);
  }
}
