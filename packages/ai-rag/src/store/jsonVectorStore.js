import { fromBase64, InMemoryBinaryStorage, toBase64 } from "./binaryStorage.js";
import { InMemoryVectorStore } from "./inMemoryVectorStore.js";
import { throwIfAborted } from "../utils/abort.js";

const JSON_VECTOR_STORE_VERSION = 2;

/**
 * @param {Float32Array} vector
 */
function float32VectorToBase64(vector) {
  const bytes = new Uint8Array(vector.buffer, vector.byteOffset, vector.byteLength);
  return toBase64(bytes);
}

/**
 * @param {string} encoded
 */
function base64ToFloat32Vector(encoded) {
  const bytes = fromBase64(encoded);
  if (bytes.byteLength % 4 !== 0) {
    throw new Error(`Invalid vector_b64 payload: expected byteLength multiple of 4, got ${bytes.byteLength}`);
  }
  return new Float32Array(bytes.buffer, bytes.byteOffset, bytes.byteLength / 4);
}

/**
 * A small, dependency-free persistent vector store.
 *
 * This keeps a full in-memory copy for queries, then snapshots state to the
 * provided storage implementation on mutation. It is not intended for very
 * large corpora, but provides workbook-scale persistence in environments that
 * don't support Node filesystem APIs (e.g. browser, Tauri webviews).
 */
export class JsonVectorStore extends InMemoryVectorStore {
  /**
   * @param {{
   *   storage?: any,
   *   dimension: number,
   *   autoSave?: boolean,
   *   resetOnCorrupt?: boolean
   * }} opts
   * @param {boolean} [opts.resetOnCorrupt]
   *   When true, invalid persisted payloads are cleared (via `storage.remove()` when
   *   available) and the store loads as empty. When false, invalid payloads are
   *   ignored but left in storage (matching the historical behaviour).
   */
  constructor(opts) {
    super({ dimension: opts.dimension });
    this._storage = opts.storage ?? new InMemoryBinaryStorage();
    this._autoSave = opts.autoSave ?? true;
    this._resetOnCorrupt = opts.resetOnCorrupt ?? true;
    this._loaded = false;
    this._dirty = false;
    /** @type {Promise<void> | null} */
    this._loadPromise = null;
    // Serialize `storage.save()` calls to prevent lost updates when multiple
    // mutations trigger overlapping async persists.
    /** @type {Promise<void>} */
    this._persistQueue = Promise.resolve();
    // Monotonic counter so we only clear `_dirty` when the persisted snapshot
    // corresponds to the latest mutation at the time the persist started.
    this._mutationVersion = 0;
    this._batchDepth = 0;
  }

  async _maybeResetOnCorrupt() {
    if (!this._resetOnCorrupt) return;
    let cleared = false;

    const remove = this._storage?.remove;
    if (typeof remove === "function") {
      try {
        await remove.call(this._storage);
        cleared = true;
      } catch {
        // ignore
      }
    }

    // Some BinaryStorage implementations may not support `remove()`. In that case,
    // best-effort overwrite the corrupted payload with a valid empty snapshot so
    // subsequent loads don't repeatedly hit the corruption path.
    if (!cleared) {
      try {
        const payload = JSON.stringify({
          version: JSON_VECTOR_STORE_VERSION,
          dimension: this.dimension,
          records: [],
        });
        await this._storage.save(new TextEncoder().encode(payload));
      } catch {
        // ignore
      }
    }
  }

  /**
   * Load records from storage (idempotent).
   */
  async load() {
    if (this._loaded) return;
    if (this._loadPromise) return await this._loadPromise;

    const promise = (async () => {
      if (this._loaded) return;

      let data;
      try {
        data = await this._storage.load();
      } catch (err) {
        await this._maybeResetOnCorrupt();
        if (this._resetOnCorrupt) {
          this._loaded = true;
          return;
        }
        throw err;
      }
      if (!data) {
        this._loaded = true;
        return;
      }

      let parsed;
      try {
        const raw = new TextDecoder().decode(data);
        parsed = JSON.parse(raw);
      } catch {
        await this._maybeResetOnCorrupt();
        this._loaded = true;
        return;
      }

      const version = parsed?.version;
      if (
        !parsed ||
        typeof parsed !== "object" ||
        (version !== 1 && version !== 2) ||
        parsed.dimension !== this.dimension ||
        !Array.isArray(parsed.records)
      ) {
        await this._maybeResetOnCorrupt();
        this._loaded = true;
        return;
      }

      let records;
      try {
        records =
          version === 1
            ? parsed.records.map((r) => ({ id: r.id, vector: r.vector, metadata: r.metadata }))
            : parsed.records.map((r) => ({
                id: r.id,
                vector: base64ToFloat32Vector(r.vector_b64),
                metadata: r.metadata,
              }));
      } catch {
        await this._maybeResetOnCorrupt();
        this._loaded = true;
        return;
      }

      try {
        await super.upsert(records);
      } catch (err) {
        await this._maybeResetOnCorrupt();
        // Ensure we still proceed as an empty store when resetOnCorrupt is enabled.
        if (this._resetOnCorrupt) {
          this._records.clear();
          this._loaded = true;
          return;
        }
        throw err;
      }

      this._loaded = true;
    })();

    this._loadPromise = promise;
    try {
      await promise;
    } finally {
      if (this._loadPromise === promise) this._loadPromise = null;
    }
  }

  async _persist() {
    if (!this._dirty) return;
    const version = this._mutationVersion;
    const records = await super.list();
    const payload = JSON.stringify({
      version: JSON_VECTOR_STORE_VERSION,
      dimension: this.dimension,
      records: records.map((r) => ({
        id: r.id,
        vector_b64: float32VectorToBase64(r.vector),
        metadata: r.metadata,
      })),
    });
    const data = new TextEncoder().encode(payload);
    await this._storage.save(data);
    if (this._mutationVersion === version) {
      this._dirty = false;
    }
  }

  async _enqueuePersist() {
    const task = () => this._persist();
    // Ensure the queue keeps flowing even if a previous persist failed.
    const next = this._persistQueue.then(task, task);
    this._persistQueue = next;
    return next;
  }

  async upsert(records) {
    await this.load();
    await super.upsert(records);
    this._dirty = true;
    this._mutationVersion += 1;
    if (this._autoSave) await this._enqueuePersist();
  }

  /**
   * Update stored metadata without touching vectors.
   *
   * @param {{ id: string, metadata: any }[]} records
   */
  async updateMetadata(records) {
    if (!records.length) return;
    await this.load();
    await super.updateMetadata(records);
    this._dirty = true;
    this._mutationVersion += 1;
    if (this._autoSave) await this._enqueuePersist();
  }

  async delete(ids) {
    await this.load();
    await super.delete(ids);
    this._dirty = true;
    this._mutationVersion += 1;
    if (this._autoSave) await this._enqueuePersist();
  }

  async deleteWorkbook(workbookId) {
    await this.load();
    const deleted = await super.deleteWorkbook(workbookId);
    if (deleted === 0) return 0;
    this._dirty = true;
    this._mutationVersion += 1;
    if (this._autoSave) await this._enqueuePersist();
    return deleted;
  }

  async clear() {
    await this.load();
    await super.clear();
    this._dirty = true;
    this._mutationVersion += 1;
    if (this._autoSave) await this._enqueuePersist();
  }

  async get(id) {
    await this.load();
    return super.get(id);
  }

  async list(opts) {
    const signal = opts?.signal;
    throwIfAborted(signal);
    await this.load();
    throwIfAborted(signal);
    return super.list(opts);
  }

  async listContentHashes(opts) {
    const signal = opts?.signal;
    throwIfAborted(signal);
    await this.load();
    throwIfAborted(signal);
    return super.listContentHashes(opts);
  }

  async query(vector, topK, opts) {
    const signal = opts?.signal;
    throwIfAborted(signal);
    await this.load();
    throwIfAborted(signal);
    return super.query(vector, topK, opts);
  }

  async close() {
    // If a load is in-flight, wait for it before deciding whether we can return early.
    if (this._loadPromise) await this.load();
    if (!this._loaded && !this._dirty) {
      // If nothing was ever loaded/mutated we can return early. Still ensure any
      // queued persists are drained (this is a no-op in the common case).
      await this._persistQueue;
      return;
    }
    await this.load();
    if (this._dirty) await this._enqueuePersist();
    await this._persistQueue;
  }

  /**
   * Batch multiple mutations into a single persistence snapshot.
   *
   * When autoSave is enabled, `upsert()`/`delete()` normally persist after each call.
   * `batch()` temporarily suppresses those intermediate saves, then persists once at
   * the end if anything changed.
   *
   * @template T
   * @param {() => Promise<T> | T} fn
   * @returns {Promise<T>}
   */
  async batch(fn) {
    const isOutermost = this._batchDepth === 0;
    const prevAutoSave = isOutermost ? this._autoSave : null;
    if (isOutermost) this._autoSave = false;
    this._batchDepth += 1;
    /** @type {any} */
    let result;
    try {
      result = await fn();
    } finally {
      this._batchDepth -= 1;
      if (isOutermost) this._autoSave = prevAutoSave;
    }

    // Only persist for successful (non-throwing) batches, and only when autoSave was
    // enabled at the start of the batch.
    if (isOutermost && prevAutoSave && this._dirty) await this._enqueuePersist();
    return result;
  }
}
