import { fromBase64, InMemoryBinaryStorage, toBase64 } from "./binaryStorage.js";
import { InMemoryVectorStore } from "./inMemoryVectorStore.js";

const JSON_VECTOR_STORE_VERSION = 2;

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

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
   * @param {{ storage?: any, dimension: number, autoSave?: boolean }} opts
   */
  constructor(opts) {
    super({ dimension: opts.dimension });
    this._storage = opts.storage ?? new InMemoryBinaryStorage();
    this._autoSave = opts.autoSave ?? true;
    this._loaded = false;
    this._dirty = false;
  }

  /**
   * Load records from storage (idempotent).
   */
  async load() {
    if (this._loaded) return;
    this._loaded = true;

    const data = await this._storage.load();
    if (!data) return;

    let parsed;
    try {
      const raw = new TextDecoder().decode(data);
      parsed = JSON.parse(raw);
    } catch {
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
      return;
    }

    await super.upsert(records);
  }

  async _persist() {
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
    this._dirty = false;
  }

  async upsert(records) {
    await this.load();
    await super.upsert(records);
    this._dirty = true;
    if (this._autoSave) await this._persist();
  }

  async delete(ids) {
    await this.load();
    await super.delete(ids);
    this._dirty = true;
    if (this._autoSave) await this._persist();
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

  async query(vector, topK, opts) {
    const signal = opts?.signal;
    throwIfAborted(signal);
    await this.load();
    throwIfAborted(signal);
    return super.query(vector, topK, opts);
  }

  async close() {
    if (!this._loaded && !this._dirty) return;
    await this.load();
    if (this._dirty) await this._persist();
  }
}
