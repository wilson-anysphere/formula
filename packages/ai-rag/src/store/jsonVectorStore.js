import { InMemoryBinaryStorage } from "./binaryStorage.js";
import { InMemoryVectorStore } from "./inMemoryVectorStore.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
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

    if (
      !parsed ||
      typeof parsed !== "object" ||
      parsed.version !== 1 ||
      parsed.dimension !== this.dimension ||
      !Array.isArray(parsed.records)
    ) {
      return;
    }

    await super.upsert(
      parsed.records.map((r) => ({ id: r.id, vector: r.vector, metadata: r.metadata }))
    );
  }

  async _persist() {
    const records = await super.list();
    const payload = JSON.stringify({
      version: 1,
      dimension: this.dimension,
      records: records.map((r) => ({
        id: r.id,
        vector: Array.from(r.vector),
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
