import { cosineSimilarity, normalizeL2, toFloat32Array } from "./vectorMath.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * @typedef {Object} VectorRecord
 * @property {string} id
 * @property {ArrayLike<number>} vector
 * @property {any} metadata
 */

/**
 * @typedef {Object} VectorSearchResult
 * @property {string} id
 * @property {number} score
 * @property {any} metadata
 */

export class InMemoryVectorStore {
  /**
   * @param {{ dimension: number }} opts
   */
  constructor(opts) {
    if (!opts || !Number.isFinite(opts.dimension) || opts.dimension <= 0) {
      throw new Error("InMemoryVectorStore requires a positive dimension");
    }
    this._dimension = opts.dimension;
    /** @type {Map<string, { vector: Float32Array, metadata: any }>} */
    this._records = new Map();
  }

  get dimension() {
    return this._dimension;
  }

  /**
   * @param {VectorRecord[]} records
   */
  async upsert(records) {
    for (const r of records) {
      const vec = toFloat32Array(r.vector);
      if (vec.length !== this._dimension) {
        throw new Error(
          `Vector dimension mismatch for id=${r.id}: expected ${this._dimension}, got ${vec.length}`
        );
      }
      // Normalize so cosineSimilarity is fast and stable across stores.
      const norm = normalizeL2(vec);
      this._records.set(r.id, { vector: norm, metadata: r.metadata });
    }
  }

  /**
   * @param {string[]} ids
   */
  async delete(ids) {
    for (const id of ids) this._records.delete(id);
  }

  /**
   * @param {string} id
   */
  async get(id) {
    const rec = this._records.get(id);
    if (!rec) return null;
    return { id, vector: rec.vector, metadata: rec.metadata };
  }

  /**
   * @param {{
   *   filter?: (metadata: any, id: string) => boolean,
   *   workbookId?: string,
   *   includeVector?: boolean
   * }} [opts]
   */
  async list(opts) {
    const signal = opts?.signal;
    const filter = opts?.filter;
    const workbookId = opts?.workbookId;
    const includeVector = opts?.includeVector ?? true;
    const out = [];
    for (const [id, rec] of this._records) {
      throwIfAborted(signal);
      if (workbookId && rec.metadata?.workbookId !== workbookId) continue;
      if (filter && !filter(rec.metadata, id)) continue;
      out.push({ id, vector: includeVector ? rec.vector : undefined, metadata: rec.metadata });
    }
    throwIfAborted(signal);
    return out;
  }

  /**
   * @param {ArrayLike<number>} vector
   * @param {number} topK
   * @param {{ filter?: (metadata: any, id: string) => boolean, workbookId?: string }} [opts]
   * @returns {Promise<VectorSearchResult[]>}
   */
  async query(vector, topK, opts) {
    const signal = opts?.signal;
    const filter = opts?.filter;
    const workbookId = opts?.workbookId;
    const q = normalizeL2(vector);
    /** @type {VectorSearchResult[]} */
    const scored = [];
    for (const [id, rec] of this._records) {
      throwIfAborted(signal);
      if (workbookId && rec.metadata?.workbookId !== workbookId) continue;
      if (filter && !filter(rec.metadata, id)) continue;
      const score = cosineSimilarity(q, rec.vector);
      scored.push({ id, score, metadata: rec.metadata });
    }
    throwIfAborted(signal);
    scored.sort((a, b) => b.score - a.score);
    throwIfAborted(signal);
    return scored.slice(0, topK);
  }

  async close() {}
}
