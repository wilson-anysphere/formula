import { cosineSimilarity, normalizeL2, toFloat32Array } from "./vectorMath.js";

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
   * @param {{ filter?: (metadata: any, id: string) => boolean }} [opts]
   */
  async list(opts) {
    const filter = opts?.filter;
    const out = [];
    for (const [id, rec] of this._records) {
      if (filter && !filter(rec.metadata, id)) continue;
      out.push({ id, vector: rec.vector, metadata: rec.metadata });
    }
    return out;
  }

  /**
   * @param {ArrayLike<number>} vector
   * @param {number} topK
   * @param {{ filter?: (metadata: any, id: string) => boolean }} [opts]
   * @returns {Promise<VectorSearchResult[]>}
   */
  async query(vector, topK, opts) {
    const filter = opts?.filter;
    const q = normalizeL2(vector);
    /** @type {VectorSearchResult[]} */
    const scored = [];
    for (const [id, rec] of this._records) {
      if (filter && !filter(rec.metadata, id)) continue;
      const score = cosineSimilarity(q, rec.vector);
      scored.push({ id, score, metadata: rec.metadata });
    }
    scored.sort((a, b) => b.score - a.score);
    return scored.slice(0, topK);
  }

  async close() {}
}
