import { cosineSimilarity, normalizeL2, toFloat32Array } from "./vectorMath.js";
import { throwIfAborted } from "../utils/abort.js";

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

    // Keep only the best `topK` results while scanning. This avoids allocating an array
    // proportional to the number of records (which can be very large) when callers only
    // need a small set of nearest neighbors.
    //
    // We preserve the old `Array#sort` semantics by tracking scan order and using it as
    // a tie-breaker when scores are identical (stable sort behavior).
    /**
     * Normalize `topK` the same way `Array.prototype.slice(0, topK)` did in the previous
     * implementation:
     *  - `undefined` means "all"
     *  - non-integers are truncated toward 0
     */
    /** @type {number} */
    let k;
    if (topK === undefined) {
      k = Number.POSITIVE_INFINITY;
    } else {
      const n = Number(topK);
      if (!Number.isFinite(n)) {
        // NaN, -Infinity => behave like slice(0, 0) and return nothing.
        k = n === Number.POSITIVE_INFINITY ? n : 0;
      } else {
        k = Math.trunc(n);
      }
    }

    /**
     * Heap item. `_idx` is a stable tie-breaker based on scan order so results match
     * the previous full-sort implementation when scores are equal.
     * @type {Array<{ id: string, score: number, metadata: any, _idx: number }>}
     */
    const heap = [];

    /**
     * Return true if `a` should come before `b` in the min-heap (i.e. `a` is "worse").
     * Lower score is worse; for equal scores, later scan order is worse.
     */
    function isWorse(a, b) {
      if (a.score !== b.score) return a.score < b.score;
      return a._idx > b._idx;
    }

    function heapSwap(i, j) {
      const tmp = heap[i];
      heap[i] = heap[j];
      heap[j] = tmp;
    }

    function heapifyUp(i) {
      while (i > 0) {
        const parent = (i - 1) >> 1;
        if (!isWorse(heap[i], heap[parent])) break;
        heapSwap(i, parent);
        i = parent;
      }
    }

    function heapifyDown(i) {
      while (true) {
        const left = i * 2 + 1;
        if (left >= heap.length) return;
        const right = left + 1;
        let smallest = left;
        if (right < heap.length && isWorse(heap[right], heap[left])) smallest = right;
        if (!isWorse(heap[smallest], heap[i])) return;
        heapSwap(i, smallest);
        i = smallest;
      }
    }

    /** @type {number} */
    let idx = 0;
    for (const [id, rec] of this._records) {
      throwIfAborted(signal);
      if (workbookId && rec.metadata?.workbookId !== workbookId) {
        idx += 1;
        continue;
      }
      if (filter && !filter(rec.metadata, id)) {
        idx += 1;
        continue;
      }

      // Fast path: if `topK` is 0/null/negative, keep nothing.
      if (k <= 0) {
        idx += 1;
        continue;
      }

      const score = cosineSimilarity(q, rec.vector);

      if (heap.length < k) {
        heap.push({ id, score, metadata: rec.metadata, _idx: idx });
        heapifyUp(heap.length - 1);
        idx += 1;
        continue;
      }

      // Heap is full; evict the worst item if this candidate is better.
      const worst = heap[0];
      if (score > worst.score || (score === worst.score && idx < worst._idx)) {
        // Reuse the existing object to avoid per-record allocations when the corpus is large.
        worst.id = id;
        worst.score = score;
        worst.metadata = rec.metadata;
        worst._idx = idx;
        heapifyDown(0);
      }
      idx += 1;
    }

    throwIfAborted(signal);

    // Sort the retained results descending (stable tie-breaker via scan index) to
    // match the previous behavior of sorting the full scored list.
    heap.sort((a, b) => b.score - a.score || a._idx - b._idx);
    throwIfAborted(signal);

    // Strip internal bookkeeping before returning.
    return heap.map(({ id, score, metadata }) => ({ id, score, metadata }));
  }

  async close() {}
}
