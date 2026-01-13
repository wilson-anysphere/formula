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
   * Update stored metadata without touching vectors.
   *
   * @param {{ id: string, metadata: any }[]} records
   */
  async updateMetadata(records) {
    for (const r of records) {
      const existing = this._records.get(r.id);
      if (!existing) continue;
      this._records.set(r.id, { vector: existing.vector, metadata: r.metadata });
    }
  }

  /**
   * @param {string[]} ids
   */
  async delete(ids) {
    for (const id of ids) this._records.delete(id);
  }

  /**
   * Delete all records associated with a workbook.
   *
   * @param {string} workbookId
   * @returns {Promise<number>} number of deleted records
   */
  async deleteWorkbook(workbookId) {
    let deleted = 0;
    for (const [id, rec] of this._records) {
      if (rec.metadata?.workbookId !== workbookId) continue;
      this._records.delete(id);
      deleted += 1;
    }
    return deleted;
  }

  /**
   * Remove all records from the store.
   */
  async clear() {
    this._records.clear();
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
   * Return `{ id, contentHash, metadataHash }` for records. This avoids returning full
   * metadata objects (e.g. `metadata.text`) when callers only need incremental change keys.
   *
   * @param {{ workbookId?: string, signal?: AbortSignal }} [opts]
   * @returns {Promise<Array<{ id: string, contentHash: string | null, metadataHash: string | null }>>}
   */
  async listContentHashes(opts) {
    const signal = opts?.signal;
    const workbookId = opts?.workbookId;
    throwIfAborted(signal);
    /** @type {Array<{ id: string, contentHash: string | null, metadataHash: string | null }>} */
    const out = [];
    for (const [id, rec] of this._records) {
      throwIfAborted(signal);
      if (workbookId && rec.metadata?.workbookId !== workbookId) continue;
      out.push({
        id,
        contentHash: typeof rec.metadata?.contentHash === "string" ? rec.metadata.contentHash : null,
        metadataHash: typeof rec.metadata?.metadataHash === "string" ? rec.metadata.metadataHash : null,
      });
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

    if (!Number.isFinite(topK)) {
      throw new Error(`Invalid topK: expected a finite number, got ${String(topK)}`);
    }
    // We deterministically floor floats (e.g. 1.9 -> 1) so vector stores behave
    // consistently even if callers compute `topK` dynamically.
    const k = Math.floor(topK);
    if (k <= 0) return [];

    const qVec = toFloat32Array(vector);
    if (qVec.length !== this._dimension) {
      throw new Error(
        `InMemoryVectorStore.query() vector dimension mismatch: expected ${this._dimension}, got ${qVec.length}`
      );
    }
    const q = normalizeL2(qVec);

    // Keep only the best `k` results while scanning. This avoids allocating an array
    // proportional to the number of records (which can be very large) when callers only
    // need a small set of nearest neighbors.
    //
    // When cosine scores tie, break ties deterministically by id (ascending) so result
    // ordering is stable across store implementations.

    /**
     * Heap items (min-heap of the "worse" result so we can keep only the topK).
     * @type {Array<{ id: string, score: number, metadata: any }>}
     */
    const heap = [];

    /**
     * Return true if `a` should come before `b` in the min-heap (i.e. `a` is "worse").
     * Lower score is worse; for equal scores, larger id is worse.
     */
    function isWorse(a, b) {
      if (a.score !== b.score) return a.score < b.score;
      return a.id > b.id;
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

    for (const [id, rec] of this._records) {
      throwIfAborted(signal);
      if (workbookId && rec.metadata?.workbookId !== workbookId) continue;
      if (filter && !filter(rec.metadata, id)) continue;

      const score = cosineSimilarity(q, rec.vector);

      if (heap.length < k) {
        heap.push({ id, score, metadata: rec.metadata });
        heapifyUp(heap.length - 1);
        continue;
      }

      // Heap is full; evict the worst item if this candidate is better.
      const worst = heap[0];
      if (score > worst.score || (score === worst.score && id < worst.id)) {
        // Reuse the existing object to avoid per-record allocations when the corpus is large.
        worst.id = id;
        worst.score = score;
        worst.metadata = rec.metadata;
        heapifyDown(0);
      }
    }

    throwIfAborted(signal);
    heap.sort((a, b) => {
      const scoreCmp = b.score - a.score;
      if (scoreCmp !== 0) return scoreCmp;
      if (a.id < b.id) return -1;
      if (a.id > b.id) return 1;
      return 0;
    });
    throwIfAborted(signal);
    return heap.map(({ id, score, metadata }) => ({ id, score, metadata }));
  }

  /**
   * Interface parity with persistent stores. In-memory stores don't need to batch,
   * but exposing this lets callers opportunistically group mutations.
   *
   * @template T
   * @param {() => Promise<T> | T} fn
   * @returns {Promise<T>}
   */
  async batch(fn) {
    return await fn();
  }

  async close() {}
}
