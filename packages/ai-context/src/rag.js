import { isCellEmpty, normalizeRange, parseA1Range, rangeToA1 } from "./a1.js";
import { extractSheetSchema } from "./schema.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

/**
 * @param {string} input
 */
function hashString(input) {
  let hash = 2166136261;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

/**
 * @param {number[]} a
 * @param {number[]} b
 */
function cosineSimilarity(a, b) {
  let dot = 0;
  let na = 0;
  let nb = 0;
  const len = Math.min(a.length, b.length);
  for (let i = 0; i < len; i++) {
    dot += a[i] * b[i];
    na += a[i] * a[i];
    nb += b[i] * b[i];
  }
  if (na === 0 || nb === 0) return 0;
  return dot / (Math.sqrt(na) * Math.sqrt(nb));
}

export class HashEmbedder {
  /**
   * @param {{ dimension?: number }} [options]
   */
  constructor(options = {}) {
    this.dimension = options.dimension ?? 128;
  }

  /**
   * Deterministic, offline hash-based embeddings.
   *
   * This is a lightweight baseline used for RAG-style similarity search without
   * requiring API keys or local model setup. Retrieval quality is lower
   * than modern ML embeddings, but it's fast and "semantic-ish" enough for basic
   * token-overlap similarity.
   *
   * @param {string} text
   * @param {{ signal?: AbortSignal }} [options]
   * @returns {Promise<number[]>}
   */
  async embed(text, options = {}) {
    const signal = options.signal;
    throwIfAborted(signal);
    const vec = Array.from({ length: this.dimension }, () => 0);
    const tokens = text.toLowerCase().match(/[a-z0-9_]+/g) ?? [];
    for (const token of tokens) {
      throwIfAborted(signal);
      const h = hashString(token);
      vec[h % this.dimension] += 1;
    }
    // L2 normalize.
    let norm = 0;
    for (const v of vec) norm += v * v;
    norm = Math.sqrt(norm);
    if (norm > 0) {
      for (let i = 0; i < vec.length; i++) vec[i] /= norm;
    }
    return vec;
  }
}

export class InMemoryVectorStore {
  constructor() {
    /** @type {Map<string, { id: string, embedding: number[], metadata: any, text: string }>} */
    this.items = new Map();
  }

  /**
   * @param {{ id: string, embedding: number[], metadata: any, text: string }[]} items
   * @param {{ signal?: AbortSignal }} [options]
   */
  async add(items, options = {}) {
    const signal = options.signal;
    for (const item of items) {
      throwIfAborted(signal);
      this.items.set(item.id, item);
    }
  }

  /**
   * @param {number[]} queryEmbedding
   * @param {number} topK
   * @param {{ signal?: AbortSignal }} [options]
   */
  async search(queryEmbedding, topK, options = {}) {
    const signal = options.signal;
    /** @type {{ item: any, score: number }[]} */
    const scored = [];
    for (const item of this.items.values()) {
      throwIfAborted(signal);
      scored.push({ item, score: cosineSimilarity(queryEmbedding, item.embedding) });
    }
    throwIfAborted(signal);
    scored.sort((a, b) => b.score - a.score);
    throwIfAborted(signal);
    return scored.slice(0, topK);
  }

  /**
   * Remove items whose ids start with a given prefix. Useful for per-sheet
   * re-indexing when the number of chunks can shrink.
   * @param {string} prefix
   * @param {{ signal?: AbortSignal }} [options]
   */
  async deleteByPrefix(prefix, options = {}) {
    const signal = options.signal;
    for (const id of this.items.keys()) {
      throwIfAborted(signal);
      if (id.startsWith(prefix)) this.items.delete(id);
    }
  }

  get size() {
    return this.items.size;
  }
}

/**
 * @param {unknown[][]} values
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 */
function slice2D(values, range) {
  /** @type {unknown[][]} */
  const out = [];
  for (let r = range.startRow; r <= range.endRow; r++) {
    const row = values[r] ?? [];
    out.push(row.slice(range.startCol, range.endCol + 1));
  }
  return out;
}

/**
 * @param {unknown[][]} matrix
 * @param {{ maxRows: number }} options
 */
function matrixToTsv(matrix, options) {
  const lines = [];
  const limit = Math.min(matrix.length, options.maxRows);
  for (let r = 0; r < limit; r++) {
    const row = matrix[r];
    lines.push(row.map((v) => (isCellEmpty(v) ? "" : String(v))).join("\t"));
  }
  if (matrix.length > limit) lines.push(`… (${matrix.length - limit} more rows)`);
  return lines.join("\n");
}

/**
 * Chunk a sheet by detected regions for a simple RAG pipeline.
 *
 * @param {{ name: string, values: unknown[][] }} sheet
 * @param {{ maxChunkRows?: number, signal?: AbortSignal }} [options]
 */
export function chunkSheetByRegions(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const schema = extractSheetSchema(sheet, { signal });
  const maxChunkRows = options.maxChunkRows ?? 30;

  return schema.dataRegions.map((region, index) => {
    throwIfAborted(signal);
    const parsed = parseRangeFromSchemaRange(region.range);
    const matrix = slice2D(sheet.values, parsed);
    const text = matrixToTsv(matrix, { maxRows: maxChunkRows });
    return {
      id: `${sheet.name}-region-${index + 1}`,
      range: region.range,
      text,
      metadata: { type: "region", sheetName: sheet.name },
    };
  });
}

/**
 * @param {string} schemaRange
 */
function parseRangeFromSchemaRange(schemaRange) {
  // Schema ranges are produced by `extractSheetSchema` and are always A1 ranges.
  // Use the shared A1 parser so sheet quoting rules stay consistent (e.g.
  // `'My Sheet'!A1:B2`, escaped quotes, etc).
  return parseA1Range(schemaRange);
}

export class RagIndex {
  /**
   * @param {{ embedder?: HashEmbedder, store?: InMemoryVectorStore }} [options]
   */
  constructor(options = {}) {
    this.embedder = options.embedder ?? new HashEmbedder();
    this.store = options.store ?? new InMemoryVectorStore();
  }

  /**
   * @param {{ name: string, values: unknown[][] }} sheet
   * @param {{ signal?: AbortSignal }} [options]
   */
  async indexSheet(sheet, options = {}) {
    const signal = options.signal;
    throwIfAborted(signal);
    // `chunkSheetByRegions()` ids are deterministic (sheet name + region index),
    // but the number of regions can change over time. Clear the previous region
    // chunks for this sheet so stale chunks don't linger in the store.
    if (typeof this.store.deleteByPrefix === "function") {
      await this.store.deleteByPrefix(`${sheet.name}-region-`, { signal });
    }

    throwIfAborted(signal);
    const chunks = chunkSheetByRegions(sheet, { signal });
    const items = [];
    for (const chunk of chunks) {
      throwIfAborted(signal);
      const embedding = await this.embedder.embed(chunk.text, { signal });
      throwIfAborted(signal);
      items.push({
        id: chunk.id,
        embedding,
        metadata: { range: chunk.range, ...chunk.metadata },
        text: chunk.text,
      });
    }
    throwIfAborted(signal);
    await this.store.add(items, { signal });
  }

  /**
   * @param {string} query
   * @param {number} [topK]
   * @param {{ signal?: AbortSignal }} [options]
   */
  async search(query, topK = 5, options = {}) {
    const signal = options.signal;
    throwIfAborted(signal);
    const queryEmbedding = await this.embedder.embed(query, { signal });
    throwIfAborted(signal);
    const results = await this.store.search(queryEmbedding, topK, { signal });
    throwIfAborted(signal);
    return results.map((r) => ({
      range: r.item.metadata.range,
      score: r.score,
      preview: r.item.text.slice(0, 200),
    }));
  }
}

/**
 * Convenience for building a single chunk from a range in a sheet.
 * @param {{ name: string, values: unknown[][] }} sheet
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 * @param {{ maxRows?: number }} [options]
 */
export function rangeToChunk(sheet, range, options = {}) {
  const normalized = normalizeRange(range);
  const matrix = slice2D(sheet.values, normalized);
  const maxRows = options.maxRows ?? 30;
  return {
    id: `${sheet.name}-${rangeToA1({ ...normalized, sheetName: sheet.name })}`,
    range: rangeToA1({ ...normalized, sheetName: sheet.name }),
    text: matrixToTsv(matrix, { maxRows }),
    metadata: { type: "range", sheetName: sheet.name },
  };
}
