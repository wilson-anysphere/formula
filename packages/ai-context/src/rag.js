import { isCellEmpty, normalizeRange, rangeToA1 } from "./a1.js";
import { extractSheetSchema } from "./schema.js";

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
   * requiring API keys or local model configuration. Retrieval quality is lower
   * than modern ML embeddings, but it's fast and "semantic-ish" enough for basic
   * token-overlap similarity.
   *
   * @param {string} text
   * @returns {Promise<number[]>}
   */
  async embed(text) {
    const vec = Array.from({ length: this.dimension }, () => 0);
    const tokens = text.toLowerCase().match(/[a-z0-9_]+/g) ?? [];
    for (const token of tokens) {
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
   */
  async add(items) {
    for (const item of items) {
      this.items.set(item.id, item);
    }
  }

  /**
   * @param {number[]} queryEmbedding
   * @param {number} topK
   */
  async search(queryEmbedding, topK) {
    const scored = Array.from(this.items.values()).map((item) => ({
      item,
      score: cosineSimilarity(queryEmbedding, item.embedding),
    }));
    scored.sort((a, b) => b.score - a.score);
    return scored.slice(0, topK);
  }

  /**
   * Remove items whose ids start with a given prefix. Useful for per-sheet
   * re-indexing when the number of chunks can shrink.
   * @param {string} prefix
   */
  async deleteByPrefix(prefix) {
    for (const id of this.items.keys()) {
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
 * @param {{ maxChunkRows?: number }} [options]
 */
export function chunkSheetByRegions(sheet, options = {}) {
  const schema = extractSheetSchema(sheet);
  const maxChunkRows = options.maxChunkRows ?? 30;

  return schema.dataRegions.map((region, index) => {
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
  // schemaRange is always in `Sheet!A1:B2` form from extractSheetSchema.
  const [sheetNameAndMaybe, a1] = schemaRange.includes("!") ? schemaRange.split("!") : ["", schemaRange];
  const sheetName = sheetNameAndMaybe || undefined;
  const match = /^(?<start>[A-Z]+\d+)(?::(?<end>[A-Z]+\d+))?$/.exec(a1);
  if (!match || !match.groups) throw new Error(`Invalid schema range: ${schemaRange}`);

  const start = match.groups.start;
  const end = match.groups.end ?? start;

  const startRef = cellFromA1(start);
  const endRef = cellFromA1(end);
  return normalizeRange({
    sheetName,
    startRow: startRef.row,
    startCol: startRef.col,
    endRow: endRef.row,
    endCol: endRef.col,
  });
}

/**
 * @param {string} a1Cell
 */
function cellFromA1(a1Cell) {
  const m = /^([A-Z]+)(\d+)$/.exec(a1Cell);
  if (!m) throw new Error(`Invalid A1 cell: ${a1Cell}`);
  const [, letters, digits] = m;
  let col = 0;
  for (const char of letters) col = col * 26 + (char.charCodeAt(0) - 64);
  col -= 1;
  const row = Number(digits) - 1;
  return { row, col };
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
   */
  async indexSheet(sheet) {
    // `chunkSheetByRegions()` ids are deterministic (sheet name + region index),
    // but the number of regions can change over time. Clear the previous region
    // chunks for this sheet so stale chunks don't linger in the store.
    if (typeof this.store.deleteByPrefix === "function") {
      await this.store.deleteByPrefix(`${sheet.name}-region-`);
    }

    const chunks = chunkSheetByRegions(sheet);
    const items = [];
    for (const chunk of chunks) {
      const embedding = await this.embedder.embed(chunk.text);
      items.push({
        id: chunk.id,
        embedding,
        metadata: { range: chunk.range, ...chunk.metadata },
        text: chunk.text,
      });
    }
    await this.store.add(items);
  }

  /**
   * @param {string} query
   * @param {number} [topK]
   */
  async search(query, topK = 5) {
    const queryEmbedding = await this.embedder.embed(query);
    const results = await this.store.search(queryEmbedding, topK);
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
