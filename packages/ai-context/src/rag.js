import { normalizeRange, parseA1Range, rangeToA1 } from "./a1.js";
import { extractSheetSchema } from "./schema.js";
import { throwIfAborted } from "./abort.js";
import { deleteLegacySheetRegionChunks, sheetChunkIdPrefix } from "./ragIds.js";
import { valuesRangeToTsv } from "./tsv.js";

/**
 * @param {any} sheet
 */
function normalizeSheetOrigin(sheet) {
  if (!sheet || typeof sheet !== "object" || !sheet.origin || typeof sheet.origin !== "object") {
    return { row: 0, col: 0 };
  }
  const row = Number.isInteger(sheet.origin.row) && sheet.origin.row >= 0 ? sheet.origin.row : 0;
  const col = Number.isInteger(sheet.origin.col) && sheet.origin.col >= 0 ? sheet.origin.col : 0;
  return { row, col };
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

/**
 * Tokenize text for hash embeddings.
 *
 * Keep this broadly aligned with `packages/ai-rag`'s HashEmbedder tokenizer:
 * - split on punctuation/whitespace
 * - treat underscores as separators (common in spreadsheet headers / identifiers)
 * - split camelCase/PascalCase + digit boundaries so `RevenueByRegion2024` matches
 *   `revenue by region 2024`
 *
 * @param {string} text
 */
function tokenize(text) {
  const raw = String(text);
  const separated = raw
    .replace(/_/g, " ")
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    .replace(/([A-Za-z])([0-9])/g, "$1 $2")
    .replace(/([0-9])([A-Za-z])/g, "$1 $2");

  return separated
    .toLowerCase()
    .split(/[^a-z0-9]+/g)
    .filter(Boolean);
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
    const tokens = tokenize(text);

    /** @type {Map<string, number>} */
    const termFreq = new Map();
    for (const token of tokens) {
      throwIfAborted(signal);
      termFreq.set(token, (termFreq.get(token) ?? 0) + 1);
    }

    for (const [token, tf] of termFreq) {
      throwIfAborted(signal);
      const h = hashString(token);
      const idx = h % this.dimension;
      const sign = (h & 0x80000000) === 0 ? 1 : -1;
      // Light TF damping: repeated tokens matter, but sublinearly.
      const w = Math.sqrt(tf);
      vec[idx] += sign * w;
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
    const count = this.items.size;
    // Mirror `Array.prototype.slice`'s ToInteger behavior for common cases.
    const k = Number.isFinite(topK) ? Math.trunc(topK) : count;
 
    /**
     * Score-descending, id-ascending ordering.
     *
     * @param {{ item: any, score: number }} a
     * @param {{ item: any, score: number }} b
     */
    function compareScored(a, b) {
      // Sort by score descending, but ensure deterministic ordering when multiple
      // items share the same similarity score (e.g. identical embeddings).
      if (a.score > b.score) return -1;
      if (a.score < b.score) return 1;
      // Tie-break by id so prompt context ordering is stable.
      if (a.item.id < b.item.id) return -1;
      if (a.item.id > b.item.id) return 1;
      return 0;
    }
 
    // If we're retrieving a strict subset, keep an in-order top-K list rather than
    // sorting all items. This makes search scale better when row-window chunking
    // increases the number of stored chunks.
    const shouldUsePartial = k > 0 && k < count;
    if (shouldUsePartial) {
      /** @type {{ item: any, score: number }[]} */
      const best = [];
      for (const item of this.items.values()) {
        throwIfAborted(signal);
        const entry = { item, score: cosineSimilarity(queryEmbedding, item.embedding) };
        // Find insertion position in the current best list.
        let i = 0;
        while (i < best.length && compareScored(entry, best[i]) > 0) i += 1;
        if (i >= k) continue;
        best.splice(i, 0, entry);
        if (best.length > k) best.pop();
      }
      throwIfAborted(signal);
      return best;
    }
 
    /** @type {{ item: any, score: number }[]} */
    const scored = [];
    for (const item of this.items.values()) {
      throwIfAborted(signal);
      scored.push({ item, score: cosineSimilarity(queryEmbedding, item.embedding) });
    }
    throwIfAborted(signal);
    scored.sort(compareScored);
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
 */
function getMatrixBounds(values) {
  const rowCount = Array.isArray(values) ? values.length : 0;
  let colCount = 0;
  if (!Array.isArray(values) || rowCount === 0) return { rowCount: 0, colCount: 0 };

  // Prefer a sparse-friendly scan for very large arrays. `for...of` visits holes,
  // which can be unexpectedly expensive when callers pass Excel-scale sparse
  // matrices (e.g. `new Array(1_048_576)` with only a few populated rows).
  const LARGE_ROW_THRESHOLD = 10_000;
  if (rowCount > LARGE_ROW_THRESHOLD) {
    for (const key in values) {
      const row = values[key];
      colCount = Math.max(colCount, row?.length ?? 0);
    }
  } else {
    for (const row of values) {
      colCount = Math.max(colCount, row?.length ?? 0);
    }
  }

  return { rowCount, colCount };
}

/**
 * Clamp a rect range (0-based, inclusive) to the bounds of a matrix.
 *
 * Returns null when the range does not intersect the provided matrix at all.
 *
 * @param {{ rowCount: number, colCount: number }} bounds
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 */
function clampRangeToMatrixBounds(bounds, range) {
  const rowCount = bounds.rowCount;
  const colCount = bounds.colCount;
  if (rowCount === 0 || colCount === 0) return null;

  if (range.endRow < 0 || range.endCol < 0) return null;
  if (range.startRow >= rowCount || range.startCol >= colCount) return null;

  const startRow = Math.max(0, Math.min(range.startRow, rowCount - 1));
  const endRow = Math.max(0, Math.min(range.endRow, rowCount - 1));
  const startCol = Math.max(0, Math.min(range.startCol, colCount - 1));
  const endCol = Math.max(0, Math.min(range.endCol, colCount - 1));

  if (endRow < startRow || endCol < startCol) return null;
  return { startRow, startCol, endRow, endCol };
}

/**
 * @param {{ startRow: number, endRow: number }} rect
 * @param {{ windowSize: number, overlap: number, maxChunks: number }} options
 */
function splitRectByRowWindows(rect, options) {
  const rowCount = rect.endRow - rect.startRow + 1;
  const windowSize = Math.max(1, Math.floor(options.windowSize));
  const overlap = Math.max(0, Math.min(Math.floor(options.overlap), windowSize - 1));
  const maxChunks = Math.max(1, Math.floor(options.maxChunks));

  if (maxChunks === 1) {
    return [{ startRow: rect.startRow, endRow: rect.endRow, index: 0 }];
  }

  if (rowCount <= windowSize) {
    return [{ startRow: rect.startRow, endRow: rect.endRow, index: 0 }];
  }

  // Default stride tries to preserve a small overlap between windows.
  let step = Math.max(1, windowSize - overlap);

  const idealChunks = Math.ceil((rowCount - windowSize) / step) + 1;
  if (idealChunks > maxChunks) {
    // If we'd generate too many chunks, increase the stride to fit within the cap.
    // This may introduce gaps, but keeps indexing bounded for very tall tables.
    step = Math.max(1, Math.ceil((rowCount - windowSize) / (maxChunks - 1)));
  }

  /** @type {{ startRow: number, endRow: number, index: number }[]} */
  const windows = [];
  for (let i = 0; i < maxChunks; i++) {
    const startRow = rect.startRow + i * step;
    if (startRow > rect.endRow) break;
    const endRow = Math.min(rect.endRow, startRow + windowSize - 1);
    windows.push({ startRow, endRow, index: i });
    if (endRow === rect.endRow) break;
  }

  // Ensure we always include a trailing window ending at `rect.endRow` so bottom-of-table
  // queries have something to retrieve, even when `maxChunks` forces a large stride.
  const last = windows[windows.length - 1];
  if (last && last.endRow < rect.endRow) {
    const startRow = Math.max(rect.startRow, rect.endRow - windowSize + 1);
    const endRow = rect.endRow;
    if (windows.length < maxChunks) {
      windows.push({ startRow, endRow, index: windows.length });
    } else {
      windows[windows.length - 1] = { startRow, endRow, index: last.index };
    }
  }

  return windows;
}

/**
 * @typedef {{
 *   id: string,
 *   range: string,
 *   text: string,
 *   metadata: any
 * }} SheetChunk
 */

/**
 * Chunk a sheet by detected regions for a simple RAG pipeline.
 *
 * @param {{ name: string, values: unknown[][], origin?: { row: number, col: number } }} sheet
 * @param {{
 *   maxChunkRows?: number,
 *   /**
 *    * Alias for `splitByRowWindows`.
 *    *\/
 *   splitRegions?: boolean,
 *   splitByRowWindows?: boolean,
 *   /**
 *    * Alias for `rowOverlap`.
 *    *\/
 *   chunkRowOverlap?: number,
 *   rowOverlap?: number,
 *   maxChunksPerRegion?: number,
 *   signal?: AbortSignal
 * }} [options]
 * @returns {SheetChunk[]}
 */
export function chunkSheetByRegions(sheet, options = {}) {
  return chunkSheetByRegionsWithSchema(sheet, options).chunks;
}

/**
 * Chunk a sheet by detected regions for a simple RAG pipeline, reusing a single
 * schema extraction pass.
 *
 * @param {{ name: string, values: unknown[][], origin?: { row: number, col: number } }} sheet
 * @param {{
 *   maxChunkRows?: number,
 *   /**
 *    * Alias for `splitByRowWindows`.
 *    *\/
 *   splitRegions?: boolean,
 *   splitByRowWindows?: boolean,
 *   /**
 *    * Alias for `rowOverlap`.
 *    *\/
 *   chunkRowOverlap?: number,
 *   rowOverlap?: number,
 *   maxChunksPerRegion?: number,
 *   signal?: AbortSignal
 * }} [options]
 * @returns {{ schema: ReturnType<typeof extractSheetSchema>, chunks: SheetChunk[] }}
 */
export function chunkSheetByRegionsWithSchema(sheet, options = {}) {
  const signal = options.signal;
  throwIfAborted(signal);
  const schema = extractSheetSchema(sheet, { signal });
  const maxChunkRows = options.maxChunkRows ?? 30;
  const splitByRowWindows = options.splitByRowWindows ?? options.splitRegions ?? false;
  const rowOverlap = options.rowOverlap ?? options.chunkRowOverlap ?? 3;
  const maxChunksPerRegion = options.maxChunksPerRegion ?? 50;
  const origin = normalizeSheetOrigin(sheet);
  const matrixBounds = getMatrixBounds(sheet.values);
  const chunkIdPrefix = sheetChunkIdPrefix(sheet.name);

  /** @type {SheetChunk[]} */
  const chunks = [];

  for (let regionIndex = 0; regionIndex < schema.dataRegions.length; regionIndex++) {
    throwIfAborted(signal);
    const region = schema.dataRegions[regionIndex];
    const parsedAbs = parseRangeFromSchemaRange(region.range);

    const regionRectAbs = normalizeRange(parsedAbs);
    const shouldPrefixHeader = splitByRowWindows && region.hasHeader;
    let headerLine = "";
    if (shouldPrefixHeader) {
      const headerRangeRaw = {
        startRow: regionRectAbs.startRow - origin.row,
        endRow: regionRectAbs.startRow - origin.row,
        startCol: regionRectAbs.startCol - origin.col,
        endCol: regionRectAbs.endCol - origin.col,
      };
      const headerRange = clampRangeToMatrixBounds(matrixBounds, headerRangeRaw);
      headerLine = headerRange ? valuesRangeToTsv(sheet.values, headerRange, { maxRows: 1, signal }) : "";
    }
    const windows = splitByRowWindows
      ? splitRectByRowWindows(
          { startRow: regionRectAbs.startRow, endRow: regionRectAbs.endRow },
          { windowSize: maxChunkRows, overlap: rowOverlap, maxChunks: maxChunksPerRegion },
        )
      : [{ startRow: regionRectAbs.startRow, endRow: regionRectAbs.endRow, index: 0 }];

    const baseId = `${chunkIdPrefix}${regionIndex + 1}`;
    const originSuffix = origin.row !== 0 || origin.col !== 0 ? `-o${origin.row}x${origin.col}` : "";

    for (const window of windows) {
      throwIfAborted(signal);
      const windowRectAbs = {
        startRow: window.startRow,
        endRow: window.endRow,
        startCol: regionRectAbs.startCol,
        endCol: regionRectAbs.endCol,
      };

      const windowRangeA1 = rangeToA1({ ...windowRectAbs, sheetName: sheet.name });

      const localRangeRaw = {
        startRow: windowRectAbs.startRow - origin.row,
        endRow: windowRectAbs.endRow - origin.row,
        startCol: windowRectAbs.startCol - origin.col,
        endCol: windowRectAbs.endCol - origin.col,
      };

      const localRange = clampRangeToMatrixBounds(matrixBounds, localRangeRaw);
      const windowText = localRange ? valuesRangeToTsv(sheet.values, localRange, { maxRows: maxChunkRows, signal }) : "";
      const text =
        headerLine && window.startRow > regionRectAbs.startRow
          ? windowText
            ? `${headerLine}\n${windowText}`
            : headerLine
          : windowText;

      chunks.push({
        id: splitByRowWindows ? `${baseId}${originSuffix}-rows-${window.startRow}` : `${baseId}${originSuffix}`,
        range: windowRangeA1,
        text,
        metadata: { type: "region", sheetName: sheet.name, regionRange: region.range },
      });
    }
  }

  return { schema, chunks };
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
   * @param {{ name: string, values: unknown[][], origin?: { row: number, col: number } }} sheet
   * @param {{
   *   maxChunkRows?: number,
   *   /**
   *    * Alias for `splitByRowWindows`.
   *    *\/
   *   splitRegions?: boolean,
   *   splitByRowWindows?: boolean,
   *   /**
   *    * Alias for `rowOverlap`.
   *    *\/
   *   chunkRowOverlap?: number,
   *   rowOverlap?: number,
   *   maxChunksPerRegion?: number,
   *   signal?: AbortSignal
   * }} [options]
   * @returns {Promise<{ schema: ReturnType<typeof extractSheetSchema>, chunkCount: number }>}
   */
  async indexSheet(sheet, options = {}) {
    const signal = options.signal;
    throwIfAborted(signal);
    // `chunkSheetByRegions()` ids are deterministic (sheet name + region index, with an
    // optional row-window start row suffix when splitting is enabled),
    // but the number of regions can change over time. Clear the previous region
    // chunks for this sheet so stale chunks don't linger in the store.
    if (typeof this.store.deleteByPrefix === "function") {
      await this.store.deleteByPrefix(sheetChunkIdPrefix(sheet.name), { signal });
    }
    // Best-effort cleanup for stale chunks created by older versions of ai-context.
    // Use a precise match to avoid resurrecting the legacy prefix collision bug.
    deleteLegacySheetRegionChunks(this.store, sheet.name, { signal });

    throwIfAborted(signal);
    const splitByRowWindows = options.splitByRowWindows ?? options.splitRegions;
    const rowOverlap = options.rowOverlap ?? options.chunkRowOverlap;
    const { schema, chunks } = chunkSheetByRegionsWithSchema(sheet, {
      signal,
      maxChunkRows: options.maxChunkRows,
      splitByRowWindows,
      rowOverlap,
      maxChunksPerRegion: options.maxChunksPerRegion,
    });
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

    return { schema, chunkCount: chunks.length };
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
 * @param {{ name: string, values: unknown[][], origin?: { row: number, col: number } }} sheet
 * @param {{ startRow: number, startCol: number, endRow: number, endCol: number }} range
 * @param {{ maxRows?: number }} [options]
 */
export function rangeToChunk(sheet, range, options = {}) {
  const normalized = normalizeRange(range);
  const origin = normalizeSheetOrigin(sheet);
  const matrixBounds = getMatrixBounds(sheet.values);
  const localRangeRaw = {
    startRow: normalized.startRow - origin.row,
    endRow: normalized.endRow - origin.row,
    startCol: normalized.startCol - origin.col,
    endCol: normalized.endCol - origin.col,
  };
  const localRange = clampRangeToMatrixBounds(matrixBounds, localRangeRaw);
  const maxRows = options.maxRows ?? 30;
  const text = localRange ? valuesRangeToTsv(sheet.values, localRange, { maxRows }) : "";
  return {
    id: `${sheet.name}-${rangeToA1({ ...normalized, sheetName: sheet.name })}`,
    range: rangeToA1({ ...normalized, sheetName: sheet.name }),
    text,
    metadata: { type: "range", sheetName: sheet.name },
  };
}
