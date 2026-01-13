import { normalizeL2 } from "../store/vectorMath.js";

function createAbortError(message = "Aborted") {
  const err = new Error(message);
  err.name = "AbortError";
  return err;
}

function throwIfAborted(signal) {
  if (signal?.aborted) throw createAbortError();
}

function fnv1a32(str) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < str.length; i += 1) {
    hash ^= str.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  // Force unsigned.
  return hash >>> 0;
}

/**
 * A deterministic, offline embedder used by Formula for workbook RAG.
 *
 * This is the supported default embedder until a Cursor-managed embedding
 * service exists.
 *
 * Formula intentionally uses a hash-based embedder so workbook indexing and
 * retrieval work without:
 * - user-supplied API keys
 * - local model downloads / setup
 * - sending workbook content to a third-party embedding provider
 *
 * This is not a true ML embedding model, so retrieval quality is meaningfully
 * lower than modern neural embeddings. However, it is fast, private, and
 * "semantic-ish" enough for basic similarity search over spreadsheet text.
 *
 * It uses a hashing trick on tokenized words, then L2-normalizes the result so
 * cosine similarity approximates shared-token overlap.
 */
export class HashEmbedder {
  /**
   * @param {{ dimension?: number }} [opts]
   */
  constructor(opts) {
    this._dimension = opts?.dimension ?? 384;
  }

  get dimension() {
    return this._dimension;
  }

  get name() {
    // Embedder identity string used in persisted metadata and cache keys.
    //
    // Include an explicit algorithm version so changes to the hashing/tokenization
    // logic can safely force a re-embed of existing vector stores (by changing
    // this string, index cache keys will change).
    return `hash:v1:${this._dimension}`;
  }

  /**
   * @param {string} text
   * @param {AbortSignal | undefined} [signal]
   */
  _embedOne(text, signal) {
    throwIfAborted(signal);
    const vec = new Float32Array(this._dimension);
    const tokens = String(text)
      .toLowerCase()
      .split(/[^a-z0-9_]+/g)
      .filter(Boolean);
    for (const token of tokens) {
      throwIfAborted(signal);
      const h = fnv1a32(token);
      const idx = h % this._dimension;
      // Simple TF weighting, lightly damped.
      vec[idx] += 1;
    }
    return normalizeL2(vec);
  }

  /**
   * @param {string[]} texts
   * @param {{ signal?: AbortSignal }} [options]
   * @returns {Promise<Float32Array[]>}
   */
  async embedTexts(texts, options = {}) {
    const signal = options.signal;
    throwIfAborted(signal);
    return texts.map((t) => this._embedOne(t, signal));
  }
}
