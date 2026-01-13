import { normalizeL2 } from "../store/vectorMath.js";
import { throwIfAborted } from "../utils/abort.js";

/**
 * Tokenize text for hash embeddings.
 *
 * Goals:
 * - stay deterministic + fully offline
 * - split on non-word boundaries (spaces/punct), underscores, camelCase/PascalCase,
 *   and digit boundaries so e.g. `RevenueByRegion` matches `revenue by region`.
 *
 * Notes:
 * - This intentionally focuses on ASCII letter/digit behavior. Non-ASCII input is
 *   treated as a separator (rather than attempting full Unicode word breaking),
 *   but should never throw.
 *
 * @param {string} text
 * @returns {string[]}
 */
function tokenize(text) {
  const raw = String(text);

  // Insert explicit separators at boundaries we care about, then split.
  //
  // We do this before lowercasing so ASCII case transitions are detectable.
  const separated = raw
    // Treat underscores as word separators (common in identifiers/table names).
    .replace(/_/g, " ")
    // Split camelCase: `fooBar` -> `foo Bar`
    .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
    // Split PascalCase acronyms: `HTTPServer` -> `HTTP Server`
    .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
    // Split digit boundaries: `Q4Revenue` -> `Q 4 Revenue`
    .replace(/([A-Za-z])([0-9])/g, "$1 $2")
    .replace(/([0-9])([A-Za-z])/g, "$1 $2");

  return separated
    .toLowerCase()
    // Keep behavior ASCII-focused: tokens are [a-z0-9]+; everything else is a separator.
    .split(/[^a-z0-9]+/g)
    .filter(Boolean);
}

const DEFAULT_TOKEN_CACHE_SIZE = 50_000;

function shouldEnableDebugCounters() {
  // Keep `HashEmbedder` browser-safe (no hard dependency on Node globals), but
  // still make it possible to validate cache hits in `node --test` suites.
  const proc = typeof globalThis !== "undefined" ? globalThis.process : undefined;
  return Boolean(
    proc &&
      (proc.env?.NODE_ENV === "test" ||
        (Array.isArray(proc.execArgv) && proc.execArgv.includes("--test")))
  );
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
   * @param {{ dimension?: number, cacheSize?: number }} [opts]
   */
  constructor(opts) {
    this._dimension = opts?.dimension ?? 384;
    const cacheSizeOpt = opts?.cacheSize;
    const cacheSize =
      cacheSizeOpt === undefined ? DEFAULT_TOKEN_CACHE_SIZE : cacheSizeOpt;
    this._tokenCacheMaxSize =
      typeof cacheSize === "number" && Number.isFinite(cacheSize) && cacheSize > 0
        ? Math.floor(cacheSize)
        : 0;
    this._tokenCache =
      this._tokenCacheMaxSize > 0 ? new Map() : null;

    if (shouldEnableDebugCounters()) {
      Object.defineProperty(this, "_debug", {
        value: { tokenCacheHits: 0, tokenCacheMisses: 0, tokenCacheClears: 0 },
        enumerable: false,
      });
    }
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
    return `hash:v2:${this._dimension}`;
  }

  /**
   * @param {string} token
   * @returns {{ idx: number, sign: number }}
   */
  _tokenToIndexSign(token) {
    // Signed hashing reduces the positive similarity bias from collisions.
    // Use a high bit so sign isn't trivially determined by `idx` for even dimensions.
    const cache = this._tokenCache;
    if (!cache) {
      const h = fnv1a32(token);
      return { idx: h % this._dimension, sign: (h & 0x80000000) === 0 ? 1 : -1 };
    }

    const cached = cache.get(token);
    if (cached) {
      // @ts-ignore - `_debug` is a test-only internal.
      if (this._debug) this._debug.tokenCacheHits += 1;
      return cached;
    }

    // @ts-ignore - `_debug` is a test-only internal.
    if (this._debug) this._debug.tokenCacheMisses += 1;

    const h = fnv1a32(token);
    const value = { idx: h % this._dimension, sign: (h & 0x80000000) === 0 ? 1 : -1 };

    if (cache.size >= this._tokenCacheMaxSize) {
      cache.clear();
      // @ts-ignore - `_debug` is a test-only internal.
      if (this._debug) this._debug.tokenCacheClears += 1;
    }

    cache.set(token, value);
    return value;
  }

  /**
   * @param {string} text
   * @param {AbortSignal | undefined} [signal]
   */
  _embedOne(text, signal) {
    throwIfAborted(signal);
    const vec = new Float32Array(this._dimension);

    const tokens = tokenize(text);
    if (tokens.length === 0) return normalizeL2(vec);

    /** @type {Map<string, number>} */
    const termFreq = new Map();
    for (const token of tokens) {
      throwIfAborted(signal);
      termFreq.set(token, (termFreq.get(token) ?? 0) + 1);
    }

    for (const [token, tf] of termFreq) {
      throwIfAborted(signal);
      const { idx, sign } = this._tokenToIndexSign(token);
      // Light TF damping: repeated tokens matter, but sublinearly.
      const w = Math.sqrt(tf);
      vec[idx] += sign * w;
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
