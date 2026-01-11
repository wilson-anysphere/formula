/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

const SIZE_MARKER_KEY = "__pq_cache_bytes";

/**
 * @param {string} text
 */
function utf8ByteLength(text) {
  if (typeof Buffer !== "undefined") {
    return Buffer.byteLength(text, "utf8");
  }
  if (typeof TextEncoder !== "undefined") {
    return new TextEncoder().encode(text).byteLength;
  }
  // Fallback: assume 1 byte per code unit.
  return text.length;
}

/**
 * Approximate the serialized size of a cache entry.
 *
 * @param {CacheEntry} entry
 */
function estimateEntrySizeBytes(entry) {
  let binaryBytes = 0;
  let json = "";
  try {
    json = JSON.stringify(entry, (_key, value) => {
      if (value instanceof Uint8Array) {
        binaryBytes += value.byteLength;
        return { [SIZE_MARKER_KEY]: value.byteLength };
      }
      if (value instanceof ArrayBuffer) {
        binaryBytes += value.byteLength;
        return { [SIZE_MARKER_KEY]: value.byteLength };
      }
      // Node Buffers define a `toJSON()` hook that runs before replacers, so we may
      // see the `{ type: "Buffer", data: number[] }` shape here instead of the
      // original `Uint8Array`. Treat it as binary to avoid JSON bloat.
      if (
        value &&
        typeof value === "object" &&
        !Array.isArray(value) &&
        // @ts-ignore - runtime inspection
        value.type === "Buffer" &&
        // @ts-ignore - runtime inspection
        Array.isArray(value.data)
      ) {
        // @ts-ignore - runtime inspection
        const bytes = Uint8Array.from(value.data);
        binaryBytes += bytes.byteLength;
        return { [SIZE_MARKER_KEY]: bytes.byteLength };
      }
      if (typeof value === "bigint") {
        return value.toString();
      }
      return value;
    });
  } catch {
    json = "";
  }

  return utf8ByteLength(json) + binaryBytes;
}

export class MemoryCacheStore {
  /**
   * @param {{ now?: () => number }} [options]
   */
  constructor(options = {}) {
    /** @type {Map<string, CacheEntry>} */
    this.map = new Map();
    /** @type {Map<string, { lastAccessMs: number, sizeBytes: number }>} */
    this._meta = new Map();
    this.now = options.now ?? (() => Date.now());
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    const entry = this.map.get(key) ?? null;
    if (!entry) return null;

    const meta = this._meta.get(key) ?? { lastAccessMs: entry.createdAtMs ?? 0, sizeBytes: estimateEntrySizeBytes(entry) };
    meta.lastAccessMs = this.now();
    this._meta.set(key, meta);

    return entry;
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    this.map.set(key, entry);
    this._meta.set(key, { lastAccessMs: this.now(), sizeBytes: estimateEntrySizeBytes(entry) });
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    this.map.delete(key);
    this._meta.delete(key);
  }

  async clear() {
    this.map.clear();
    this._meta.clear();
  }

  /**
   * Proactively delete expired entries.
   *
   * @param {number} [nowMs]
   */
  async pruneExpired(nowMs = Date.now()) {
    for (const [key, entry] of this.map.entries()) {
      if (entry?.expiresAtMs != null && entry.expiresAtMs <= nowMs) {
        this.map.delete(key);
        this._meta.delete(key);
      }
    }
  }

  /**
   * Prune expired entries and enforce optional entry/byte quotas using LRU eviction.
   *
   * @param {{ nowMs: number, maxEntries?: number, maxBytes?: number }} options
   */
  async prune(options) {
    const maxEntries = options.maxEntries;
    const maxBytes = options.maxBytes;
    if (maxEntries == null && maxBytes == null) {
      await this.pruneExpired(options.nowMs);
      return;
    }

    /** @type {Array<{ key: string, sizeBytes: number, lastAccessMs: number }>} */
    const items = [];
    let totalBytes = 0;

    for (const [key, entry] of this.map.entries()) {
      if (entry?.expiresAtMs != null && entry.expiresAtMs <= options.nowMs) {
        this.map.delete(key);
        this._meta.delete(key);
        continue;
      }
      let meta = this._meta.get(key);
      if (!meta || typeof meta.sizeBytes !== "number" || typeof meta.lastAccessMs !== "number") {
        meta = {
          lastAccessMs: typeof entry.createdAtMs === "number" ? entry.createdAtMs : 0,
          sizeBytes: estimateEntrySizeBytes(entry),
        };
        this._meta.set(key, meta);
      }
      totalBytes += meta.sizeBytes;
      items.push({ key, sizeBytes: meta.sizeBytes, lastAccessMs: meta.lastAccessMs });
    }

    let totalEntries = items.length;

    items.sort((a, b) => {
      if (a.lastAccessMs !== b.lastAccessMs) return a.lastAccessMs - b.lastAccessMs;
      return a.key.localeCompare(b.key);
    });

    let idx = 0;
    while (
      (maxEntries != null && totalEntries > maxEntries) ||
      (maxBytes != null && totalBytes > maxBytes)
    ) {
      const victim = items[idx++];
      if (!victim) break;
      if (!this.map.has(victim.key)) continue;
      this.map.delete(victim.key);
      this._meta.delete(victim.key);
      totalEntries -= 1;
      totalBytes -= victim.sizeBytes;
    }
  }
}
