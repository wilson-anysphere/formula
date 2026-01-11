/**
 * Connector-agnostic caching primitives.
 *
 * The cache manager is intentionally simple:
 * - deterministic keys (caller provided)
 * - optional TTL
 * - pluggable storage backends (e.g. `MemoryCacheStore`, `FileSystemCacheStore`,
 *   and `EncryptedFileSystemCacheStore` for encrypted-at-rest Node caching)
 */

/**
 * @typedef {Object} CacheEntry
 * @property {unknown} value
 *   The cached payload. Cache stores should be able to persist structured-cloneable
 *   values (objects, arrays, typed arrays like `Uint8Array`, etc).
 * @property {number} createdAtMs
 * @property {number | null} expiresAtMs
 */

/**
 * @typedef {Object} CacheStore
 * @property {(key: string) => Promise<CacheEntry | null>} get
 * @property {(key: string, entry: CacheEntry) => Promise<void>} set
 * @property {(key: string) => Promise<void>} delete
 * @property {(() => Promise<void>) | undefined} [clear]
 * @property {((nowMs?: number) => Promise<void>) | undefined} [pruneExpired]
 */

/**
 * @typedef {{
 *   store: CacheStore;
 *   now?: () => number;
 * }} CacheManagerOptions
 */

export class CacheManager {
  /**
   * @param {CacheManagerOptions} options
   */
  constructor(options) {
    this.store = options.store;
    this.now = options.now ?? (() => Date.now());
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async getEntry(key) {
    const entry = await this.store.get(key);
    if (!entry) return null;
    if (entry.expiresAtMs != null && entry.expiresAtMs <= this.now()) {
      // Cache eviction is best-effort; treat delete failures as a miss.
      try {
        await this.store.delete(key);
      } catch {
        // ignore
      }
      return null;
    }
    return entry;
  }

  /**
   * @param {string} key
   * @returns {Promise<unknown | null>}
   */
  async get(key) {
    const entry = await this.getEntry(key);
    return entry ? entry.value : null;
  }

  /**
   * @param {string} key
   * @param {unknown} value
   * @param {{ ttlMs?: number }} [options]
   */
  async set(key, value, options = {}) {
    const createdAtMs = this.now();
    const expiresAtMs = options.ttlMs != null ? createdAtMs + options.ttlMs : null;
    await this.store.set(key, { value, createdAtMs, expiresAtMs });
  }

  /**
   * Manual invalidation.
   * @param {string} key
   */
  async delete(key) {
    await this.store.delete(key);
  }

  async clear() {
    if (this.store.clear) await this.store.clear();
  }

  /**
   * Best-effort proactive pruning for stores that support it (e.g. IndexedDB).
   *
   * @param {number} [nowMs]
   */
  async pruneExpired(nowMs = this.now()) {
    if (!this.store.pruneExpired) return;
    try {
      await this.store.pruneExpired(nowMs);
    } catch {
      // Cache pruning should never be fatal; callers may invoke this opportunistically.
    }
  }
}
