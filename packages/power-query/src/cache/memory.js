/**
 * @typedef {import("./cache.js").CacheEntry} CacheEntry
 */

export class MemoryCacheStore {
  constructor() {
    /** @type {Map<string, CacheEntry>} */
    this.map = new Map();
  }

  /**
   * @param {string} key
   * @returns {Promise<CacheEntry | null>}
   */
  async get(key) {
    return this.map.get(key) ?? null;
  }

  /**
   * @param {string} key
   * @param {CacheEntry} entry
   */
  async set(key, entry) {
    this.map.set(key, entry);
  }

  /**
   * @param {string} key
   */
  async delete(key) {
    this.map.delete(key);
  }

  async clear() {
    this.map.clear();
  }
}

