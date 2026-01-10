export class LRUCache {
  /**
   * @param {number} maxEntries
   */
  constructor(maxEntries = 200) {
    if (!Number.isFinite(maxEntries) || maxEntries <= 0) {
      throw new Error(`LRUCache maxEntries must be a positive number, got ${maxEntries}`);
    }
    this.maxEntries = maxEntries;
    /** @type {Map<string, any>} */
    this.map = new Map();
  }

  /**
   * @param {string} key
   */
  has(key) {
    return this.map.has(key);
  }

  /**
   * @param {string} key
   */
  get(key) {
    if (!this.map.has(key)) return undefined;
    const value = this.map.get(key);
    // refresh LRU order
    this.map.delete(key);
    this.map.set(key, value);
    return value;
  }

  /**
   * @param {string} key
   * @param {any} value
   */
  set(key, value) {
    if (this.map.has(key)) {
      this.map.delete(key);
    }
    this.map.set(key, value);
    this.#evictIfNeeded();
  }

  /**
   * @param {string} key
   */
  delete(key) {
    return this.map.delete(key);
  }

  clear() {
    this.map.clear();
  }

  get size() {
    return this.map.size;
  }

  #evictIfNeeded() {
    while (this.map.size > this.maxEntries) {
      const oldestKey = this.map.keys().next().value;
      if (oldestKey === undefined) return;
      this.map.delete(oldestKey);
    }
  }
}
