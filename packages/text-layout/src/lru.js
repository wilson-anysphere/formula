export class LRUCache {
  /**
   * @param {number} maxEntries
   */
  constructor(maxEntries) {
    this.maxEntries = maxEntries;
    /** @type {Map<string, any>} */
    this.map = new Map();
  }

  /**
   * @template T
   * @param {string} key
   * @returns {T | undefined}
   */
  get(key) {
    if (!this.map.has(key)) return undefined;
    const value = this.map.get(key);
    // Refresh key.
    this.map.delete(key);
    this.map.set(key, value);
    return value;
  }

  /**
   * @param {string} key
   * @param {any} value
   */
  set(key, value) {
    if (this.map.has(key)) this.map.delete(key);
    this.map.set(key, value);
    if (this.map.size > this.maxEntries) {
      const oldestKey = this.map.keys().next().value;
      this.map.delete(oldestKey);
    }
  }

  clear() {
    this.map.clear();
  }
}

