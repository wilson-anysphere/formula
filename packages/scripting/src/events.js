export class TypedEventEmitter {
  constructor() {
    /** @type {Map<string, Set<(payload: any) => void>>} */
    this.listeners = new Map();
  }

  /**
   * @param {string} event
   * @param {(payload: any) => void} listener
   * @returns {() => void}
   */
  on(event, listener) {
    let bucket = this.listeners.get(event);
    if (!bucket) {
      bucket = new Set();
      this.listeners.set(event, bucket);
    }
    bucket.add(listener);
    return () => this.off(event, listener);
  }

  /**
   * @param {string} event
   * @param {(payload: any) => void} listener
   */
  off(event, listener) {
    const bucket = this.listeners.get(event);
    if (!bucket) return;
    bucket.delete(listener);
    if (bucket.size === 0) {
      this.listeners.delete(event);
    }
  }

  /**
   * @param {string} event
   * @param {any} payload
   */
  emit(event, payload) {
    const bucket = this.listeners.get(event);
    if (!bucket) return;
    for (const listener of bucket) {
      listener(payload);
    }
  }
}
