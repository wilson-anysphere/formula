export function createMemoryStorage() {
  /** @type {Record<string, string>} */
  const data = Object.create(null);

  return {
    getItem(key) {
      return Object.prototype.hasOwnProperty.call(data, key) ? data[key] : null;
    },
    setItem(key, value) {
      data[key] = String(value);
    },
    removeItem(key) {
      delete data[key];
    }
  };
}

export function getDefaultStorage() {
  // `localStorage` is available in the desktop webview (Tauri) and in browsers.
  // Tests and non-DOM environments should inject a storage implementation.
  try {
    if (typeof globalThis !== "undefined" && globalThis.localStorage) {
      return globalThis.localStorage;
    }
  } catch {
    // Some environments throw on access; fall back to memory storage.
  }

  return createMemoryStorage();
}
