/**
 * Vitest runs in Node, and newer Node versions can expose an experimental
 * `localStorage` implementation that throws unless configured (e.g.
 * `--localstorage-file`). That breaks tests that intentionally run in a DOM
 * environment.
 *
 * These helpers install a simple in-memory `localStorage` shim when the global
 * one is missing or unusable.
 */

export function ensureTestLocalStorage(): Storage {
  try {
    // Accessing `localStorage` can throw in some Node environments.
    // eslint-disable-next-line no-undef
    const storage = localStorage as Storage;
    storage.setItem("__vitest__", "1");
    storage.removeItem("__vitest__");
    return storage;
  } catch {
    const data = new Map<string, string>();

    const storage: Storage = {
      get length() {
        return data.size;
      },
      clear() {
        data.clear();
      },
      getItem(key: string) {
        return data.has(key) ? (data.get(key) ?? null) : null;
      },
      key(index: number) {
        return Array.from(data.keys())[index] ?? null;
      },
      removeItem(key: string) {
        data.delete(key);
      },
      setItem(key: string, value: string) {
        data.set(key, String(value));
      },
    };

    Object.defineProperty(globalThis, "localStorage", {
      value: storage,
      configurable: true,
      enumerable: true,
      writable: true,
    });

    return storage;
  }
}

