import assert from "node:assert/strict";
import test from "node:test";

import { ChunkedLocalStorageBinaryStorage } from "../src/store/binaryStorage.js";

class MemoryLocalStorage {
  constructor() {
    /** @type {Map<string, string>} */
    this._data = new Map();
  }

  get length() {
    return this._data.size;
  }

  clear() {
    this._data.clear();
  }

  getItem(key) {
    return this._data.get(key) ?? null;
  }

  key(index) {
    return Array.from(this._data.keys())[index] ?? null;
  }

  removeItem(key) {
    this._data.delete(key);
  }

  setItem(key, value) {
    this._data.set(key, value);
  }
}

/**
 * @param {Storage} storage
 * @param {string} prefix
 */
function listKeysWithPrefix(storage, prefix) {
  /** @type {string[]} */
  const out = [];
  for (let i = 0; i < storage.length; i += 1) {
    const key = storage.key(i);
    if (key && key.startsWith(prefix)) out.push(key);
  }
  out.sort();
  return out;
}

test("ChunkedLocalStorageBinaryStorage clears corrupted meta json on load", async () => {
  const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const ls = new MemoryLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { value: ls, configurable: true });
  try {
    const storage = new ChunkedLocalStorageBinaryStorage({
      namespace: "formula.test.rag",
      workbookId: "corrupt-meta",
      chunkSizeChars: 8,
    });

    ls.setItem(`${storage.key}:meta`, "{not json");
    ls.setItem(`${storage.key}:0`, "ignored");

    const loaded = await storage.load();
    assert.equal(loaded, null);

    assert.deepEqual(listKeysWithPrefix(ls, `${storage.key}:`), []);
    assert.equal(ls.getItem(storage.key), null);
  } finally {
    if (originalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
    } else {
      // If the environment had no `localStorage`, remove the property we added so
      // subsequent tests see the original global shape.
      // eslint-disable-next-line no-undef
      delete globalThis.localStorage;
    }
  }
});

test("ChunkedLocalStorageBinaryStorage clears missing chunk keys on load", async () => {
  const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const ls = new MemoryLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { value: ls, configurable: true });
  try {
    const storage = new ChunkedLocalStorageBinaryStorage({
      namespace: "formula.test.rag",
      workbookId: "missing-chunk",
      chunkSizeChars: 8,
    });

    ls.setItem(`${storage.key}:meta`, JSON.stringify({ chunks: 2 }));
    ls.setItem(`${storage.key}:0`, "partial");
    // Missing `${storage.key}:1`

    const loaded = await storage.load();
    assert.equal(loaded, null);

    assert.deepEqual(listKeysWithPrefix(ls, `${storage.key}:`), []);
    assert.equal(ls.getItem(storage.key), null);
  } finally {
    if (originalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
    } else {
      // eslint-disable-next-line no-undef
      delete globalThis.localStorage;
    }
  }
});
