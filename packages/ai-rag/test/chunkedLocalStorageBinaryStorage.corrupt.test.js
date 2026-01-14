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

class NoRemoveItemLocalStorage {
  constructor() {
    /** @type {Map<string, string>} */
    this._data = new Map();
  }

  getItem(key) {
    return this._data.has(key) ? this._data.get(key) : null;
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

test("ChunkedLocalStorageBinaryStorage clears orphaned chunks when meta is missing", async () => {
  const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const ls = new MemoryLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { value: ls, configurable: true });
  try {
    const storage = new ChunkedLocalStorageBinaryStorage({
      namespace: "formula.test.rag",
      workbookId: "orphan-chunks",
      chunkSizeChars: 8,
    });

    // Simulate a partial write where chunk 0 was written but :meta wasn't.
    ls.setItem(`${storage.key}:0`, "orphaned");
    assert.equal(ls.getItem(`${storage.key}:meta`), null);

    const loaded = await storage.load();
    assert.equal(loaded, null);
    assert.deepEqual(listKeysWithPrefix(ls, `${storage.key}:`), []);
  } finally {
    if (originalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
    } else {
      // eslint-disable-next-line no-undef
      delete globalThis.localStorage;
    }
  }
});

test("ChunkedLocalStorageBinaryStorage clears corrupted meta even when localStorage lacks removeItem", async () => {
  const originalLocalStorage = Object.getOwnPropertyDescriptor(globalThis, "localStorage");
  const ls = new NoRemoveItemLocalStorage();
  Object.defineProperty(globalThis, "localStorage", { value: ls, configurable: true });
  try {
    const storage = new ChunkedLocalStorageBinaryStorage({
      namespace: "formula.test.rag",
      workbookId: "corrupt-meta-no-remove",
      chunkSizeChars: 8,
    });

    ls.setItem(`${storage.key}:meta`, "{not json");
    ls.setItem(`${storage.key}:0`, "orphaned");

    const loaded = await storage.load();
    assert.equal(loaded, null);

    // When `removeItem` is unavailable, we should still self-heal by overwriting
    // with a falsy value so future loads treat the payload as missing.
    assert.equal(ls.getItem(`${storage.key}:meta`), "");
    assert.equal(ls.getItem(`${storage.key}:0`), "");

    const loaded2 = await storage.load();
    assert.equal(loaded2, null);
  } finally {
    if (originalLocalStorage) {
      Object.defineProperty(globalThis, "localStorage", originalLocalStorage);
    } else {
      // eslint-disable-next-line no-undef
      delete globalThis.localStorage;
    }
  }
});
