import assert from "node:assert/strict";
import test from "node:test";

import { ChunkedLocalStorageBinaryStorage, LocalStorageBinaryStorage } from "../src/store/binaryStorage.js";

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

function withTempGlobalProp(name, value, fn) {
  const original = Object.getOwnPropertyDescriptor(globalThis, name);
  Object.defineProperty(globalThis, name, { value, configurable: true });
  return Promise.resolve()
    .then(fn)
    .finally(() => {
      if (original) Object.defineProperty(globalThis, name, original);
      else delete globalThis[name];
    });
}

test("LocalStorageBinaryStorage clears invalid base64 payloads on load (browser path)", async () => {
  // Force the atob/btoa codepath in fromBase64/toBase64.
  const originalBuffer = globalThis.Buffer;
  // eslint-disable-next-line no-undef
  globalThis.Buffer = undefined;
  try {
    await withTempGlobalProp("localStorage", new MemoryLocalStorage(), async () => {
      const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "bad-b64" });
      globalThis.localStorage.setItem(storage.key, "%%%not-base64%%%");
      const loaded = await storage.load();
      assert.equal(loaded, null);
      assert.equal(globalThis.localStorage.getItem(storage.key), null, "expected invalid base64 key to be cleared");
    });
  } finally {
    globalThis.Buffer = originalBuffer;
  }
});

test("ChunkedLocalStorageBinaryStorage clears invalid legacy base64 payloads on load (browser path)", async () => {
  const originalBuffer = globalThis.Buffer;
  // eslint-disable-next-line no-undef
  globalThis.Buffer = undefined;
  try {
    await withTempGlobalProp("localStorage", new MemoryLocalStorage(), async () => {
      const legacy = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "legacy-bad" });
      globalThis.localStorage.setItem(legacy.key, "%%%not-base64%%%");

      const chunked = new ChunkedLocalStorageBinaryStorage({
        namespace: "formula.test.rag",
        workbookId: "legacy-bad",
        chunkSizeChars: 8,
      });
      const loaded = await chunked.load();
      assert.equal(loaded, null);
      assert.equal(globalThis.localStorage.getItem(legacy.key), null, "expected invalid legacy key to be cleared");
    });
  } finally {
    globalThis.Buffer = originalBuffer;
  }
});

test("ChunkedLocalStorageBinaryStorage clears invalid chunk base64 payloads on load (browser path)", async () => {
  const originalBuffer = globalThis.Buffer;
  // eslint-disable-next-line no-undef
  globalThis.Buffer = undefined;
  try {
    await withTempGlobalProp("localStorage", new MemoryLocalStorage(), async () => {
      const storage = new ChunkedLocalStorageBinaryStorage({
        namespace: "formula.test.rag",
        workbookId: "chunk-bad",
        chunkSizeChars: 8,
      });

      // Pretend we have a chunked payload, but the chunk content is invalid base64.
      globalThis.localStorage.setItem(`${storage.key}:meta`, JSON.stringify({ chunks: 1 }));
      globalThis.localStorage.setItem(`${storage.key}:0`, "%%%not-base64%%%");

      const loaded = await storage.load();
      assert.equal(loaded, null);
      assert.equal(globalThis.localStorage.getItem(`${storage.key}:meta`), null);
      assert.equal(globalThis.localStorage.getItem(`${storage.key}:0`), null);
    });
  } finally {
    globalThis.Buffer = originalBuffer;
  }
});

