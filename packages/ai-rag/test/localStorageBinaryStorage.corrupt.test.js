import assert from "node:assert/strict";
import { Buffer as NodeBuffer } from "node:buffer";
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

async function withBrowserBase64(fn) {
  const originalBuffer = Object.getOwnPropertyDescriptor(globalThis, "Buffer");
  const originalAtob = Object.getOwnPropertyDescriptor(globalThis, "atob");
  const originalBtoa = Object.getOwnPropertyDescriptor(globalThis, "btoa");

  // Ensure we exercise the browser path in toBase64/fromBase64 by hiding global Buffer.
  if (originalBuffer) {
    if ("value" in originalBuffer) {
      // Data descriptor; keep original attributes.
      Object.defineProperty(globalThis, "Buffer", { ...originalBuffer, value: undefined });
    } else if (originalBuffer.configurable) {
      // Accessor descriptor (Node 20+); redefine as a data property.
      Object.defineProperty(globalThis, "Buffer", { value: undefined, configurable: true });
    } else if (typeof originalBuffer.set === "function") {
      // Non-configurable accessor: fall back to assignment via setter.
      // eslint-disable-next-line no-undef
      globalThis.Buffer = undefined;
    }
  }

  // Some Node versions don't expose atob/btoa; polyfill them for the duration of the test.
  if (!originalAtob) {
    Object.defineProperty(globalThis, "atob", {
      value: (encoded) => NodeBuffer.from(encoded, "base64").toString("binary"),
      configurable: true,
    });
  }
  if (!originalBtoa) {
    Object.defineProperty(globalThis, "btoa", {
      value: (binary) => NodeBuffer.from(binary, "binary").toString("base64"),
      configurable: true,
    });
  }

  try {
    await fn();
  } finally {
    if (originalBuffer) Object.defineProperty(globalThis, "Buffer", originalBuffer);
    else delete globalThis.Buffer;
    if (originalAtob) Object.defineProperty(globalThis, "atob", originalAtob);
    else delete globalThis.atob;
    if (originalBtoa) Object.defineProperty(globalThis, "btoa", originalBtoa);
    else delete globalThis.btoa;
  }
}

test("LocalStorageBinaryStorage clears invalid base64 payloads on load (browser path)", async () => {
  await withBrowserBase64(async () => {
    await withTempGlobalProp("localStorage", new MemoryLocalStorage(), async () => {
      const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "bad-b64" });
      globalThis.localStorage.setItem(storage.key, "%%%not-base64%%%");
      const loaded = await storage.load();
      assert.equal(loaded, null);
      assert.equal(globalThis.localStorage.getItem(storage.key), null, "expected invalid base64 key to be cleared");
    });
  });
});

test("LocalStorageBinaryStorage clears invalid base64 payloads on load (Buffer path)", async () => {
  await withTempGlobalProp("localStorage", new MemoryLocalStorage(), async () => {
    const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "bad-b64-buffer" });
    globalThis.localStorage.setItem(storage.key, "%%%not-base64%%%");
    const loaded = await storage.load();
    assert.equal(loaded, null);
    assert.equal(globalThis.localStorage.getItem(storage.key), null, "expected invalid base64 key to be cleared");
  });
});

test("LocalStorageBinaryStorage clears invalid base64 even when localStorage lacks removeItem", async () => {
  await withTempGlobalProp("localStorage", new NoRemoveItemLocalStorage(), async () => {
    const storage = new LocalStorageBinaryStorage({ namespace: "formula.test.rag", workbookId: "bad-b64-no-remove" });
    globalThis.localStorage.setItem(storage.key, "%%%not-base64%%%");
    const loaded = await storage.load();
    assert.equal(loaded, null);
    // Key should be overwritten with a falsy value so future loads treat it as missing.
    assert.equal(globalThis.localStorage.getItem(storage.key), "");

    const loaded2 = await storage.load();
    assert.equal(loaded2, null);
  });
});

test("ChunkedLocalStorageBinaryStorage clears invalid legacy base64 payloads on load (browser path)", async () => {
  await withBrowserBase64(async () => {
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
  });
});

test("ChunkedLocalStorageBinaryStorage clears invalid chunk base64 payloads on load (browser path)", async () => {
  await withBrowserBase64(async () => {
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
  });
});
